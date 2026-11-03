"""
End-to-end tests of the template-creation workflow (no network needed).
Run:  python -m tests.test_template_workflow
"""
from __future__ import annotations

import json
import re

from pptx import Presentation

import config
from core import alias_engine as eng
from core.alias_schema import AliasDef, AliasDocument
from core.fidelity import compare
from core.markers import classify, detect_markers, resolve_color, theme_colors
from core.metadata_extractor import extract_metadata
from core.template_catalog import TemplateCatalog
from integrations.stories import LocalStorySource
from template_creation.mapper import LlmMapper
from template_creation.runner import TemplateCreationWorkflow
from template_creation.services import Services
from tests.fakes import FakeOpenAI, temp_repos

ROOT, MOCK = config.ROOT, config.MOCK_DIR
OK = "\u2713"
TOK = re.compile(r"\{\{[A-Z]+\}\}")


def check(cond, msg):
    assert cond, msg
    print(f"  {OK} {msg}")


def workflow(repos, source, steps):
    client = FakeOpenAI(steps)
    return TemplateCreationWorkflow(Services(repos, source, LlmMapper(client=client), renderer=None)), client


def shape(prs, sid):
    return next(s for s in prs.slides[0].shapes if s.shape_id == sid)


def alias_of(doc: dict, path: str) -> str:
    return next(n for n, a in doc["aliases"].items() if a["path"] == path)


def register_bond1(repos):
    repos.datasources.register({"id": "ds_american_bond_1", "name": "American bond (structured)",
                                "type": "json_file", "location": "file://mock/American_bond_1.json", "params": []})


PAGE_NO = [("add_alias", {"shape_id": 500, "paragraph": 0, "text": "5", "path": "$.slide_metadata.page_number"})]
BOND_STEPS = [[("get_alias_proposal", {}), ("get_data_source_response", {})], PAGE_NO, [("submit", {})], "done"]


def test_bond_create_unchanged_update():
    print("\n[1] American bond + American_bond_1.json: aliases, labels, round trip, versions")
    repos = temp_repos("bond"); register_bond1(repos)
    src = LocalStorySource(MOCK / "platinum_american_bond.pptx", "Bond", "American_bond_1", "ds_american_bond_1")
    wf, client = workflow(repos, src, BOND_STEPS)
    r = wf.start("STORY-1")
    check(r.status == "awaiting_review", f"paused for review (status={r.status})")
    doc = r.review["alias_doc"]
    check(len(doc["occurrences"]) >= 60, f"{len(doc['occurrences'])} placeholders, {len(doc['aliases'])} aliases")
    check(any(o["text"] == "5" and doc["aliases"][o["alias"]]["path"] == "$.slide_metadata.page_number"
              for o in doc["occurrences"]), "agent added the page number code left unresolved")
    r = wf.resume(r.run_id, {"action": "approve"})
    check(r.status == "created" and r.state["fidelity"]["fails"] == 0,
          f"created v1, fidelity {r.state['fidelity']['status']} (0 formatting failures)")

    h = TemplateCatalog(repos).find("bond", "american_bond_1")
    t, doc = h.presentation(), h.aliases.model_dump()
    check(shape(t, 3).text_frame.text == "{{" + alias_of(doc, "$.slide_title") + "}}", "title -> placeholder for $.slide_title")
    check(shape(t, 2).text_frame.text == "{{" + alias_of(doc, "$.composite_name") + "}}", "subtitle -> $.composite_name")
    fee = shape(t, 31).text_frame.text
    check(len(TOK.findall(fee)) == 3 and "Gross" not in fee, f"fee box: three placeholders in one paragraph ({fee})")
    tbl = shape(t, 600).table
    check(re.fullmatch(r"Data for \{\{[A-Z]+\}\} \(%\)", tbl.cell(0, 0).text) is not None,
          f"partial match keeps the rest literal: {tbl.cell(0, 0).text!r}")
    check(tbl.cell(1, 1).text == "Cumulative total returns" and tbl.cell(4, 0).text.strip() == "— gross of fees",
          "key-name labels stay static ('Cumulative total returns', '— gross of fees')")
    check(TOK.fullmatch(tbl.cell(2, 1).text.strip()) and TOK.fullmatch(tbl.cell(5, 5).text.strip()),
          "period headers and numbers are placeholders")
    check(re.fullmatch(r"Lifetime: \{\{[A-Z]+\}\}", tbl.cell(7, 0).text.strip()) is not None,
          f"'Lifetime' (a key name) stays literal: {tbl.cell(7, 0).text!r}")
    check(shape(t, 500).text_frame.text.count("{{") == 4, f"data-backup line: {shape(t, 500).text_frame.text}")
    a110 = doc["aliases"][TOK.fullmatch(tbl.cell(5, 5).text.strip()).group(0)[2:-2]]
    check(a110["format"] == "{value:.2f}", "1.10 keeps two decimals ({value:.2f})")
    dash = doc["aliases"][TOK.fullmatch(tbl.cell(4, 6).text.strip()).group(0)[2:-2]]
    check(dash["null"] == "—" and "5_years" in dash["path"], "null 5-year value shows '—'")

    prev = Presentation(str(ROOT / h.location / "preview_filled.pptx"))
    plat = Presentation(str(MOCK / "platinum_american_bond.pptx"))
    diffs = [(s.name, r_, c_) for s in plat.slides[0].shapes if getattr(s, "has_table", False) and s.has_table
             for r_ in range(len(s.table.rows)) for c_ in range(len(s.table.columns))
             if eng.normws(s.table.cell(r_, c_).text) != eng.normws(shape(prev, s.shape_id).table.cell(r_, c_).text)]
    texts = [s.shape_id for s in plat.slides[0].shapes if getattr(s, "has_text_frame", False) and s.has_text_frame
             and eng.normws(s.text_frame.text) != eng.normws(shape(prev, s.shape_id).text_frame.text)]
    check(not diffs and not texts, "preview filled from the JSON reproduces the Platinum copy exactly")

    wf2, _ = workflow(repos, src, [])
    check(wf2.start("STORY-2").status == "unchanged", "same Platinum copy again -> unchanged")

    changed = ROOT / "tests" / "_out" / "bond" / "platinum_v2.pptx"
    p2 = Presentation(str(MOCK / "platinum_american_bond.pptx"))
    shape(p2, 3).text_frame.paragraphs[0].runs[0].text = "Investment results (revised)"
    p2.save(str(changed))
    wf3, _ = workflow(repos, LocalStorySource(changed, "BOND", "american bond 1"), BOND_STEPS)
    r3 = wf3.start("STORY-3")
    d3 = r3.review["alias_doc"]
    check(r3.state["mode"] == "update" and alias_of(d3, "$.slide_title") == alias_of(doc, "$.slide_title"),
          "update: alias names reused from v1 for the same data")
    check(r3.review["diff"]["removed"] == [], "update diff shown to the reviewer (no data fields lost)")
    r3 = wf3.resume(r3.run_id, {"action": "approve"})
    t3 = TemplateCatalog(repos).find("Bond", "American_bond_1").presentation()
    check(r3.status == "updated" and re.fullmatch(r"\{\{[A-Z]+\}\} \(revised\)", shape(t3, 3).text_frame.text) is not None,
          f"v2 saved; new title text partially aliased: {shape(t3, 3).text_frame.text!r}")


def test_markers():
    print("\n[2] Annotated Platinum copy: callouts, red bold locks, design elements")
    repos = temp_repos("annotated")
    repos.datasources.register({"id": "ds_annotated", "name": "Annotated", "type": "json_file",
                                "location": "file://mock/annotated.json", "params": []})
    src = LocalStorySource(MOCK / "platinum_annotated.pptx", "Equity", "ABCF", "ds_annotated")
    ids = {s.name: s.shape_id for s in Presentation(str(MOCK / "platinum_annotated.pptx")).slides[0].shapes}
    decisions = [{"shape_id": ids["KPI 5"], "decision": "design_element"},
                 {"shape_id": ids["Note 6"], "decision": "instruction", "instruction_text": "date from as_of_date",
                  "target_shape_id": ids["Date 2"], "how_applied": "date run is {{alias}} for $.as_of_date"},
                 {"shape_id": ids["Note 8"], "decision": "instruction", "instruction_text": "fund name from data",
                  "target_shape_id": ids["Title 1"], "how_applied": "fund name part is {{alias}}"}]
    steps = [[("get_alias_proposal", {}), ("get_markers", {})], [("submit", {})],
             [("keep_static", {"occurrence_id": "o1", "reason": ""})],
             [("decide_callouts", {"decisions": decisions})], [("submit", {})], "done"]
    wf, client = workflow(repos, src, steps)
    r = wf.start("STORY-9")
    errs = next(x["errors"] for x in client.results if isinstance(x, dict) and x.get("status") == "invalid")
    check(any("has no decision" in e for e in errs), "validator: undecided callouts rejected")
    check(any("reason is required" in str(x.get("error", "")) for x in client.results if isinstance(x, dict)),
          "keep_static without a reason is refused")
    check(any(x.get("design_element_placeholders_added") == 1 for x in client.results if isinstance(x, dict)),
          "design element (green KPI box) got its placeholder from code")
    check(r.status == "awaiting_review", "reached review")
    r = wf.resume(r.run_id, {"action": "approve"})
    check(r.status == "created", f"created (fidelity {r.state['fidelity']['status']})")

    h = TemplateCatalog(repos).find("Equity", "ABCF")
    t = h.presentation()
    names = {s.name for s in t.slides[0].shapes}
    check(not {"Note 6", "Note 8", "Arrow 7"} & names, "instruction callouts and their arrow removed")
    runs = shape(t, ids["Date 2"]).text_frame.paragraphs[0].runs
    th = theme_colors(t)
    check(runs[0].text == "As of " and runs[0].font.bold and classify(resolve_color(runs[0].font.color, th)) == "red"
          and TOK.fullmatch(runs[1].text), f"'As of ' stays red bold, date is a placeholder: {[r.text for r in runs]}")
    check(shape(t, ids["Footer 3"]).text_frame.text.startswith("Source: Capital Group"), "fully locked footer untouched")
    check(shape(t, ids["Returns 4"]).table.cell(2, 0).text == "Benchmark", "locked table cell untouched")
    check(shape(t, ids["Returns 4"]).table.cell(1, 0).text == "Fund", "'Fund' (a key name) stays a static label")

    data = json.loads((MOCK / "annotated.json").read_text(encoding="utf-8"))
    data.update(fund_name="XYZ Fund", as_of_date="30 Sep 2026", kpi_text="+9.1% YTD")
    data["returns"]["fund"] = {"3M": 1.6, "1Y": 9.1}
    warnings = eng.fill(t, h.aliases, data)
    check(not warnings, "Agent 2 fill: no warnings")
    check(shape(t, ids["Title 1"]).text_frame.text == "Portfolio Overview — XYZ Fund", "fill: title updated, rest literal")
    check(shape(t, ids["Date 2"]).text_frame.text == "As of 30 Sep 2026", "fill: locked 'As of ' kept, date filled")
    check(shape(t, ids["KPI 5"]).text_frame.text == "+9.1% YTD", "fill: design element filled")
    check(shape(t, ids["Returns 4"]).table.cell(1, 1).text == "1.6", "fill: table value filled with slide format")
    mk = json.loads((ROOT / h.location / "markers.json").read_text(encoding="utf-8"))
    rep = compare(json.loads((ROOT / h.location / "metadata.json").read_text(encoding="utf-8")),
                  extract_metadata(t), removed_ids={ids["Note 6"], ids["Note 8"], ids["Arrow 7"]},
                  locked_runs=mk["locked_runs"], mode="exhibit")
    check(rep["fails"] == 0, f"filled exhibit keeps the Platinum formatting (fidelity {rep['status']})")


def test_alias_validation():
    print("\n[3] Validation catches the alias.json mistakes found in the hand-made sample")
    prs = Presentation(str(MOCK / "platinum_american_bond.pptx"))
    data = json.loads((MOCK / "American_bond_1.json").read_text(encoding="utf-8"))
    mk = detect_markers(prs)
    base = eng.propose(prs, data, mk).doc

    def errors_after(change):
        doc = AliasDocument.model_validate(base.model_dump())
        change(doc)
        return eng.validate(doc, prs, data, mk, set(), set())

    occ = next(o for o in base.occurrences if o.text == "3 months")
    e, _ = errors_after(lambda d: eng.set_definition(d, occ.id, AliasDef(
        path="$.investment_performance.cumulative_total_returns.Periods[0]")))
    check(any("finds nothing" in x and "case-sensitive" in x for x in e), "'Periods[0]' (wrong case) -> finds nothing")
    num = next(o for o in base.occurrences if o.location.cell == [4, 1])
    e, _ = errors_after(lambda d: eng.set_definition(d, num.id, AliasDef(
        path="$.investment_performance.cumulative_total_returns.returns[0].gross_of_fees.3_months")))
    check(any("invalid JSONPath" in x for x in e), "'gross_of_fees.3_months' -> needs ['3_months']")
    e, _ = errors_after(lambda d: d.occurrences.remove(next(o for o in d.occurrences if o.text == "Investment results")))
    check(any("'Investment results'" in x and "no placeholder" in x for x in e), "dropping a data-backed placeholder -> error")
    e, w = errors_after(lambda d: eng.keep_static(d, next(o for o in d.occurrences if o.text == "Investment results").id,
                                                  "Title fixed by Compliance"))
    check(not e and any("Compliance" in x for x in w), "keep_static with a reason -> allowed, shown as a warning")
    e, _ = errors_after(lambda d: d.aliases.update(ZZ=AliasDef(path="$.slide_title")))
    check(any("ZZ is not used" in x for x in e), "alias defined but not placed -> error")


def test_jira_and_excel():
    print("\n[4] JIRA module (fake server) and Excel datasource")
    from integrations.jira.client import JiraClient
    from integrations.jira.source import JiraStorySource

    class FakeJira(JiraClient):
        def __init__(self): pass
        def get_issue(self, key, fields="*all"):
            return {"key": key, "fields": {"summary": "Investment results slide",
                    "description": "Please build this.\nAsset Type: Fixed Income\n**Acronym**: IMNASMA\nDatasource: ds_american_bond",
                    "attachment": [{"filename": "old.pptx", "created": "2026-01-01", "content": "u1"},
                                   {"filename": "Platinum_v2.pptx", "created": "2026-02-01", "content": "u2"},
                                   {"filename": "notes.pdf", "created": "2026-02-02", "content": "u3"}]}}
        def download(self, url):
            return (MOCK / "platinum_american_bond.pptx").read_bytes()

    jira = JiraStorySource(FakeJira(), field_ids={})
    st = jira.get_story("PB-42")
    check((st.asset_type, st.acronym, st.datasource_id) == ("Fixed Income", "IMNASMA", "ds_american_bond"),
          "story parser read asset type / acronym / datasource from description")
    check([a.filename for a in st.pptx_attachments()] == ["old.pptx", "Platinum_v2.pptx"], "only .pptx attachments")
    repos = temp_repos("jira")
    wf, _ = workflow(repos, jira, [])
    r = wf.start("PB-42")
    check("Platinum_v2.pptx" in json.dumps(r.state["story"]), "newest 'platinum' attachment chosen")
    check(r.status == "failed" and "after 3 attempt" in r.error, "LLM never submitted -> retries exhausted -> 'failed'")

    from openpyxl import Workbook
    from integrations.datasources import build_datasource
    wb = Workbook(); ws = wb.active; ws.title = "Returns"
    for row in [["Period", "Fund", "Index"], ["1Y", 4.31, 4.58], ["3Y", 2.62, 2.98]]:
        ws.append(row)
    xl = ROOT / "tests" / "_out" / "jira" / "returns.xlsx"; wb.save(xl)
    data = build_datasource({"id": "x", "type": "excel_file", "location": str(xl), "params": []}).fetch()
    check(data["sheets"]["Returns"]["grid"][1] == ["1Y", 4.31, 4.58], "Excel -> grid")
    check(data["sheets"]["Returns"]["records"][1]["Index"] == 2.98, "Excel -> records")


def test_fidelity_detects_drift():
    print("\n[5] Fidelity checker catches formatting drift")
    from pptx.util import Pt
    prs = Presentation(str(MOCK / "platinum_annotated.pptx"))
    ref = extract_metadata(prs)
    byname = {s.name: s for s in prs.slides[0].shapes}
    byname["Title 1"].text_frame.paragraphs[0].runs[0].font.size = Pt(20)
    byname["Returns 4"].left += 50000
    byname["Date 2"].text_frame.paragraphs[0].runs[0].text = "Changed "
    rep = compare(ref, extract_metadata(prs), locked_runs=[{"shape_id": byname["Date 2"].shape_id,
                                                            "paragraph": 0, "run": 0}])
    paths = {i["path"] for i in rep["issues"]}
    check("text.p0.font" in paths, "font size change detected")
    check("geometry_emu" in paths, "moved table detected")
    check(any(p.startswith("locked_run") for p in paths), "edited locked (red bold) text detected")


if __name__ == "__main__":
    test_bond_create_unchanged_update()
    test_markers()
    test_alias_validation()
    test_jira_and_excel()
    test_fidelity_detects_drift()
    print("\nAll template-creation workflow tests passed.")

"""
End-to-end tests of Agent 2 (exhibit creation). Builds a template with Agent 1 first.
Run:  python -m tests.test_exhibit_workflow
"""
from __future__ import annotations

import json
import re
import shutil

from pptx import Presentation

import config
from core import alias_engine as eng
from core.exhibit_catalog import ExhibitCatalog
from exhibit_creation.runner import ExhibitCreationWorkflow
from exhibit_creation.services import ExhibitServices
from integrations.stories import LocalStorySource
from tests.fakes import temp_repos
from tests.test_template_workflow import BOND_STEPS, MOCK, ROOT, check, register_bond1, shape, workflow


def make_template(repos):
    src = LocalStorySource(MOCK / "platinum_american_bond.pptx", "Bond", "American_bond_1", "ds_american_bond_1")
    wf, _ = workflow(repos, src, BOND_STEPS)
    r = wf.start("STORY-1")
    r = wf.resume(r.run_id, {"action": "approve"})
    assert r.status == "created", r.error


def agent2(repos):
    return ExhibitCreationWorkflow(ExhibitServices(repos, renderer=None))


def test_create_and_store():
    print("\n[1] Create an exhibit from the template + datasource")
    repos = temp_repos("exhibit"); register_bond1(repos); make_template(repos)
    r = agent2(repos).run("Bond", "American_bond_1")
    check(r.status == "created", f"created (status={r.status}, warnings={r.warnings})")
    loc = r.state["location"]
    check(re.fullmatch(r"tests/_out/exhibit/exhibits/[0-9a-f]{16}__bond__american_bond_1/v1", loc) is not None,
          f"stored under <key_hash>__<asset_type>__<acronym>/v1: {loc}")
    files = sorted(p.name for p in (ROOT / loc).iterdir())
    check("BOND_AMERICAN_BOND_1.pptx" in files and {"exhibit.json", "data.json", "fidelity_report.json"} <= set(files),
          f"files: {files}")
    prs = Presentation(str(ROOT / loc / "BOND_AMERICAN_BOND_1.pptx"))
    check(len(prs.slides) == 1 and prs.slides[0].name == "BOND_AMERICAN_BOND_1", "one slide, named asset type + acronym")
    check(prs.core_properties.title == "BOND_AMERICAN_BOND_1", "file title set to the exhibit name")
    check(not eng.tokens_in(prs), "no placeholders left on the slide")
    plat = Presentation(str(MOCK / "platinum_american_bond.pptx"))
    same = all(eng.normws(s.text_frame.text) == eng.normws(shape(prs, s.shape_id).text_frame.text)
               for s in plat.slides[0].shapes if getattr(s, "has_text_frame", False) and s.has_text_frame)
    check(same, "slide text equals the Platinum copy (same data)")
    check(r.state["fidelity"]["status"] in ("pass", "warn") and r.state["fidelity"]["fails"] == 0,
          f"fidelity vs template: {r.state['fidelity']['status']}")
    meta = json.loads((ROOT / loc / "exhibit.json").read_text(encoding="utf-8"))
    check(meta["template"]["version"] == 1 and len(meta["data_sha256"]) == 64, "exhibit.json records template version + data hash")

    h = ExhibitCatalog(repos).find("bond", "AMERICAN-BOND-1")
    check(h is not None and h.version == 1 and h.presentation().slides[0].name == "BOND_AMERICAN_BOND_1",
          "Agent 3 can fetch it by asset type + acronym (any spelling)")
    check([x.key for x in ExhibitCatalog(repos).search(asset_type="Bond")] == ["BOND::AMERICAN_BOND_1"],
          "search by asset type")

    r2 = agent2(repos).run("Bond", "American_bond_1")
    check(r2.status == "unchanged" and r2.state["version"] == 1, "same template + same data -> unchanged, no duplicate")
    r3 = agent2(repos).run("Bond", "American_bond_1", force=True)
    check(r3.status == "updated" and r3.state["version"] == 2, "force -> v2, v1 kept")


def test_new_data():
    print("\n[2] New data produces a new version with the new values")
    repos = temp_repos("exhibit_new"); register_bond1(repos); make_template(repos)
    agent2(repos).run("Bond", "American_bond_1")
    data = json.loads((MOCK / "American_bond_1.json").read_text(encoding="utf-8"))
    data["slide_title"] = "Investment results Q3"
    data["investment_performance"]["reporting_period"] = "Periods ended September 30, 2026"
    data["investment_performance"]["cumulative_total_returns"]["returns"][0]["gross_of_fees"]["3_months"] = 0.4
    data["investment_performance"]["average_annual_total_returns"]["returns"][0]["gross_of_fees"]["5_years"] = 3.1
    new = ROOT / "tests" / "_out" / "exhibit_new" / "q3.json"
    new.write_text(json.dumps(data), encoding="utf-8")
    repos.datasources.register({**repos.datasources.get("ds_american_bond_1"), "location": str(new)})
    r = agent2(repos).run("Bond", "American_bond_1")
    check(r.status == "updated" and r.state["version"] == 2, "data changed -> v2")
    s = Presentation(r.state["exhibit_path"])
    t = shape(s, 600).table
    check(shape(s, 3).text_frame.text == "Investment results Q3", "title from new data")
    check(t.cell(0, 0).text == "Data for periods ended September 30, 2026 (%)", "partial text + case rule applied")
    check(t.cell(4, 1).text.strip() == "0.40", "0.4 printed with the slide's two decimals")
    check(t.cell(4, 6).text.strip() == "3.10", "former null now a number, formatted")
    check(t.cell(5, 6).text.strip() == "—", "still-null value shows '—'")


def test_failures():
    print("\n[3] Clear failures")
    repos = temp_repos("exhibit_fail"); register_bond1(repos)
    r = agent2(repos).run("Bond", "American_bond_1")
    check(r.status == "failed" and "Create it with the template workflow" in r.error, "no template -> clear error")
    make_template(repos)
    data = json.loads((MOCK / "American_bond_1.json").read_text(encoding="utf-8"))
    del data["slide_title"]
    data["fee_disclosure"]["model_annual_fee_percentage"] = "1.5"
    bad = ROOT / "tests" / "_out" / "exhibit_fail" / "bad.json"
    bad.write_text(json.dumps(data), encoding="utf-8")
    repos.datasources.register({**repos.datasources.get("ds_american_bond_1"), "location": str(bad)})
    r = agent2(repos).run("Bond", "American_bond_1")
    check(r.status == "failed" and "$.slide_title" in r.error, "missing field -> fails before building, names the field")
    check(not repos.exhibit_registry.get("BOND::AMERICAN_BOND_1"), "nothing saved on failure")

    data = json.loads((MOCK / "American_bond_1.json").read_text(encoding="utf-8"))
    data["fee_disclosure"]["model_annual_fee_percentage"] = "1.5"        # unused field changes type
    drift = ROOT / "tests" / "_out" / "exhibit_fail" / "drift.json"
    drift.write_text(json.dumps(data), encoding="utf-8")
    repos.datasources.register({**repos.datasources.get("ds_american_bond_1"), "location": str(drift)})
    r = agent2(repos).run("Bond", "American_bond_1")
    check(r.status == "created" and any("model_annual_fee_percentage" in w for w in r.warnings),
          f"API structure change -> built, but warned: {r.warnings[:1]}")


if __name__ == "__main__":
    test_create_and_store()
    test_new_data()
    test_failures()
    shutil.rmtree(config.ROOT / "tests" / "_out", ignore_errors=True)
    print("\nAll exhibit workflow tests passed.")

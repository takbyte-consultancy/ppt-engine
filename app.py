"""
Streamlit UI for the template-creation workflow.   Run:  streamlit run app.py

The UI only talks to TemplateCreationWorkflow (start / resume / result). It never calls nodes,
repository internals or LangGraph directly, so it can later be replaced by an API.
"""
import json
import io
import re

import streamlit as st
from pptx import Presentation

import config
from core.exhibit_catalog import ExhibitCatalog
from core.template_catalog import TemplateCatalog
from exhibit_creation.runner import ExhibitCreationWorkflow
from exhibit_creation.services import default_exhibit_services
from integrations.stories import LocalStorySource
from repositories.base import get_repositories
from template_creation.runner import TemplateCreationWorkflow
from template_creation.services import default_services
from pitchbook_creation.runner import PitchbookCreationWorkflow
from pitchbook_creation.services import default_pitchbook_services

st.set_page_config(page_title="PPT-AI · Workspace", page_icon="◆", layout="wide")
repos = get_repositories()
ss = st.session_state

dark = ss.get("theme_choice", "Light") == "Dark"
canvas, surface, surface_alt, ink, muted, border, accent = (
    ("#111629", "#1b2340", "#222c4d", "#edf0ff", "#a6aec7", "#343d5a", "#a99cff") if dark else
    ("#f4f6fb", "#ffffff", "#f8f9fd", "#17213b", "#66708a", "#e2e7f0", "#5549d5"))
theme_css = f"""
<style>
  :root {{ color-scheme: {'dark' if dark else 'light'}; }}
  .stApp, [data-testid="stAppViewContainer"] {{ background: {canvas} !important; color: {ink} !important; }}
  [data-testid="stHeader"] {{ background: {canvas} !important; }}
  [data-testid="stAppViewContainer"] .main .block-container {{ max-width: 1500px; padding: 1.65rem 2.2rem 3rem; }}
  .stApp [data-testid="stMarkdownContainer"], .stApp [data-testid="stMarkdownContainer"] *,
  .stApp [data-testid="stWidgetLabel"], .stApp [data-testid="stWidgetLabel"] *,
  .stApp [data-testid="stCaptionContainer"], .stApp [data-testid="stCaptionContainer"] *,
  .stApp [data-testid="stRadio"] label, .stApp [data-testid="stCheckbox"] label {{ color: {ink} !important; }}
  .stApp input, .stApp textarea {{ color: {ink} !important; background: {surface} !important; border-color: {border} !important; }}
  .stApp [data-baseweb="select"] > div {{ background: {surface} !important; border-color: {border} !important; color: {ink} !important; }}
  [data-testid="stFileUploader"] section {{ background: {surface_alt}; border: 1px dashed {border}; border-radius: 12px; }}
  [data-testid="stFileUploader"] section button {{ color: {ink}; background: {surface}; border: 1px solid {border}; }}
  [data-testid="stMetric"] {{ background: {surface}; color: {ink}; border: 1px solid {border}; padding: 14px 16px; border-radius: 14px; }}
  [data-testid="stMetricLabel"] {{ color: {muted} !important; }}
  div[data-testid="stVerticalBlockBorderWrapper"] {{ border-color: {border}; border-radius: 14px; }}
  [data-testid="stForm"] {{ background: {surface}; border: 1px solid {border}; border-radius: 14px; padding: 1rem; }}
  [data-testid="stDataFrame"] {{ border: 1px solid {border}; border-radius: 12px; overflow: hidden; }}
  div[data-testid="stAlert"] {{ border-radius: 12px; }}
  [data-testid="stTabs"] [role="tablist"] {{ gap: .35rem; border-bottom: 1px solid {border}; }}
  [data-testid="stTabs"] [data-baseweb="tab-highlight"] {{ background: {accent} !important; }}
  [data-testid="stTabs"] [data-baseweb="tab-border"] {{ background: {border} !important; }}
  [data-testid="stTabs"] [data-baseweb="tab"] {{ color: {muted} !important; }}
  [data-testid="stTabs"] [data-baseweb="tab"][aria-selected="true"] {{ color: {accent} !important; }}
  [data-testid="stTabs"] [data-baseweb="tab"] p {{ color: inherit !important; }}
  [data-testid="stTabs"] button[role="tab"] {{ min-height: 44px; padding: .6rem .9rem; color: {muted}; border-radius: 10px 10px 0 0; }}
  [data-testid="stTabs"] button[role="tab"][aria-selected="true"] {{ color: {accent} !important; background: {surface}; border-bottom: 2px solid {accent}; font-weight: 650; }}
  [data-testid="stExpander"] {{ border: 1px solid {border}; border-radius: 12px; overflow: hidden; }}
  [data-testid="stExpander"] summary, [data-testid="stExpander"] summary * {{ color: {ink} !important; }}
  [data-testid="stExpander"] summary {{ background: {surface}; }}
  [data-testid="stExpander"] summary:hover {{ background: {surface_alt} !important; }}
  [data-testid="stExpander"] summary:hover, [data-testid="stExpander"] summary:hover * {{ color: {ink} !important; }}
  [role="listbox"], [role="option"] {{ color: {ink} !important; background: {surface} !important; }}
  [role="option"]:hover, [role="option"][data-highlighted="true"] {{ color: {ink} !important; background: {surface_alt} !important; }}
  [data-testid="stBaseButton-primary"] {{ background: {accent} !important; border-color: {accent} !important; border-radius: 10px; }}
  [data-testid="stBaseButton-primary"] p {{ color: #ffffff !important; }}
  .review-actions div[data-testid="stHorizontalBlock"] > div:nth-child(1) [data-testid^="stBaseButton"] {{ background: #15803d !important; border-color: #15803d !important; color: #fff !important; }}
  .review-actions div[data-testid="stHorizontalBlock"] > div:nth-child(2) [data-testid^="stBaseButton"] {{ background: #facc15 !important; border-color: #eab308 !important; color: #422006 !important; }}
  .review-actions div[data-testid="stHorizontalBlock"] > div:nth-child(3) [data-testid^="stBaseButton"] {{ background: #dc2626 !important; border-color: #dc2626 !important; color: #fff !important; }}
  .review-actions [data-testid^="stBaseButton"] p {{ color: inherit !important; }}
  .review-actions [data-testid^="stBaseButton"]:hover,
  .review-actions [data-testid^="stBaseButton"]:focus-visible {{ filter: brightness(.9); color: inherit !important; }}
  [data-testid^="stBaseButton"] {{ border-radius: 10px; }}
  [data-testid="stBaseButton-secondary"] {{ background: {surface} !important; border-color: {border} !important; color: {ink} !important; }}
  [data-testid="stBaseButton-secondary"] p, [data-testid="stBaseButton-secondary"] span {{ color: {ink} !important; }}
  [data-testid="stBaseButton-primary"] p, [data-testid="stBaseButton-primary"] span {{ color: #ffffff !important; }}
  .tk-brand {{ font-size: 1.8rem; font-weight: 750; letter-spacing: -.045em; color: {ink}; line-height: 1.1; }}
  .tk-brand span {{ color: {accent}; }}
  .tk-muted {{ color: {muted}; font-size: .92rem; }}
  .tk-kicker {{ color: {accent}; text-transform: uppercase; font-size: .72rem; font-weight: 700; letter-spacing: .12em; }}
  .tk-session {{ display: inline-block; padding: .42rem .7rem; border: 1px solid {border}; border-radius: 999px; background: {surface}; color: {muted}; font-size: .82rem; }}
  .tk-step {{ min-height: 82px; padding: .9rem 1rem; border: 1px solid {border}; border-radius: 14px; background: {surface}; }}
  .tk-step b {{ display: block; margin: .2rem 0; color: {ink}; }}
  .tk-step span {{ color: {muted}; font-size: .84rem; }}
  .stApp input:focus, .stApp textarea:focus {{ outline: 2px solid {accent} !important; outline-offset: 1px; }}
  @media (hover: hover) and (pointer: fine) {{
    [data-testid="stBaseButton-primary"] {{ transition: background-color 180ms ease, opacity 180ms ease; }}
    [data-testid="stBaseButton-primary"]:hover {{ opacity: .9; }}
    [data-testid="stBaseButton-secondary"] {{ transition: background-color 180ms ease, border-color 180ms ease; }}
    [data-testid="stBaseButton-secondary"]:hover {{ background: {surface_alt}; border-color: {accent}; }}
  }}
  @media (prefers-reduced-motion: reduce) {{ *, *::before, *::after {{ scroll-behavior: auto !important; transition-duration: .01ms !important; animation-duration: .01ms !important; }} }}
</style>
"""
st.markdown(theme_css, unsafe_allow_html=True)

brand, product, theme = st.columns([1.25, 3.4, .8], vertical_alignment="center")
with brand:
    st.markdown('<div class="tk-brand">◆ <span>PPT-AI</span></div>', unsafe_allow_html=True)
with product:
    st.markdown('<div class="tk-muted">PowerPoint workflow workspace · templates, exhibits, and pitchbooks</div>', unsafe_allow_html=True)
with theme:
    st.selectbox("Theme", ["Light", "Dark"], key="theme_choice", label_visibility="collapsed")
tab_overview, tab_run, tab_exhibit, tab_pitchbook, tab_registry, tab_ds = st.tabs(
    ["Overview", "Template workflow", "Exhibit creation", "Pitchbook studio", "Template registry", "Datasources"])

template_entries = TemplateCatalog(repos).list()
exhibit_entries = repos.exhibit_registry.list() if repos.exhibit_registry else []
pitchbook_entries = repos.pitchbook_registry.list() if repos.pitchbook_registry else []

with tab_overview:
    st.markdown('<div class="tk-kicker">PPT-AI · workspace</div>', unsafe_allow_html=True)
    st.title("PowerPoint workflow overview")
    st.caption("Move from reusable template, to data-filled exhibit, to reviewed pitchbook.")
    m1, m2, m3 = st.columns(3)
    m1.metric("Template families", len(template_entries), help="Registered asset type and acronym keys.")
    m2.metric("Exhibit families", len(exhibit_entries), help="Keys with at least one saved exhibit version.")
    m3.metric("Pitchbooks", len(pitchbook_entries), help="Saved pitchbook workflows.")
    st.subheader("Three-step workflow")
    a, b, c = st.columns(3)
    a.markdown('<div class="tk-step"><div class="tk-kicker">01 · Agent 1</div><b>Create a template</b><span>Map Platinum slide content to datasource fields. Review aliases before approval.</span></div>', unsafe_allow_html=True)
    b.markdown('<div class="tk-step"><div class="tk-kicker">02 · Agent 2</div><b>Generate exhibits</b><span>Fill saved templates, check fidelity, and keep immutable slide versions.</span></div>', unsafe_allow_html=True)
    c.markdown('<div class="tk-step"><div class="tk-kicker">03 · Agent 3</div><b>Assemble a pitchbook</b><span>Map sample UIDs or select exhibits, reorder, collect feedback, then approve.</span></div>', unsafe_allow_html=True)
    st.info("Use the workflow tabs above. Registry, datasource setup, review tables, run logs, previews, and downloads remain available in their original steps.")


def render_events(container):
    def on_event(kind, payload):
        if kind == "node":
            container.write(f"▸ {payload}")
        elif kind == "tool_call":
            container.caption(f"   🔧 {payload['tool']}")
        elif kind == "tool_result" and isinstance(payload["result"], dict) and payload["result"].get("errors"):
            container.caption(f"   ⚠️ validator: {len(payload['result']['errors'])} issue(s), agent is fixing them")
        elif kind == "error":
            container.error(payload)
    return on_event


# =========================================================================== #
with tab_run:
    st.markdown('<div class="tk-kicker">Agent 1 · Template design</div>', unsafe_allow_html=True)
    st.header("Create or update a template")
    st.caption("Start from a Jira story or a Platinum copy. Review proposed mappings before the template is saved.")
    left, right = st.columns([1, 2])

    with left:
        st.subheader("Template inputs")
        sources = (["JIRA story"] if config.JIRA_BASE_URL else []) + ["Upload Platinum copy"]
        src_kind = st.radio("Source", sources, help="JIRA appears once JIRA_BASE_URL is set in .env")
        story_ref, upload = "", None
        if src_kind == "JIRA story":
            story_ref = st.text_input("Story key", placeholder="PB-123")
        else:
            upload = st.file_uploader("Platinum copy (.pptx)", type=["pptx"])

        st.markdown("**Template key fields**")
        st.caption("Read from the story when available; anything typed here overrides it.")
        asset_type = st.text_input("Asset type")
        acronym = st.text_input("Acronym")
        ds_ids = [d["id"] for d in repos.datasources.list()]
        ds_choice = st.selectbox("Datasource", ["(from story / existing template)"] + ds_ids)

        with st.expander("Options"):
            slide_no = st.number_input("Slide number (multi-slide Platinum copies)", 0, 100, 0,
                                       help="0 = the file must contain exactly one slide")
            force = st.checkbox("Re-process even if the Platinum copy is unchanged")
            vision = st.checkbox("Send a slide image to the LLM (needs LibreOffice)", value=True)

        if st.button("Run workflow", type="primary", width="stretch"):
            overrides = {"asset_type": asset_type, "acronym": acronym,
                         "datasource_id": "" if ds_choice.startswith("(") else ds_choice}
            try:
                if src_kind == "JIRA story":
                    from integrations.jira import JiraStorySource
                    source = JiraStorySource()
                else:
                    if not upload:
                        st.error("Upload a Platinum copy first."); st.stop()
                    path = config.WORK_DIR / "uploads" / upload.name
                    path.parent.mkdir(parents=True, exist_ok=True)
                    path.write_bytes(upload.getvalue())
                    source = LocalStorySource(path, asset_type, acronym, overrides["datasource_id"])
                    story_ref = f"LOCAL-{path.stem}"
                status = st.status("Running…", expanded=True)
                svc = default_services(source, render_events(status), use_vision=vision)
                ss.wf = TemplateCreationWorkflow(svc)
                ss.result = ss.wf.start(story_ref, overrides=overrides,
                                        slide_index=(slide_no - 1) if slide_no else None, force=force)
                status.update(label=f"Status: {ss.result.status}", state="complete")
            except Exception as e:
                st.exception(e)

    with right:
        r = ss.get("result")
        if not r:
            st.info("Run the workflow to see the review and results here.")
        elif r.status == "awaiting_review":
            rv = r.review
            st.subheader(f"Review · {rv['template_key']} · {rv['mode'].upper()}")
            if rv.get("errors"):
                st.error("Your edits have problems:\n\n" + "\n\n".join(rv["errors"]))
            for w in rv.get("warnings", []):
                st.warning(w)
            doc = rv["alias_doc"]
            if rv["mode"] == "update" and rv.get("diff"):
                d = rv["diff"]
                st.markdown(f"**Changes vs current version:** {len(d['added'])} data fields added, "
                            f"{len(d['removed'])} removed, {len(d['changed'])} changed")
                if d["added"] or d["removed"] or d["changed"]:
                    st.json(d, expanded=False)
            uses: dict = {}
            for o in doc["occurrences"]:
                uses.setdefault(o["alias"], []).append(o)
            st.markdown(f"**Placeholders** · {len(doc['occurrences'])} in the template, "
                        f"{len(doc['aliases'])} aliases")
            st.dataframe([{"alias": "{{" + n + "}}", "replaces": uses[n][0]["text"][:60],
                           "used": len(uses.get(n, [])), "JSON path": a["path"],
                           "format": a.get("format") or "", "case": a.get("case") or "", "null": a.get("null", ""),
                           "confidence": min((o["confidence"] for o in uses.get(n, [])),
                                             key=["low", "medium", "high"].index, default="")}
                          for n, a in doc["aliases"].items()], width="stretch")
            if doc.get("labels"):
                with st.expander(f"Static labels ({len(doc['labels'])}): text that names a JSON key, kept as-is"):
                    st.dataframe([{"text": l["text"], "JSON key": l["key_path"]} for l in doc["labels"]],
                                 width="stretch")
            if doc.get("kept_static"):
                st.markdown("**Kept static although data-backed**")
                st.dataframe(doc["kept_static"], width="stretch")
            decisions = {c["shape_id"]: c for c in doc.get("callouts", [])}
            if rv["callouts"]:
                st.markdown("**Yellow / green callouts**")
                st.dataframe([{"colour": c["colour"], "callout text": c["text"],
                               "decision": decisions.get(c["shape_id"], {}).get("decision"),
                               "target id": decisions.get(c["shape_id"], {}).get("target_shape_id"),
                               "how applied": decisions.get(c["shape_id"], {}).get("how_applied")}
                              for c in rv["callouts"]], width="stretch")
            if rv["locked_runs"]:
                st.markdown("**Locked red bold text (kept exactly as-is)**")
                st.dataframe([{"shape": l["original_name"], "where": f"cell {l['cell']}" if "cell" in l else
                               f"p{l['paragraph']}.r{l['run']}", "text": l["text"]} for l in rv["locked_runs"]],
                             width="stretch")
            st.markdown("**alias.json** (edit and use *Approve with my edits*)")
            edited = st.text_area("alias.json", json.dumps(doc, indent=2, ensure_ascii=False), height=380,
                                  label_visibility="collapsed")
            st.markdown('<div class="review-actions">', unsafe_allow_html=True)
            c1, c2, c3 = st.columns(3)
            if c1.button("Approve", type="primary"):
                ss.result = ss.wf.resume(r.run_id, {"action": "approve"}); st.rerun()
            if c2.button("Approve with my edits"):
                try:
                    ss.result = ss.wf.resume(r.run_id, {"action": "approve", "alias_doc": json.loads(edited)})
                    st.rerun()
                except json.JSONDecodeError as e:
                    st.error(f"Not valid JSON: {e}")
            reason = c3.text_input("Reject reason", label_visibility="collapsed", placeholder="Reject reason")
            if c3.button("Reject"):
                ss.result = ss.wf.resume(r.run_id, {"action": "reject", "reason": reason}); st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)
        else:
            s = r.state
            if r.status in ("created", "updated"):
                st.success(f"{s['key']['key']} {r.status}: version {s['version']} at `{s['location']}`")
            elif r.status == "unchanged":
                st.info(f"{s['key']['key']}: Platinum copy unchanged since v{s['version']}. Nothing to do.")
            else:
                st.error(f"{r.status}: {r.error}")
            if s.get("fidelity"):
                fd = s["fidelity"]
                st.markdown(f"**Fidelity check: {fd['status'].upper()}** · {fd['fails']} failures · "
                            f"{fd['warnings']} warnings")
                if fd["issues"]:
                    st.dataframe(fd["issues"], width="stretch")
            if s.get("placeholders"):
                st.caption(f"{s['placeholders']} placeholders written; removed: {s.get('removed_shapes') or 'nothing'}")
            if r.status in ("created", "updated"):
                c1, c2, c3 = st.columns(3)
                c1.download_button("template.pptx", repos.templates.read(s["location"], "template.pptx"),
                                   "template.pptx")
                c2.download_button("alias.json", repos.templates.read(s["location"], "alias.json"), "alias.json")
                c3.download_button("preview (filled with today's data)",
                                   repos.templates.read(s["location"], "preview_filled.pptx"), "preview_filled.pptx")
            with st.expander("Run log"):
                st.write(s.get("log", []))

# =========================================================================== #
with tab_exhibit:
    st.markdown('<div class="tk-kicker">Agent 2 · Deterministic slide fill</div>', unsafe_allow_html=True)
    st.header("Create an exhibit")
    st.write("Fill a saved template with its datasource and save a versioned slide for pitchbook assembly.")
    with st.form("exhibit_creation_form"):
        c1, c2 = st.columns(2)
        exhibit_asset_type = c1.text_input("Asset type", placeholder="Fixed Income")
        exhibit_acronym = c2.text_input("Acronym", placeholder="IMNASMA")
        c3, c4 = st.columns(2)
        template_version_text = c3.text_input("Template version", placeholder="Latest")
        force_exhibit = c4.checkbox("Regenerate even when data is unchanged")
        params_text = st.text_area("Datasource parameters (JSON)", value="{}", height=90,
                                   help="Runtime parameters passed to the template's registered datasource.")
        run_exhibit = st.form_submit_button("Create exhibit", type="primary")
    if run_exhibit:
        ss.pop("exhibit_result", None)
        try:
            params = json.loads(params_text or "{}")
            if not isinstance(params, dict):
                raise ValueError("Datasource parameters must be a JSON object")
            version = int(template_version_text) if template_version_text.strip() else None
            wf = ExhibitCreationWorkflow(default_exhibit_services())
            ss.exhibit_result = wf.run(exhibit_asset_type, exhibit_acronym, params=params,
                                       template_version=version, force=force_exhibit)
        except Exception as e:
            ss.exhibit_error = str(e)
        else:
            ss.pop("exhibit_error", None)
    if ss.get("exhibit_error"):
        st.error(ss.exhibit_error)
    exhibit_result = ss.get("exhibit_result")
    if exhibit_result:
        state = exhibit_result.state
        if exhibit_result.status in ("created", "updated"):
            st.success(f"{state['exhibit_name']} {exhibit_result.status} · v{state['version']}")
            c1, c2 = st.columns([1, 2])
            with c1:
                st.download_button("Download exhibit slide", repos.exhibits.read(
                    state["location"], f"{state['exhibit_name']}.pptx"),
                    file_name=f"{state['exhibit_name']}.pptx", type="primary")
            with c2:
                handle = ExhibitCatalog(repos).find(exhibit_asset_type or state["key"]["fields"].get("asset_type", ""),
                    exhibit_acronym or state["key"]["fields"].get("acronym", ""), state["version"])
                if handle and handle.preview_png():
                    st.image(handle.preview_png(), caption="Saved exhibit preview", width=560)
        elif exhibit_result.status == "unchanged":
            st.info(f"No changes detected · existing version v{state['version']} is current.")
        else:
            st.error(state.get("error", exhibit_result.status))
        if state.get("fidelity"):
            st.caption(f"Fidelity: {state['fidelity']['status']} · {state['fidelity']['fails']} issues")
        if exhibit_result.warnings:
            with st.expander(f"Warnings ({len(exhibit_result.warnings)})"):
                st.write(exhibit_result.warnings)
        with st.expander("Exhibit run log"):
            st.write(state.get("log", []))

# =========================================================================== #
with tab_registry:
    entries = template_entries
    if not entries:
        st.info("No templates yet.")
    for e in entries:
        with st.expander(f"{e['key']} · v{e['current_version']} · {e['datasource_id']}"):
            st.write({"key_hash": e["key_hash"], "fields": e["fields"], "location": e["current_location"],
                      "updated_at": e["updated_at"]})
            st.dataframe(e["versions"], width="stretch")

with tab_ds:
    st.dataframe([{"id": d["id"], "type": d["type"], "location": d["location"], "params": d.get("params"),
                   "used by": ", ".join(repos.datasources.templates_using(d["id"]))}
                  for d in repos.datasources.list()], width="stretch")
    with st.form("register_ds"):
        st.markdown("**Register a datasource**")
        c1, c2 = st.columns(2)
        ds_id = c1.text_input("ID", "ds_")
        ds_type = c2.selectbox("Type", ["json_file", "excel_file", "http"])
        ds_name = c1.text_input("Name")
        ds_loc = c2.text_input("Location", "file://mock/input.json",
                               help="file://… for json/excel, https://… for http; {param} placeholders allowed")
        ds_params = st.text_input("Run-time params (comma separated)")
        if st.form_submit_button("Register"):
            repos.datasources.register({"id": ds_id, "name": ds_name, "type": ds_type, "location": ds_loc,
                                        "params": [p.strip() for p in ds_params.split(",") if p.strip()]})
            st.rerun()

# =========================================================================== #
with tab_pitchbook:
    st.markdown('<div class="tk-kicker">Agent 3 · Book assembly</div>', unsafe_allow_html=True)
    st.header("Create a pitchbook")
    st.write("Assemble saved exhibit slides into a reviewed, versioned PowerPoint.")
    catalog = ExhibitCatalog(repos)
    handles = []
    fields_by_key = {}
    if repos.exhibit_registry:
        for entry in repos.exhibit_registry.list():
            fields = entry.get("fields", {})
            fields_by_key[entry["key"]] = fields
            for version_record in entry.get("versions", []):
                handle = catalog.find(fields.get("asset_type", ""), fields.get("acronym", ""),
                                      version_record["version"])
                if handle:
                    handles.append(handle)
    handles.sort(key=lambda h: (h.key, h.version))
    ref_by_label = {f"{h.key} · v{h.version} · {h.record['file_name']}":
                    {"key": h.key, "version": h.version} for h in handles}
    labels = list(ref_by_label)
    ss.setdefault("pitchbook_feedback", [])

    # Give the 49-slide sample mapper the full canvas until a draft exists; split
    # the page only when there is a review panel to show beside the inputs.
    left, right = st.columns([0.75, 2.25] if ss.get("pitchbook_result") else [1, 0.03], gap="large")
    with left:
        mode = st.radio("Assembly method", ["From a sample deck", "From a request"], horizontal=True)
        title = st.text_input("Pitchbook name", value="PPT-AI Pitchbook", key="pitchbook_title")
        refs, criteria = [], {}
        if mode == "From a sample deck":
            sample = st.file_uploader("Reference pitchbook (.pptx)", type=["pptx"], key="pitchbook_sample")
            sample_profile = {}
            if sample:
                try:
                    sample_prs = Presentation(io.BytesIO(sample.getvalue()))
                    sample_profile = {"slide_count": len(sample_prs.slides),
                        "width_inches": round(sample_prs.slide_width / 914400, 2),
                        "height_inches": round(sample_prs.slide_height / 914400, 2)}
                    st.info(f"Reference example · {sample_profile['slide_count']} slides · "
                            f"{sample_profile['width_inches']} × {sample_profile['height_inches']} in. "
                            "Visible UID markers will be mapped to saved exhibit versions; the example’s existing content is not copied.")
                    for index, slide in enumerate(sample_prs.slides, 1):
                        text = " ".join(shape.text.replace("\n", " ") for shape in slide.shapes
                                        if getattr(shape, "has_text_frame", False) and shape.has_text_frame)
                        markers = list(dict.fromkeys(re.findall(
                            r"(?<![A-Z0-9])(?:[A-Z]{2,}(?:@[A-Z]+)?\.)[A-Z0-9]+(?:[._][A-Z0-9]+)*(?![A-Z0-9])",
                            text.upper())))
                        title_text = next((" ".join(shape.text.split()) for shape in slide.shapes
                                           if getattr(shape, "has_text_frame", False) and shape.has_text_frame and shape.text.strip()), "")
                        with st.expander(f"Slide {index} · {title_text[:72] or 'Untitled'} · {len(markers)} UID(s)"):
                            if markers:
                                for marker_index, marker in enumerate(markers, 1):
                                    st.caption(f"Sample UID · {marker}")
                                    options = ["(skip this UID)"] + labels
                                    possible = [label for label, ref in ref_by_label.items()
                                                if marker.casefold() in f"{ref['key']} {label}".casefold()]
                                    default_index = options.index(possible[0]) if possible else 0
                                    choice = st.selectbox(f"Saved exhibit for {marker}", options, index=default_index,
                                        key=f"sample_map_{index}_{marker_index}")
                                    if choice != "(skip this UID)":
                                        refs.append({**ref_by_label[choice], "sample_uid": marker,
                                                     "sample_slide": index})
                            else:
                                st.caption("No visible UID found. Map this page manually if it represents an exhibit.")
                                choice = st.selectbox(f"Saved exhibit for sample slide {index}", ["(skip this slide)"] + labels,
                                                      key=f"sample_map_{index}")
                                if choice != "(skip this slide)":
                                    refs.append({**ref_by_label[choice], "sample_uid": None, "sample_slide": index})
                except Exception as e:
                    st.error(f"Could not read sample pitchbook: {e}")
            criteria = {"sample_file": sample.name if sample else "", "sample_profile": sample_profile}
        else:
            request = st.text_area("Describe the pitchbook and preferred slide order",
                placeholder="For IMNASMA, create a FY2026 pitchbook. Start with the market overview, then show performance and risks.", height=95,
                help="The saved exhibit/acronym pair below selects the source slide. This request can guide the narrative order; AI ordering is optional.")
            families = {}
            for handle in handles:
                fields = fields_by_key.get(handle.key, {})
                family_key = (str(fields.get("acronym", "")).strip(), str(fields.get("asset_type", "")).strip())
                if all(family_key):
                    families.setdefault(family_key, []).append(handle)
            acronyms = sorted({family[0] for family in families}, key=str.casefold)
            if acronyms:
                acronym_filter = st.selectbox("Acronym · required", acronyms, key="pitchbook_query_acronym")
                asset_types = sorted({asset for acronym, asset in families if acronym == acronym_filter}, key=str.casefold)
                asset_filter = st.multiselect("Exhibits / asset types · select one or more",
                    asset_types, default=asset_types,
                    key=f"pitchbook_query_asset_types_{re.sub(r'[^A-Za-z0-9]+', '_', acronym_filter)}",
                    help="Each selected exhibit + acronym pair resolves to its latest saved slide version.")
                for asset_type_choice in asset_filter:
                    matches = families.get((acronym_filter, asset_type_choice), [])
                    latest = max(matches, key=lambda handle: int(handle.version)) if matches else None
                    if latest:
                        registry_entry = repos.exhibit_registry.get(latest.key) if repos.exhibit_registry else None
                        uid = f"{registry_entry['key_hash']}:v{latest.version}" if registry_entry else f"{latest.key}:v{latest.version}"
                        st.caption(f"{asset_type_choice} · resolved UID {uid} · {latest.record['file_name']}")
                year_filter = st.text_input("Financial year (optional filter)", placeholder="FY2026")
            else:
                acronym_filter, asset_filter, year_filter = "", [], ""
                st.info("Create and save an exhibit first. Query-based pitchbooks resolve from registered exhibit + acronym pairs.")
            use_ai_order = st.checkbox("Use AI to interpret the requested slide order",
                value=bool(config.OPENAI_API_KEY), disabled=not bool(config.OPENAI_API_KEY),
                help=("AI can reorder only the saved slides resolved from this exact pair. You can review and change the order before approval." if config.OPENAI_API_KEY
                      else "Configure OPENAI_API_KEY to enable AI ordering. The saved slide order and manual reordering remain available."),
                key="pitchbook_use_ai_order")
            if not config.OPENAI_API_KEY:
                st.caption("AI ordering is currently unavailable; the saved exhibit order will be used and can be changed in review.")
            criteria = {"request": request, "acronym": acronym_filter, "financial_year": year_filter,
                        "asset_type": asset_filter, "use_llm_order": use_ai_order}

        can_build = bool(refs) if mode == "From a sample deck" else bool(criteria.get("acronym") and criteria.get("asset_type"))
        if mode == "From a sample deck" and sample and not labels:
            st.warning("No saved exhibits are available yet. Create exhibits first, then map at least one sample UID to an exhibit.")
        if st.button("Build pitchbook draft", type="primary", disabled=False, key="build_pitchbook"):
            if not can_build:
                st.warning("Select a saved exhibit/acronym pair, or map at least one sample UID to a saved exhibit.")
                st.stop()
            try:
                wf = PitchbookCreationWorkflow(default_pitchbook_services())
                ss.pitchbook_wf = wf
                ss.pitchbook_result = wf.start(refs, title=title,
                    mode="sample" if mode == "From a sample deck" else "request", criteria=criteria)
                ss.pitchbook_feedback = []
                ss.pop("pitchbook_order", None)
                ss["pitchbook_final_name"] = title
                st.rerun()
            except Exception as e:
                st.error(f"Pitchbook draft failed: {e}")

    with right:
        result = ss.get("pitchbook_result")
        if result is None:
            pass
        elif result.status == "awaiting_review":
            st.subheader("Review the draft")
            st.caption("Reorder by changing the position number. Source exhibit versions stay immutable.")
            if result.state.get("log"):
                st.info(result.state["log"][-1])
            slides = result.review["slides"]
            order_rows = [{"position": i + 1, "slide": i + 1, "exhibit": x["exhibit_name"],
                           "source": f"{x['key']} · v{x['version']}", "source_file": x["file_name"]}
                          for i, x in enumerate(slides)]
            edited_order = st.data_editor(order_rows, hide_index=True, width="stretch", key="pitchbook_order",
                column_config={"position": st.column_config.NumberColumn("Position", min_value=1,
                    max_value=len(slides), step=1), "slide": st.column_config.NumberColumn(disabled=True),
                    "exhibit": st.column_config.TextColumn(disabled=True), "source": st.column_config.TextColumn(disabled=True),
                    "source_file": st.column_config.TextColumn(disabled=True)})
            for i, x in enumerate(slides):
                handle = catalog.find(x["fields"].get("asset_type", ""), x["fields"].get("acronym", ""), x["version"])
                with st.expander(f"Slide {i + 1} · {x['exhibit_name']} · {x['key']} v{x['version']}"):
                    if handle and handle.preview_png():
                        st.image(handle.preview_png(), width="stretch")
                    component_rows = result.review["components"][i] if i < len(result.review["components"]) else []
                    st.dataframe(component_rows, hide_index=True, width="stretch")

            st.markdown("#### Feedback")
            st.caption("Attach feedback to the whole pitchbook, a slide, or a source component. Each note is retained in the final manifest.")
            run_key = result.run_id
            scope = st.selectbox("Applies to", ["Select scope", "Pitchbook", "Slide", "Component"],
                                 key=f"feedback_scope_{run_key}")
            target_slide = None
            target_component = None
            comp_options = []
            components = []
            replacement_text = ""
            if scope in ("Slide", "Component"):
                target_slide = st.selectbox("Slide", list(range(1, len(slides) + 1)),
                                            key=f"feedback_slide_{run_key}")
            if scope == "Component":
                components = result.review["components"][target_slide - 1]
                comp_options = [f"{c['name']} · shape {c['shape_id']} · {c['text'][:60]}" for c in components]
                if comp_options:
                    target_component = st.selectbox("Component", comp_options,
                                                     key=f"feedback_component_{run_key}")
                    replacement_text = st.text_input("Replacement text (optional; exact text edit)",
                        help="Leave blank to keep this as a comment for review without changing the component.",
                        key=f"feedback_replacement_{run_key}")
                else:
                    st.info("This slide has no editable text components.")
            with st.form(f"pitchbook_feedback_form_{run_key}", clear_on_submit=True):
                note = st.text_area("Requested change", placeholder="Describe the requested edit")
                submitted = st.form_submit_button("Add feedback", disabled=scope == "Select scope")
            if submitted:
                if note.strip():
                    ss.pitchbook_feedback.append({"scope": scope.lower(), "slide_number": target_slide,
                        "component": target_component, "source": (slides[target_slide - 1] if target_slide else None),
                        "shape_id": (components[comp_options.index(target_component)]["shape_id"]
                                     if scope == "Component" and target_component in comp_options else None),
                        "replacement_text": replacement_text.strip() if replacement_text.strip() else None,
                        "comment": note.strip()})
                    st.rerun()
                else:
                    st.warning("Write a feedback note first.")
            if ss.pitchbook_feedback:
                st.dataframe([{**f, "source": f["source"]["key"] + " v" + str(f["source"]["version"])
                               if f.get("source") else "Entire pitchbook"} for f in ss.pitchbook_feedback],
                             hide_index=True, width="stretch")
            filename = st.text_input("Final deck name", value=result.state.get("title", title), key="pitchbook_final_name")
            use_llm_feedback = st.checkbox("Use AI to interpret requested changes",
                value=bool(config.OPENAI_API_KEY), disabled=not bool(config.OPENAI_API_KEY),
                help=("AI turns notes into validated slide, title, or component operations. "
                      "Only those operations are applied; other slide content stays unchanged." if config.OPENAI_API_KEY
                      else "Configure OPENAI_API_KEY to enable AI feedback interpretation."),
                key=f"use_llm_feedback_{result.run_id}")
            if not config.OPENAI_API_KEY:
                st.caption("AI feedback is unavailable until OPENAI_API_KEY is configured. Exact component replacements remain available.")
            if st.button("Apply feedback and review revisions", type="primary", key="approve_pitchbook"):
                try:
                    order_records = (edited_order.to_dict(orient="records")
                                     if hasattr(edited_order, "to_dict") else list(edited_order))
                    positions = [int(row["position"]) for row in order_records]
                    if sorted(positions) != list(range(1, len(slides) + 1)):
                        st.error("Each slide needs a unique position from 1 to the number of slides.")
                    else:
                        permutation = [next(i for i, row in enumerate(order_records) if int(row["position"]) == p)
                                       for p in range(1, len(slides) + 1)]
                        ss.pitchbook_result = ss.pitchbook_wf.resume(result.run_id, {
                            "action": "approve", "order": permutation, "title": filename,
                            "feedback": ss.pitchbook_feedback, "use_llm_feedback": use_llm_feedback})
                        st.rerun()
                except Exception as e:
                    st.error(f"Could not save pitchbook: {e}")
            if st.button("Reject draft", key="reject_pitchbook"):
                ss.pitchbook_result = ss.pitchbook_wf.resume(result.run_id, {"action": "reject"})
                st.rerun()
        elif result.status == "awaiting_feedback_review":
            st.subheader("Review feedback agent changes")
            st.caption("The feedback agent prepared these revisions. Confirm them before the final deck is saved.")
            st.markdown(f"**Final deck name:** {result.review.get('title')}")
            st.dataframe([{"position": i + 1, "exhibit": x["exhibit_name"],
                           "source": f"{x['key']} · v{x['version']}"}
                          for i, x in enumerate(result.review.get("slides", []))],
                         hide_index=True, width="stretch")
            actions = result.review.get("actions", [])
            if actions:
                st.markdown("**Prepared actions**")
                st.dataframe(actions, hide_index=True, width="stretch")
            else:
                st.info("No automatic edits were prepared. Feedback notes will remain attached to the pitchbook.")
            feedback_rows = result.review.get("feedback", [])
            if feedback_rows:
                st.markdown("**Feedback status**")
                st.dataframe([{k: v for k, v in item.items() if k not in ("source",)} |
                              {"source": (item.get("source") or {}).get("key", "Pitchbook-level")}
                              for item in feedback_rows], hide_index=True, width="stretch")
            resolve_all = st.checkbox("Resolve all feedback after reviewing these changes", value=False,
                                      key="resolve_pitchbook_feedback")
            c1, c2 = st.columns(2)
            if c1.button("Approve final pitchbook", type="primary", key="finalize_pitchbook"):
                ss.pitchbook_result = ss.pitchbook_wf.resume(result.run_id,
                    {"action": "approve", "resolve_all": resolve_all})
                st.rerun()
            if c2.button("Reject revised draft", key="reject_revised_pitchbook"):
                ss.pitchbook_result = ss.pitchbook_wf.resume(result.run_id, {"action": "reject"})
                st.rerun()
        elif result.status == "completed":
            st.success(f"Pitchbook completed · v{result.state['version']} · {result.state['location']}")
            manifest = json.loads(repos.pitchbooks.read(result.state["location"], "pitchbook.json"))
            st.dataframe([{"position": i + 1, "exhibit": x["exhibit_name"], "source": x["key"],
                           "version": x["version"]} for i, x in enumerate(manifest["slides"])],
                         hide_index=True, width="stretch")
            st.download_button("Download final PowerPoint", repos.pitchbooks.read(
                result.state["location"], result.state.get("file_name", "")) if result.state.get("file_name") else
                repos.pitchbooks.read(result.state["location"], manifest["title"].replace(" ", "_") + ".pptx"),
                file_name=manifest["title"].replace(" ", "_") + ".pptx", type="primary")
            st.json(manifest["feedback"], expanded=False)
        elif result.status == "rejected":
            st.warning(result.state.get("error", "Draft rejected"))
        else:
            st.error(result.state.get("error", result.status))

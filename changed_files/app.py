"""
Streamlit UI for the template-creation workflow.   Run:  streamlit run app.py

The UI only talks to TemplateCreationWorkflow (start / resume / result). It never calls nodes,
repository internals or LangGraph directly, so it can later be replaced by an API.
"""
import json

import streamlit as st

import config
from core.template_catalog import TemplateCatalog
from integrations.stories import LocalStorySource
from repositories.base import get_repositories
from template_creation.runner import TemplateCreationWorkflow
from template_creation.services import default_services

st.set_page_config(page_title="Template creation", layout="wide")
st.title("Template creation workflow")
repos = get_repositories()
ss = st.session_state

tab_run, tab_registry, tab_ds = st.tabs(["Create / update template", "Template registry", "Datasources"])


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
    left, right = st.columns([1, 2])

    with left:
        st.subheader("Input")
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
with tab_registry:
    entries = TemplateCatalog(repos).list()
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

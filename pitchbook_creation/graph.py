from __future__ import annotations

import json
import io
import re
from datetime import datetime, timezone

from langgraph.checkpoint.memory import InMemorySaver
from langgraph.graph import END, START, StateGraph
from langgraph.types import interrupt
from pptx import Presentation

import config
from core.pptx_assembler import assemble_exhibits, slide_components
from pitchbook_creation.state import PitchbookState
from pitchbook_creation.services import PitchbookServices
from pitchbook_creation.feedback import interpret_feedback


def _target_size(state: PitchbookState) -> tuple[int, int] | None:
    profile = state.get("criteria", {}).get("sample_profile", {})
    width, height = profile.get("width_inches"), profile.get("height_inches")
    if width and height:
        return round(float(width) * 914400), round(float(height) * 914400)
    return None


def _resolve(state: PitchbookState, svc: PitchbookServices):
    registry = svc.repos.exhibit_registry
    if not registry or not svc.repos.exhibits:
        raise ValueError("Exhibit storage is not configured")
    refs = list(state.get("references", []))
    criteria = state.get("criteria", {})
    if not refs and criteria.get("acronym"):
        acronym = str(criteria["acronym"]).casefold()
        asset_type = str(criteria.get("asset_type", "")).casefold()
        year = str(criteria.get("financial_year", "")).casefold()
        for entry in registry.list():
            fields = entry.get("fields", {})
            if acronym not in str(fields.get("acronym", "")).casefold():
                continue
            if asset_type and asset_type not in str(fields.get("asset_type", "")).casefold():
                continue
            latest = entry.get("versions", [])[-1:] 
            for rec in latest:
                if year and year not in json.dumps({"params": rec.get("params", {}), "key": entry["key"]}).casefold():
                    continue
                refs.append({"key": entry["key"], "version": rec["version"]})
    selected = []
    for ref in refs:
        key = ref.get("key", "")
        entry = registry.get(key)
        if not entry:
            raise ValueError(f"No saved exhibit matches {key!r}")
        rec = next((v for v in entry["versions"] if v["version"] == int(ref.get("version", 0))), None)
        if not rec:
            raise ValueError(f"Exhibit {key!r} has no version {ref.get('version')}")
        selected.append({"key": key, "key_hash": entry["key_hash"], "fields": entry["fields"],
                         "uid": f"{entry['key_hash']}:v{rec['version']}",
                         "version": rec["version"], "location": rec["location"],
                         "file_name": rec["file_name"], "exhibit_name": entry["exhibit_name"],
                         "sample_uid": ref.get("sample_uid"), "sample_slide": ref.get("sample_slide")})
    if not selected:
        raise ValueError("Select at least one exhibit slide")
    return {"selected": selected, "log": [f"resolved {len(selected)} exhibit reference(s)"]}


def _assemble(state: PitchbookState, svc: PitchbookServices):
    work = config.WORK_DIR / f"pitchbook_{state['run_id']}"
    work.mkdir(parents=True, exist_ok=True)
    items = [(ref, svc.repos.exhibits.read(ref["location"], ref["file_name"])) for ref in state["selected"]]
    payload = assemble_exhibits(items, target_size=_target_size(state))
    prs = Presentation(io.BytesIO(payload))
    title = state.get("title") or "PPT-AI Pitchbook"
    prs.core_properties.title = title
    prs.core_properties.subject = "; ".join(x["key"] for x in state["selected"])
    prs.core_properties.keywords = "pitchbook; " + "; ".join(x["key_hash"] for x in state["selected"])
    draft_name = re.sub(r"[^A-Za-z0-9._ -]+", "", title).strip().replace(" ", "_") or "pitchbook"
    path = work / f"{draft_name}.pptx"
    prs.save(path)
    return {"draft_path": str(path), "draft_name": f"{draft_name}.pptx",
            "review": {"slides": state["selected"], "components": slide_components(path.read_bytes())},
            "log": [f"assembled draft with {len(items)} slide(s)"]}


def _review(state: PitchbookState, svc: PitchbookServices):
    decision = interrupt({"stage": "draft_review", "title": state.get("title", "PPT-AI Pitchbook"),
                          "slides": state["review"]["slides"], "components": state["review"]["components"]})
    if (decision or {}).get("action") != "approve":
        return {"status": "rejected", "error": "Pitchbook draft rejected", "log": ["rejected"]}
    order = decision.get("order") or list(range(len(state["selected"])))
    if sorted(order) != list(range(len(state["selected"]))):
        return {"status": "failed", "error": "Slide order must include each slide exactly once", "log": ["invalid order"]}
    selected = [state["selected"][i] for i in order]
    title = (decision.get("title") or state.get("title") or "PPT-AI Pitchbook").strip()
    feedback = decision.get("feedback") or []
    for index, item in enumerate(feedback):
        item.setdefault("id", f"FB-{index + 1:03d}")
        item["status"] = "open"
    return {"selected": selected, "title": title, "feedback": feedback,
            "use_llm_feedback": bool(decision.get("use_llm_feedback", True)),
            "log": [f"reviewed; {len(feedback)} feedback item(s)"]}


def _feedback_agent(state: PitchbookState, svc: PitchbookServices):
    selected = state["selected"]
    title = state.get("title", "PPT-AI Pitchbook")
    components = []
    for ref in selected:
        source_index = next(i for i, candidate in enumerate(state["review"]["slides"])
                            if candidate["key"] == ref["key"] and candidate["version"] == ref["version"])
        components.append(state["review"]["components"][source_index])
    feedback = state.get("feedback", [])
    try:
        actions = interpret_feedback(feedback, selected, components, svc.emit,
                                     use_llm=state.get("use_llm_feedback", True))
    except Exception as e:
        actions = []
        for item in feedback:
            if item.get("status") != "unresolved":
                item["status"] = "open"
                item["resolution_note"] = f"Feedback agent could not complete this pass: {e}"
    for action in actions:
        if action["type"] == "rename":
            title = action["title"]
        elif action["type"] == "move":
            index = next(i for i, ref in enumerate(selected)
                         if ref["key"] == action["key"] and ref["version"] == action["version"])
            ref = selected.pop(index)
            selected.insert(action["position"] - 1, ref)
    applied_feedback = {action.get("feedback_id") for action in actions}
    for item in feedback:
        if item.get("id") in applied_feedback:
            item["applied"] = True
    components = []
    for ref in selected:
        source_index = next(i for i, candidate in enumerate(state["review"]["slides"])
                            if candidate["key"] == ref["key"] and candidate["version"] == ref["version"])
        components.append(state["review"]["components"][source_index])
    return {"selected": selected, "title": title, "feedback": feedback, "feedback_actions": actions,
            "final_review": {"title": title, "slides": selected, "components": components,
                             "actions": actions, "feedback": feedback},
            "log": [f"feedback agent prepared {len(actions)} revision action(s)"]}


def _final_review(state: PitchbookState, svc: PitchbookServices):
    decision = interrupt({"stage": "feedback_review", **state["final_review"]})
    if (decision or {}).get("action") != "approve":
        return {"status": "rejected", "error": "Revised pitchbook draft rejected", "log": ["revised draft rejected"]}
    feedback = state.get("feedback", [])
    if decision.get("resolve_all"):
        for item in feedback:
            if item.get("status") != "unresolved":
                item["status"] = "resolved"
    return {"feedback": feedback, "status": "approved", "log": ["revised draft approved"]}


def _persist(state: PitchbookState, svc: PitchbookServices):
    selected = state["selected"]
    items = [(ref, svc.repos.exhibits.read(ref["location"], ref["file_name"])) for ref in selected]
    payload = assemble_exhibits(items, target_size=_target_size(state))
    prs = Presentation(io.BytesIO(payload))
    prs.core_properties.title = state["title"]
    prs.core_properties.subject = "; ".join(x["key"] for x in selected)
    # Apply validated component edits while leaving saved exhibit source files untouched.
    for action in state.get("feedback_actions", []):
        if action.get("type") != "replace_text":
            continue
        slide_index = next((i for i, ref in enumerate(selected)
                            if ref["key"] == action.get("key") and ref["version"] == action.get("version")), None)
        shape_id = action.get("shape_id")
        if slide_index is None or shape_id is None:
            continue
        shape = next((shape for shape in prs.slides[slide_index].shapes if shape.shape_id == int(shape_id)), None)
        if shape is None or not getattr(shape, "has_text_frame", False):
            continue
        paragraphs = shape.text_frame.paragraphs
        if paragraphs and paragraphs[0].runs:
            paragraphs[0].runs[0].text = action["replacement_text"]
            for paragraph in paragraphs:
                for run in paragraph.runs[1:] if paragraph is paragraphs[0] else paragraph.runs:
                    run.text = ""
        elif paragraphs:
            paragraphs[0].text = action["replacement_text"]
        for feedback in state.get("feedback", []):
            source = feedback.get("source") or {}
            if (source.get("key") == action.get("key") and
                    source.get("version") == action.get("version") and
                    (feedback.get("shape_id") == shape_id or feedback.get("id") == action.get("feedback_id"))):
                feedback["applied"] = True
    out = io.BytesIO()
    prs.save(out)
    entry = svc.repos.pitchbook_registry.get(state["run_id"])
    version = (entry["current_version"] + 1) if entry else 1
    now = datetime.now(timezone.utc).isoformat(timespec="seconds")
    title = re.sub(r"[^A-Za-z0-9._ -]+", "", state["title"]).strip().replace(" ", "_") or "pitchbook"
    file_name = f"{title}.pptx"
    manifest = {"id": state["run_id"], "title": state["title"], "version": version,
                "created_at": now, "status": "completed", "slides": selected,
                "criteria": state.get("criteria", {}), "feedback": state.get("feedback", [])}
    location = svc.repos.pitchbooks.save_version(state["run_id"], version,
        {file_name: out.getvalue(), "pitchbook.json": json.dumps(manifest, indent=2, ensure_ascii=False),
         "feedback.json": json.dumps(state.get("feedback", []), indent=2, ensure_ascii=False)})
    history = (entry or {}).get("versions", []) + [{"version": version, "location": location,
        "file_name": file_name, "generated_at": now, "slides": len(selected)}]
    svc.repos.pitchbook_registry.upsert({"key": state["run_id"], "title": state["title"],
        "current_version": version, "current_location": location, "created_at": (entry or {}).get("created_at", now),
        "updated_at": now, "versions": history})
    return {"location": location, "version": version, "file_name": file_name,
            "status": "completed", "log": [f"saved pitchbook v{version}"]}


def build_graph(svc: PitchbookServices, checkpointer=None):
    g = StateGraph(PitchbookState)
    g.add_node("resolve", lambda s: _resolve(s, svc))
    g.add_node("assemble", lambda s: _assemble(s, svc))
    g.add_node("review", lambda s: _review(s, svc))
    g.add_node("feedback_agent", lambda s: _feedback_agent(s, svc))
    g.add_node("final_review", lambda s: _final_review(s, svc))
    g.add_node("persist", lambda s: _persist(s, svc))
    g.add_edge(START, "resolve")
    g.add_edge("resolve", "assemble")
    g.add_edge("assemble", "review")
    g.add_conditional_edges("review", lambda s: END if s.get("status") else "feedback_agent", ["feedback_agent", END])
    g.add_edge("feedback_agent", "final_review")
    g.add_conditional_edges("final_review", lambda s: "persist" if s.get("status") == "approved" else END,
                            ["persist", END])
    g.add_edge("persist", END)
    return g.compile(checkpointer=checkpointer or InMemorySaver())

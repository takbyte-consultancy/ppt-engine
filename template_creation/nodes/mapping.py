"""Nodes: map_components (code proposes, the LLM reviews/completes), validate."""
from __future__ import annotations

from pathlib import Path

from pptx import Presentation

from core import alias_engine as eng
from core.alias_schema import AliasDocument
from template_creation import rules
from template_creation.mapper import MappingContext


def map_components(state, svc):
    attempts = state.get("attempts", 0)
    out = {}
    if attempts == 0:
        prs = Presentation(state["platinum_path"])
        previous = AliasDocument.model_validate(state["previous_doc"]) if state.get("previous_doc") else None
        prop = eng.propose(prs, state["sample"], state["markers"], rules.callout_ids(state["markers"]) |
                           set(state["markers"].get("callout_connector_ids", [])), previous)
        rules.enrich(prop.doc, state["key"], state["datasource_id"], state.get("required_params", []))
        out.update(proposal=prop.doc.model_dump(exclude_none=True), unresolved=prop.unresolved)
        svc.emit("node", f"Code proposed {len(prop.doc.occurrences)} placeholders "
                         f"({sum(o.confidence == 'low' for o in prop.doc.occurrences)} ambiguous), "
                         f"{len(prop.doc.labels)} static labels, {len(prop.unresolved)} unresolved")
    proposal = out.get("proposal") or state["proposal"]
    ctx = MappingContext(
        platinum_pptx=state["platinum_path"], sample=state["sample"], markers=state["markers"],
        key=state["key"], datasource_id=state["datasource_id"], required_params=state.get("required_params", []),
        draft=state.get("alias_doc") or proposal, unresolved=out.get("unresolved", state.get("unresolved", [])),
        previous=state.get("previous_doc"), feedback=state.get("validation_errors") or [],
        preview_png=Path(state["preview_path"]).read_bytes() if state.get("preview_path") else None)
    svc.emit("node", f"Mapping agent, attempt {attempts + 1}")
    doc = svc.mapper.map(ctx, on_event=svc.on_event)
    out.update(alias_doc=doc, attempts=attempts + 1, log=[f"mapping attempt {attempts + 1}"])
    return out


def check(state, doc_dict: dict | None) -> tuple[dict | None, list[str], list[str]]:
    """Shared by validate and human_review: parse, enrich, validate."""
    if not doc_dict:
        return None, ["The mapping agent did not submit a valid alias document"], []
    try:
        doc = AliasDocument.model_validate(doc_dict)
    except Exception as e:
        return None, [f"Alias document does not match the schema: {e}"], []
    rules.enrich(doc, state["key"], state["datasource_id"], state.get("required_params", []))
    errors, warnings = rules.validate_all(doc, Presentation(state["platinum_path"]), state["sample"], state["markers"])
    return doc.model_dump(exclude_none=True), errors, warnings


def validate(state, svc):
    import config
    doc, errors, warnings = check(state, state.get("alias_doc"))
    svc.emit("node", "Validation passed" if not errors else f"Validation: {len(errors)} error(s)")
    out = {"validation_errors": errors, "validation_warnings": warnings, "log": [f"validate: {len(errors)} errors"]}
    if errors and state.get("attempts", 0) >= config.MAX_MAPPING_ATTEMPTS:
        out.update(status="failed", error=f"Alias document still invalid after {state['attempts']} attempt(s): "
                                          + "; ".join(errors[:5]))
    if doc:
        out["alias_doc"] = doc
        out["alias_diff"] = rules.diff_docs(state.get("previous_doc"), doc)
    return out

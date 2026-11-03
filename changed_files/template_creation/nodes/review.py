"""Node: human_review. Pauses the graph (LangGraph interrupt) until a person approves or edits."""
from __future__ import annotations

from langgraph.types import interrupt

from template_creation import rules
from template_creation.nodes.mapping import check


def human_review(state, svc):
    decision = interrupt({
        "template_key": state["key"]["key"],
        "mode": state["mode"],
        "alias_doc": state["alias_doc"],
        "diff": state.get("alias_diff"),
        "callouts": state["markers"]["callouts"],
        "locked_runs": state["markers"]["locked_runs"],
        "warnings": state.get("validation_warnings", []),
        "errors": state.get("review_errors", []),
    })
    action = (decision or {}).get("action", "approve")
    if action == "reject":
        return {"status": "rejected", "error": decision.get("reason", "Rejected in review"),
                "log": ["rejected in review"]}
    edited = decision.get("alias_doc")
    if edited is None:
        return {"review_errors": [], "log": ["approved in review"]}
    doc, errors, warnings = check(state, edited)
    if errors:
        return {"review_errors": errors, "log": ["edited alias document invalid, back to review"]}
    return {"alias_doc": doc, "review_errors": [], "validation_warnings": warnings,
            "alias_diff": rules.diff_docs(state.get("previous_doc"), doc), "log": ["approved with edits"]}

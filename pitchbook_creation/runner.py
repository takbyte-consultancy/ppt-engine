from __future__ import annotations

import uuid
from dataclasses import dataclass

from langgraph.types import Command

from pitchbook_creation.graph import build_graph
from pitchbook_creation.services import PitchbookServices


@dataclass
class PitchbookResult:
    run_id: str
    status: str
    review: dict | None
    state: dict


class PitchbookCreationWorkflow:
    def __init__(self, services: PitchbookServices, checkpointer=None):
        self.graph = build_graph(services, checkpointer)

    def start(self, references: list[dict], *, title: str = "PPT-AI Pitchbook", mode: str = "manual",
              criteria: dict | None = None) -> PitchbookResult:
        run_id = uuid.uuid4().hex[:12]
        self.graph.invoke({"run_id": run_id, "references": references, "title": title,
                           "mode": mode, "criteria": criteria or {}, "log": []}, self._cfg(run_id))
        return self.result(run_id)

    def resume(self, run_id: str, decision: dict) -> PitchbookResult:
        self.graph.invoke(Command(resume=decision), self._cfg(run_id))
        return self.result(run_id)

    def result(self, run_id: str) -> PitchbookResult:
        snap = self.graph.get_state(self._cfg(run_id))
        interrupts = getattr(snap, "interrupts", ()) or ()
        values = dict(snap.values)
        stage = interrupts[0].value.get("stage") if interrupts and isinstance(interrupts[0].value, dict) else None
        status = values.get("status") or ({"draft_review": "awaiting_review",
            "feedback_review": "awaiting_feedback_review"}.get(stage, "running") if interrupts else "running")
        return PitchbookResult(run_id, status, interrupts[0].value if interrupts else None, values)

    @staticmethod
    def _cfg(run_id: str) -> dict:
        return {"configurable": {"thread_id": run_id}}

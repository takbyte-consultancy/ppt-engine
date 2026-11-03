"""
Thin facade over the graph so callers (Streamlit, CLI, tests, a future API) never deal with
LangGraph threads, checkpoints or interrupts directly.
"""
from __future__ import annotations

import uuid
from dataclasses import dataclass

from langgraph.types import Command

import config
from template_creation.graph import build_graph
from template_creation.services import Services


@dataclass
class RunResult:
    run_id: str
    status: str                # awaiting_review | created | updated | unchanged | failed | rejected | fidelity_failed
    review: dict | None        # payload to show a reviewer when status == awaiting_review
    state: dict

    @property
    def error(self) -> str:
        return self.state.get("error", "")


class TemplateCreationWorkflow:
    def __init__(self, services: Services, checkpointer=None):
        self.svc = services
        self.graph = build_graph(services, checkpointer)

    def start(self, story_ref: str, *, overrides: dict | None = None, params: dict | None = None,
              slide_index: int | None = None, attachment_name: str | None = None,
              force: bool = False) -> RunResult:
        run_id = uuid.uuid4().hex[:12]
        work = config.WORK_DIR / run_id
        work.mkdir(parents=True, exist_ok=True)
        self.graph.invoke({"run_id": run_id, "work_dir": str(work), "story_ref": story_ref,
                           "overrides": overrides or {}, "params": params or {},
                           "slide_index": slide_index, "attachment_name": attachment_name,
                           "force": force, "attempts": 0, "log": []}, self._cfg(run_id))
        return self.result(run_id)

    def resume(self, run_id: str, decision: dict) -> RunResult:
        """decision: {"action": "approve"} | {"action": "approve", "mapping": {...}} | {"action": "reject", "reason": "..."}"""
        self.graph.invoke(Command(resume=decision), self._cfg(run_id))
        return self.result(run_id)

    def result(self, run_id: str) -> RunResult:
        snap = self.graph.get_state(self._cfg(run_id))
        interrupts = getattr(snap, "interrupts", ()) or ()
        values = dict(snap.values)
        status = values.get("status") or ("awaiting_review" if interrupts else "running")
        return RunResult(run_id, status, interrupts[0].value if interrupts else None, values)

    @staticmethod
    def _cfg(run_id: str) -> dict:
        return {"configurable": {"thread_id": run_id}}

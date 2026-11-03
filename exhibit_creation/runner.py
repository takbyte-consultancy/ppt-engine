"""Facade over the Exhibit graph: one call creates one slide."""
from __future__ import annotations

import uuid
from dataclasses import dataclass

import config
from exhibit_creation.graph import build_graph
from exhibit_creation.services import ExhibitServices


@dataclass
class ExhibitResult:
    status: str                # created | updated | unchanged | failed
    state: dict

    @property
    def error(self) -> str:
        return self.state.get("error", "")

    @property
    def warnings(self) -> list[str]:
        return self.state.get("fill_warnings", []) + self.state.get("schema_warnings", [])


class ExhibitCreationWorkflow:
    def __init__(self, services: ExhibitServices):
        self.graph = build_graph(services)

    def run(self, asset_type: str, acronym: str, *, params: dict | None = None,
            template_version: int | None = None, force: bool = False) -> ExhibitResult:
        run_id = uuid.uuid4().hex[:12]
        work = config.WORK_DIR / f"exhibit_{run_id}"
        work.mkdir(parents=True, exist_ok=True)
        state = self.graph.invoke({"run_id": run_id, "work_dir": str(work), "asset_type": asset_type,
                                   "acronym": acronym, "params": params or {}, "template_version": template_version,
                                   "force": force, "log": []})
        return ExhibitResult(state.get("status", "failed"), state)

from __future__ import annotations

import operator
from typing import Annotated, TypedDict


class PitchbookState(TypedDict, total=False):
    run_id: str
    title: str
    mode: str
    references: list[dict]
    criteria: dict
    selected: list[dict]
    draft_path: str
    draft_name: str
    review: dict
    final_review: dict
    feedback_actions: list[dict]
    use_llm_feedback: bool
    feedback: list[dict]
    location: str
    version: int
    status: str
    error: str
    log: Annotated[list[str], operator.add]

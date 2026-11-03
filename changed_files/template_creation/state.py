"""The state passed between graph nodes. JSON-serialisable so LangGraph can checkpoint it."""
from __future__ import annotations

import operator
from typing import Annotated, Any, TypedDict

TERMINAL = {"failed", "unchanged", "rejected", "fidelity_failed", "created", "updated"}


class TemplateState(TypedDict, total=False):
    # inputs
    run_id: str
    story_ref: str                 # JIRA key, or a label for local runs
    overrides: dict                # asset_type / acronym / datasource_id typed by a human
    params: dict                   # datasource run-time params
    slide_index: int | None        # which slide of a multi-slide Platinum copy
    attachment_name: str | None
    force: bool                    # re-process even if the Platinum copy is unchanged
    # produced along the way
    work_dir: str
    story: dict
    platinum_path: str
    platinum_sha256: str
    key: dict
    mode: str                      # create | update
    registry_entry: dict | None
    previous_doc: dict | None      # previous version's alias document (updates)
    datasource_id: str
    required_params: list[str]
    metadata_path: str
    markers: dict
    original_names: dict
    preview_path: str | None
    sample: dict
    response_schema: dict
    proposal: dict                 # code's alias proposal (before the agent)
    unresolved: list[dict]
    alias_doc: dict | None         # the alias document (alias.json) being built
    validation_errors: list[str]
    validation_warnings: list[str]
    attempts: int
    review_errors: list[str]
    alias_diff: dict
    template_path: str
    template_metadata_path: str
    fidelity: dict
    preview_filled_path: str
    placeholders: int
    removed_shapes: list[str]
    location: str
    version: int
    # outcome
    status: str
    error: str
    log: Annotated[list[str], operator.add]


class WorkflowError(Exception):
    """An expected failure: the graph ends with status 'failed' and this message."""

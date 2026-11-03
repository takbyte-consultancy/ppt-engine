"""State passed between Exhibit-creation nodes (JSON-serialisable for LangGraph checkpoints)."""
from __future__ import annotations

import operator
from typing import Annotated, TypedDict

TERMINAL = {"failed", "unchanged", "created", "updated"}


class ExhibitState(TypedDict, total=False):
    # inputs
    run_id: str
    asset_type: str
    acronym: str
    params: dict                   # datasource run-time params
    template_version: int | None   # None = latest
    force: bool                    # regenerate even if template + data are unchanged
    work_dir: str
    # produced
    key: dict
    exhibit_name: str              # e.g. BOND_AMERICAN_BOND_1
    folder: str                    # <key_hash>__<asset_type>__<acronym>
    template: dict                 # {version, location, datasource_id}
    aliases: dict                  # alias.json of that template version
    data: dict
    data_sha256: str
    schema_warnings: list[str]
    exhibit_path: str
    preview_path: str | None
    fidelity: dict
    fill_warnings: list[str]
    location: str
    version: int
    # outcome
    status: str
    error: str
    log: Annotated[list[str], operator.add]


class ExhibitError(Exception):
    """An expected failure: the graph ends with status 'failed' and this message."""

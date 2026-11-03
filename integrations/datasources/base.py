"""
DataSource interface + factory. A datasource definition (from the repository) looks like:
  {"id": "ds_x", "type": "json_file" | "excel_file" | "http", "location": "...", "params": [...], ...}
Every implementation returns a plain dict, so mapping works the same whatever the source is.
"""
from __future__ import annotations

from pathlib import Path
from typing import Protocol

import config


class DataSource(Protocol):
    definition: dict

    def fetch(self, params: dict | None = None) -> dict: ...


def resolve_location(location: str, params: dict) -> str:
    loc = location.format(**params) if params else location
    return loc[len("file://"):] if loc.startswith("file://") else loc


def local_path(location: str, params: dict) -> Path:
    p = Path(resolve_location(location, params))
    return p if p.is_absolute() else config.ROOT / p


def check_params(definition: dict, params: dict | None) -> dict:
    params = params or {}
    missing = [p for p in definition.get("params", []) if not params.get(p)]
    if missing:
        raise ValueError(f"Datasource '{definition['id']}' needs params: {missing}")
    return params


def build_datasource(definition: dict) -> DataSource:
    kind = definition.get("type")
    if kind == "json_file":
        from integrations.datasources.json_source import JsonFileSource
        return JsonFileSource(definition)
    if kind == "excel_file":
        from integrations.datasources.excel_source import ExcelFileSource
        return ExcelFileSource(definition)
    if kind == "http":
        from integrations.datasources.http_source import HttpSource
        return HttpSource(definition)
    raise ValueError(f"Unknown datasource type '{kind}' for '{definition.get('id')}'")

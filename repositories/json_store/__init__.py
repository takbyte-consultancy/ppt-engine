"""JSON-file implementations of the repository interfaces."""
from __future__ import annotations

import json
import os
import tempfile
from pathlib import Path

import config


def _read_json(path: Path, default):
    return json.loads(path.read_text(encoding="utf-8")) if path.exists() else default


def _write_json_atomic(path: Path, data) -> None:
    """Write to a temp file then rename, so a crash never leaves a half-written registry."""
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp = tempfile.mkstemp(dir=path.parent, suffix=".tmp")
    with os.fdopen(fd, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    os.replace(tmp, path)


class JsonRegistryRepository:
    def __init__(self, path: Path | None = None):
        self.path = path or config.DATA_DIR / "template_registry.json"

    def get(self, key: str) -> dict | None:
        return _read_json(self.path, {}).get(key)

    def upsert(self, entry: dict) -> None:
        data = _read_json(self.path, {})
        data[entry["key"]] = entry
        _write_json_atomic(self.path, data)

    def list(self) -> list[dict]:
        return list(_read_json(self.path, {}).values())


class FileTemplateRepository:
    def __init__(self, root: Path | None = None):
        self.root = root or config.TEMPLATES_DIR

    def save_version(self, template_id: str, version: int, files: dict[str, bytes | str]) -> str:
        out = self.root / template_id / f"v{version}"
        if out.exists():
            raise FileExistsError(f"{out} already exists; versions are immutable")
        out.mkdir(parents=True)
        for name, content in files.items():
            mode = "wb" if isinstance(content, bytes) else "w"
            with open(out / name, mode, **({} if mode == "wb" else {"encoding": "utf-8"})) as f:
                f.write(content)
        return str(out.relative_to(config.ROOT)).replace("\\", "/")

    def read(self, location: str, name: str) -> bytes:
        return (config.ROOT / location / name).read_bytes()

    def exists(self, location: str, name: str) -> bool:
        return (config.ROOT / location / name).exists()


class JsonDatasourceRepository:
    def __init__(self, registry: JsonRegistryRepository, path: Path | None = None):
        self.path = path or config.DATA_DIR / "datasources.json"
        self.registry = registry

    def get(self, ds_id: str) -> dict:
        data = _read_json(self.path, {})
        if ds_id not in data:
            raise KeyError(f"Datasource '{ds_id}' is not registered. Known: {sorted(data)}")
        return {"id": ds_id, **data[ds_id]}

    def list(self) -> list[dict]:
        return [{"id": k, **v} for k, v in _read_json(self.path, {}).items()]

    def register(self, definition: dict) -> None:
        data = _read_json(self.path, {})
        d = dict(definition)
        data[d.pop("id")] = d
        _write_json_atomic(self.path, data)

    def templates_using(self, ds_id: str) -> list[str]:
        return sorted(e["key"] for e in self.registry.list() if e.get("datasource_id") == ds_id)

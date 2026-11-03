"""
Read-side API over saved exhibits. This is what the Book assembly agent (Agent 3) uses:
it asks by asset type / acronym and never deals with folders or registry internals.
"""
from __future__ import annotations

import io
import json
import re
from dataclasses import dataclass

from pptx import Presentation

import config
from core.template_key import build_template_key, normalise
from repositories.base import Repositories, get_repositories


def exhibit_name(fields: dict) -> str:
    """Slide / file name: asset type + acronym, e.g. BOND_AMERICAN_BOND_1."""
    return "_".join(normalise(fields[f]) for f in config.TEMPLATE_KEY_FIELDS)


def exhibit_folder(key: dict) -> str:
    """Folder per template key: <key_hash>__<asset_type>__<acronym>."""
    return f"{key['key_hash']}__{key['template_id']}"


@dataclass
class ExhibitHandle:
    key: str
    version: int
    location: str
    file_name: str
    record: dict
    repos: Repositories

    @property
    def metadata(self) -> dict:
        return json.loads(self.repos.exhibits.read(self.location, "exhibit.json"))

    @property
    def path(self):
        return config.ROOT / self.location / self.file_name

    def pptx_bytes(self) -> bytes:
        return self.repos.exhibits.read(self.location, self.file_name)

    def presentation(self) -> Presentation:
        return Presentation(io.BytesIO(self.pptx_bytes()))

    def preview_png(self) -> bytes | None:
        return self.repos.exhibits.read(self.location, "preview.png") \
            if self.repos.exhibits.exists(self.location, "preview.png") else None


class ExhibitCatalog:
    def __init__(self, repos: Repositories | None = None):
        self.repos = repos or get_repositories()

    def _handle(self, entry: dict, rec: dict) -> ExhibitHandle:
        return ExhibitHandle(entry["key"], rec["version"], rec["location"], rec["file_name"], rec, self.repos)

    def find(self, asset_type: str, acronym: str, version: int | None = None) -> ExhibitHandle | None:
        key = build_template_key({"asset_type": asset_type, "acronym": acronym})
        entry = self.repos.exhibit_registry.get(key.key)
        if not entry:
            return None
        rec = entry["versions"][-1] if version is None else \
            next((v for v in entry["versions"] if v["version"] == version), None)
        return self._handle(entry, rec) if rec else None

    def search(self, asset_type: str | None = None, acronym: str | None = None,
               text: str | None = None) -> list[ExhibitHandle]:
        """Latest exhibit per key, filtered by asset type / acronym (normalised) or free text."""
        at = normalise(asset_type) if asset_type else None
        ac = normalise(acronym) if acronym else None
        out = []
        for e in sorted(self.repos.exhibit_registry.list(), key=lambda e: e["key"]):
            f = {k: normalise(v) for k, v in e["fields"].items()}
            if at and f.get("asset_type") != at:
                continue
            if ac and f.get("acronym") != ac:
                continue
            if text and not re.search(re.escape(text), json.dumps(e), re.IGNORECASE):
                continue
            out.append(self._handle(e, e["versions"][-1]))
        return out

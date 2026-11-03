"""
Read-side API over saved templates. This is what the Exhibit (slide creation) agent will use:
it only knows asset type + acronym, never folders or registry internals.
"""
from __future__ import annotations

import io
import json
from dataclasses import dataclass

from pptx import Presentation

from core.alias_schema import AliasDocument
from core.template_key import build_template_key
from repositories.base import Repositories, get_repositories


@dataclass
class TemplateHandle:
    key: str
    version: int
    location: str
    datasource_id: str
    entry: dict
    repos: Repositories

    @property
    def aliases(self) -> AliasDocument:
        return AliasDocument.model_validate_json(self.repos.templates.read(self.location, "alias.json"))

    @property
    def metadata(self) -> dict:
        return json.loads(self.repos.templates.read(self.location, "metadata.json"))

    def presentation(self) -> Presentation:
        return Presentation(io.BytesIO(self.repos.templates.read(self.location, "template.pptx")))


class TemplateCatalog:
    def __init__(self, repos: Repositories | None = None):
        self.repos = repos or get_repositories()

    def find(self, asset_type: str, acronym: str, version: int | None = None) -> TemplateHandle | None:
        key = build_template_key({"asset_type": asset_type, "acronym": acronym})
        entry = self.repos.registry.get(key.key)
        if not entry:
            return None
        rec = entry["versions"][-1] if version is None else \
            next((v for v in entry["versions"] if v["version"] == version), None)
        if rec is None:
            return None
        return TemplateHandle(entry["key"], rec["version"], rec["location"], entry["datasource_id"], entry, self.repos)

    def list(self) -> list[dict]:
        return sorted(self.repos.registry.list(), key=lambda e: e["key"])

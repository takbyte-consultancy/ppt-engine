"""
Storage interfaces. The workflow depends only on these; today's implementation is JSON files
(repositories/json_store), tomorrow's is SQL tables with the same methods.
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Protocol


class RegistryRepository(Protocol):
    """template key -> template record (current version, path, history)."""
    def get(self, key: str) -> dict | None: ...
    def upsert(self, entry: dict) -> None: ...
    def list(self) -> list[dict]: ...


class TemplateRepository(Protocol):
    """Stores the files of one template version."""
    def save_version(self, template_id: str, version: int, files: dict[str, bytes | str]) -> str: ...
    def read(self, location: str, name: str) -> bytes: ...
    def exists(self, location: str, name: str) -> bool: ...


class DatasourceRepository(Protocol):
    def get(self, ds_id: str) -> dict: ...
    def list(self) -> list[dict]: ...
    def register(self, definition: dict) -> None: ...
    def templates_using(self, ds_id: str) -> list[str]: ...


@dataclass
class Repositories:
    registry: RegistryRepository                 # template key -> template versions
    templates: TemplateRepository                # template version files
    datasources: DatasourceRepository
    exhibit_registry: RegistryRepository | None = None   # template key -> exhibit versions
    exhibits: TemplateRepository | None = None           # exhibit version files (same versioned layout)
    pitchbook_registry: RegistryRepository | None = None
    pitchbooks: TemplateRepository | None = None


def get_repositories() -> Repositories:
    """The single place to switch storage backends."""
    import config
    from repositories.json_store import JsonDatasourceRepository, JsonRegistryRepository, FileTemplateRepository
    registry = JsonRegistryRepository()
    return Repositories(registry, FileTemplateRepository(), JsonDatasourceRepository(registry),
                        JsonRegistryRepository(config.DATA_DIR / "exhibit_registry.json"),
                        FileTemplateRepository(config.EXHIBITS_DIR),
                        JsonRegistryRepository(config.DATA_DIR / "pitchbook_registry.json"),
                        FileTemplateRepository(config.PITCHBOOKS_DIR))

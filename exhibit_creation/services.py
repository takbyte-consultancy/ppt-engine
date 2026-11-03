"""Dependencies injected into the Exhibit-creation graph."""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable

from repositories.base import Repositories, get_repositories


@dataclass
class ExhibitServices:
    repos: Repositories
    renderer: Callable[[Path], bytes | None] | None = None     # slide -> PNG thumbnail (for Agent 3's UI)
    on_event: Callable[[str, Any], None] | None = None

    def emit(self, kind: str, payload: Any) -> None:
        if self.on_event:
            self.on_event(kind, payload)


def default_exhibit_services(on_event=None, thumbnails: bool = True) -> ExhibitServices:
    from core.preview import render_png
    return ExhibitServices(get_repositories(), render_png if thumbnails else None, on_event)

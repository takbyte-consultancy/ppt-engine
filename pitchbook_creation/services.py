from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Callable

from repositories.base import Repositories, get_repositories


@dataclass
class PitchbookServices:
    repos: Repositories
    on_event: Callable[[str, Any], None] | None = None

    def emit(self, kind: str, payload: Any) -> None:
        if self.on_event:
            self.on_event(kind, payload)


def default_pitchbook_services(on_event=None) -> PitchbookServices:
    return PitchbookServices(get_repositories(), on_event)

"""Dependency container injected into the graph: every external thing the workflow touches."""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable

from integrations.stories import StorySource
from repositories.base import Repositories, get_repositories
from template_creation.mapper import LlmMapper, Mapper


@dataclass
class Services:
    repos: Repositories
    story_source: StorySource
    mapper: Mapper
    renderer: Callable[[Path], bytes | None] | None = None     # slide -> PNG for the LLM (optional)
    on_event: Callable[[str, Any], None] | None = None          # UI progress stream

    def emit(self, kind: str, payload: Any) -> None:
        if self.on_event:
            self.on_event(kind, payload)


def default_services(story_source: StorySource, on_event=None, use_vision: bool = True) -> Services:
    from core.preview import render_png
    return Services(get_repositories(), story_source, LlmMapper(),
                    render_png if use_vision else None, on_event)

"""
StorySource interface. The template workflow only ever sees a Story object, never JIRA itself.
Implementations: JiraStorySource (integrations/jira) and LocalStorySource (manual upload).
"""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Protocol


@dataclass
class Attachment:
    filename: str
    size: int = 0
    created: str = ""
    url: str = ""                    # remote location (JIRA); empty for local files
    local_path: str = ""             # set once downloaded


@dataclass
class Story:
    key: str                         # e.g. PROJ-123, or LOCAL-<hash> for uploads
    summary: str = ""
    description: str = ""
    asset_type: str = ""
    acronym: str = ""
    datasource_id: str = ""
    attachments: list[Attachment] = field(default_factory=list)
    raw_fields: dict = field(default_factory=dict)

    def pptx_attachments(self) -> list[Attachment]:
        return [a for a in self.attachments if a.filename.lower().endswith(".pptx")]


class StorySource(Protocol):
    def get_story(self, story_ref: str) -> Story: ...
    def download_attachment(self, story: Story, attachment: Attachment, dest_dir: Path) -> Path: ...


class LocalStorySource:
    """Use when JIRA isn't available: a Platinum copy file plus fields typed in the UI."""

    def __init__(self, platinum_path: Path, asset_type: str, acronym: str,
                 datasource_id: str = "", summary: str = ""):
        self.path = Path(platinum_path)
        self.fields = dict(asset_type=asset_type, acronym=acronym,
                           datasource_id=datasource_id, summary=summary)

    def get_story(self, story_ref: str = "") -> Story:
        att = Attachment(self.path.name, self.path.stat().st_size, local_path=str(self.path))
        return Story(key=story_ref or f"LOCAL-{self.path.stem}", attachments=[att], **self.fields)

    def download_attachment(self, story: Story, attachment: Attachment, dest_dir: Path) -> Path:
        return Path(attachment.local_path)

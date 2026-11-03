"""JiraStorySource: the StorySource implementation backed by JIRA."""
from __future__ import annotations

from pathlib import Path

import config
from integrations.jira.client import JiraClient
from integrations.jira.story_parser import parse_issue
from integrations.stories import Attachment, Story


class JiraStorySource:
    def __init__(self, client: JiraClient | None = None, field_ids: dict | None = None):
        self.client = client or JiraClient(config.JIRA_BASE_URL, config.JIRA_AUTH_TYPE,
                                           config.JIRA_EMAIL, config.JIRA_API_TOKEN)
        self.field_ids = field_ids if field_ids is not None else config.JIRA_FIELD_IDS

    def get_story(self, story_ref: str) -> Story:
        return parse_issue(self.client.get_issue(story_ref), self.field_ids)

    def download_attachment(self, story: Story, attachment: Attachment, dest_dir: Path) -> Path:
        dest_dir.mkdir(parents=True, exist_ok=True)
        out = dest_dir / attachment.filename
        out.write_bytes(self.client.download(attachment.url))
        attachment.local_path = str(out)
        return out

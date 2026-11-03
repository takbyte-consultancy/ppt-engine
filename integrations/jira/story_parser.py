"""
Turns a raw JIRA issue into a Story.

Asset type / acronym / datasource are read from (in order):
  1. configured custom fields  (JIRA_FIELD_ASSET_TYPE=customfield_10050, ...)
  2. "Label: value" lines in the description  (Asset Type: Fixed Income / Acronym: IMNASMA)
The workflow lets a human override anything still missing.
"""
from __future__ import annotations

import re

from integrations.stories import Attachment, Story

LABELS = {
    "asset_type": r"asset[\s_-]*type",
    "acronym": r"acronym",
    "datasource_id": r"data[\s_-]*source(?:[\s_-]*id)?",
}


def _field_value(v) -> str:
    if v is None:
        return ""
    if isinstance(v, dict):                       # select lists: {"value": "..."} / {"name": "..."}
        return str(v.get("value") or v.get("name") or "")
    if isinstance(v, list):
        return ", ".join(_field_value(x) for x in v)
    return str(v)


def _from_description(text: str, label_regex: str) -> str:
    m = re.search(rf"^\s*\**\s*{label_regex}\s*\**\s*[:=\-]\s*(.+?)\s*$", text or "",
                  flags=re.IGNORECASE | re.MULTILINE)
    return m.group(1).strip() if m else ""


def parse_issue(issue: dict, field_ids: dict[str, str]) -> Story:
    f = issue.get("fields", {})
    desc = f.get("description") or ""
    if isinstance(desc, dict):                    # Atlassian Document Format (API v3) -> plain text
        desc = " ".join(re.findall(r'"text":\s*"([^"]*)"', str(desc)))
    values = {}
    for name, label in LABELS.items():
        fid = field_ids.get(name)
        values[name] = _field_value(f.get(fid)) if fid else ""
        values[name] = values[name] or _from_description(desc, label)
    atts = [Attachment(a["filename"], a.get("size", 0), a.get("created", ""), a.get("content", ""))
            for a in f.get("attachment", [])]
    return Story(key=issue["key"], summary=f.get("summary", ""), description=desc,
                 attachments=atts, raw_fields={k: f.get(v) for k, v in field_ids.items() if v},
                 **values)

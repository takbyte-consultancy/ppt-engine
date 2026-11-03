"""Nodes: load_story, extract_platinum, resolve_key."""
from __future__ import annotations

import hashlib
import json
from dataclasses import asdict
from pathlib import Path

from pptx import Presentation

from core.pptx_inspector import keep_only_slide
from core.template_key import build_template_key
from integrations.stories import Attachment, Story
from template_creation.state import WorkflowError


def _story(d: dict) -> Story:
    d = dict(d)
    d["attachments"] = [Attachment(**a) for a in d.get("attachments", [])]
    return Story(**d)


def load_story(state, svc):
    story = svc.story_source.get_story(state["story_ref"])
    for k, v in (state.get("overrides") or {}).items():
        if v:
            setattr(story, k, v)
    missing = [f for f in ("asset_type", "acronym") if not getattr(story, f)]
    if missing:
        raise WorkflowError(f"Story {story.key} has no {missing}. Add them to the story or enter them manually.")
    svc.emit("node", f"Story {story.key}: asset type '{story.asset_type}', acronym '{story.acronym}'")
    return {"story": asdict(story), "log": [f"loaded story {story.key}"]}


def extract_platinum(state, svc):
    story = _story(state["story"])
    atts = story.pptx_attachments()
    if not atts:
        raise WorkflowError(f"Story {story.key} has no .pptx attachment (Platinum copy)")
    wanted = state.get("attachment_name")
    if wanted:
        att = next((a for a in atts if a.filename == wanted), None)
        if att is None:
            raise WorkflowError(f"Attachment '{wanted}' not found on {story.key}")
    else:
        named = [a for a in atts if "platinum" in a.filename.lower()]
        att = max(named or atts, key=lambda a: a.created or "")
    work = Path(state["work_dir"])
    src = svc.story_source.download_attachment(story, att, work)
    raw = Path(src).read_bytes()

    prs = Presentation(str(src))
    n, idx = len(prs.slides), state.get("slide_index")
    if n > 1 and idx is None:
        raise WorkflowError(f"Platinum copy '{att.filename}' has {n} slides; choose which one to template")
    idx = idx or 0
    if idx >= n:
        raise WorkflowError(f"Slide {idx + 1} requested but '{att.filename}' has {n} slide(s)")
    if n > 1:
        keep_only_slide(prs, idx)                       # templates are exactly one slide
    out = work / "platinum.pptx"
    prs.save(str(out))
    sha = hashlib.sha256(raw + f":slide{idx}".encode()).hexdigest()
    svc.emit("node", f"Platinum copy '{att.filename}' (slide {idx + 1} of {n})")
    return {"platinum_path": str(out), "platinum_sha256": sha,
            "story": {**state["story"], "attachments": [asdict(a) for a in story.attachments]},
            "log": [f"extracted {att.filename} slide {idx + 1}"]}


def resolve_key(state, svc):
    s = state["story"]
    key = build_template_key({"asset_type": s["asset_type"], "acronym": s["acronym"]})
    entry = svc.repos.registry.get(key.key)
    ds_id = s.get("datasource_id") or (entry or {}).get("datasource_id")
    if not ds_id:
        raise WorkflowError(f"No datasource for {key.key}: set it on the story or choose one")
    ds = svc.repos.datasources.get(ds_id)
    base = {"key": key.as_dict(), "registry_entry": entry, "datasource_id": ds_id,
            "required_params": ds.get("params", [])}

    if entry:
        latest = entry["versions"][-1]
        if latest["platinum_sha256"] == state["platinum_sha256"] and not state.get("force"):
            svc.emit("node", f"{key.key}: Platinum copy unchanged since v{latest['version']}, nothing to do")
            return {**base, "mode": "update", "status": "unchanged", "location": latest["location"],
                    "version": latest["version"], "log": ["unchanged platinum copy, skipped"]}
        prev = None
        if svc.repos.templates.exists(latest["location"], "alias.json"):
            prev = json.loads(svc.repos.templates.read(latest["location"], "alias.json"))
        else:   # created before placeholders existed (mapping.json): nothing to reuse, start fresh
            svc.emit("node", f"v{latest['version']} predates alias.json; aliases will be created from scratch")
        svc.emit("node", f"{key.key} exists (v{latest['version']}): updating")
        return {**base, "mode": "update", "previous_doc": prev, "log": [f"update from v{latest['version']}"]}
    svc.emit("node", f"{key.key} is new: creating")
    return {**base, "mode": "create", "previous_doc": None, "log": ["create new template"]}

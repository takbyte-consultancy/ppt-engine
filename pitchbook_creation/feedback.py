"""Constrained feedback agent for approved pitchbook revisions."""
from __future__ import annotations

from core.llm_agent import Tool, run_agent
import config


def interpret_feedback(feedback: list[dict], slides: list[dict], components: list[list[dict]],
                       on_event=None, use_llm: bool = True) -> list[dict]:
    """Turn free-text notes into validated title, order, and component replacement actions."""
    if not feedback:
        return []
    known = {(slide["key"], int(slide["version"])): set(
        int(c["shape_id"]) for c in components[index] if c.get("shape_id") is not None)
        for index, slide in enumerate(slides)}
    actions: list[dict] = []

    def rename_deck(title: str):
        title = str(title).strip()
        if not title or len(title) > 150:
            raise ValueError("Deck title must be 1–150 characters")
        actions.append({"type": "rename", "title": title})
        return {"ok": True, "title": title}

    def move_slide(key: str, version: int, position: int):
        source = (key, int(version))
        if source not in known:
            raise ValueError("Unknown source exhibit/version")
        if position < 1 or position > len(slides):
            raise ValueError("Target position is outside the pitchbook")
        actions.append({"type": "move", "key": key, "version": int(version), "position": position})
        return {"ok": True, "source": key, "position": position}

    def replace_component(key: str, version: int, shape_id: int, replacement_text: str):
        source = (key, int(version))
        if source not in known or int(shape_id) not in known[source]:
            raise ValueError("Unknown component for that exhibit version")
        if len(replacement_text) > 4000:
            raise ValueError("Replacement text is too long")
        actions.append({"type": "replace_text", "key": key, "version": int(version),
                        "shape_id": int(shape_id), "replacement_text": replacement_text})
        return {"ok": True, "source": key, "shape_id": int(shape_id)}

    tools = [
        Tool("rename_pitchbook", "Change the output pitchbook filename/title when explicitly requested.",
             {"type": "object", "properties": {"title": {"type": "string"}}, "required": ["title"]}, rename_deck),
        Tool("move_slide", "Move a source exhibit to a requested 1-based position.",
             {"type": "object", "properties": {"key": {"type": "string"}, "version": {"type": "integer"},
                "position": {"type": "integer"}}, "required": ["key", "version", "position"]}, move_slide),
        Tool("replace_component_text", "Replace editable text for a specific source exhibit shape.",
             {"type": "object", "properties": {"key": {"type": "string"}, "version": {"type": "integer"},
                "shape_id": {"type": "integer"}, "replacement_text": {"type": "string"}},
              "required": ["key", "version", "shape_id", "replacement_text"]}, replace_component),
    ]
    inventory = [{"position": index + 1, "source": {"key": slide["key"], "version": slide["version"]},
                  "components": components[index]} for index, slide in enumerate(slides)]
    for item in feedback:
        action_start = len(actions)
        if item.get("replacement_text") and item.get("shape_id") is not None:
            try:
                replace_component(item["source"]["key"], item["source"]["version"],
                                  item["shape_id"], item["replacement_text"])
            except (KeyError, TypeError, ValueError):
                item["status"] = "unresolved"
                item["resolution_note"] = "The component replacement could not be validated."
        else:
            if not use_llm:
                item["status"] = "open"
                item["resolution_note"] = "AI interpretation was turned off; review this note manually."
                continue
            if not config.OPENAI_API_KEY:
                item["status"] = "open"
                item["resolution_note"] = "Feedback agent is unavailable until OPENAI_API_KEY is configured."
                continue
            run_agent(
                "You are the pitchbook feedback agent. Interpret only the supplied reviewer note. "
                "Use the available tools for explicit deck rename, slide reorder, or exact component text edits. "
                "Do not invent data, change financial values, or make unrelated edits. If the note is not actionable "
                "with these tools, explain that no supported edit was made.",
                f"Reviewer note: {item.get('comment', '')}\nAvailable source slides and components: {inventory}",
                tools, max_steps=4, on_event=on_event)
        for action in actions[action_start:]:
            action["feedback_id"] = item.get("id")
    return actions

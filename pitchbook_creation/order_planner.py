"""Optional LLM ordering for already-resolved, saved exhibit slides."""
from __future__ import annotations

import json

import config
from core.llm_agent import Tool, run_agent


def plan_slide_order(request: str, slides: list[dict], on_event=None) -> tuple[list[dict], str]:
    """Return a validated permutation. The model can only reorder known saved slide UIDs."""
    original = list(slides)
    if not request.strip():
        return original, "No ordering instruction supplied; using the saved exhibit order."
    if not config.OPENAI_API_KEY:
        return original, "AI ordering is unavailable until OPENAI_API_KEY is configured; using the saved exhibit order."

    by_uid = {slide["uid"]: slide for slide in original}
    proposed: list[str] = []

    def set_slide_order(uids: list[str]):
        nonlocal proposed
        if len(uids) != len(by_uid) or len(set(uids)) != len(uids) or set(uids) != set(by_uid):
            raise ValueError("Order must contain every available saved slide UID exactly once")
        proposed = list(uids)
        return {"accepted": True, "ordered_uids": proposed}

    tool = Tool("set_slide_order", "Set the complete order of the available saved exhibit slides.",
        {"type": "object", "properties": {"ordered_uids": {"type": "array", "items": {"type": "string"}}},
         "required": ["ordered_uids"]}, set_slide_order)
    inventory = [{"uid": s["uid"], "exhibit": s["exhibit_name"], "key": s["key"],
                  "version": s["version"], "saved_file": s["file_name"]} for s in original]
    try:
        run_agent(
            "You are a constrained pitchbook ordering assistant. Interpret the user's requested narrative order "
            "and call set_slide_order with every supplied UID exactly once. Never invent, omit, or rename slides. "
            "If the wording is vague, choose the most reasonable narrative order from exhibit names and explain briefly.",
            f"Ordering request: {request}\nAvailable saved slides: {json.dumps(inventory, ensure_ascii=False)}",
            [tool], max_steps=3, on_event=on_event)
        if not proposed:
            return original, "AI could not produce a validated order; using the saved exhibit order."
        return [by_uid[uid] for uid in proposed], "AI proposed an order from the request; review and adjust it before approval."
    except Exception as exc:
        return original, f"AI ordering failed ({exc}); using the saved exhibit order."

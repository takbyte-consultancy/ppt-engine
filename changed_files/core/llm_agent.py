"""
A minimal, transparent agent loop on the OpenAI Responses API (/v1/responses).
Works with reasoning models (which need this API for tool calls) and non-reasoning models.

    goal + tools  ->  LLM picks a tool  ->  we run it  ->  result goes back to LLM  ->  repeat
                                                            until the LLM answers without a tool call
"""
from __future__ import annotations

import json
import traceback
from dataclasses import dataclass
from typing import Any, Callable

from openai import OpenAI
import config


@dataclass
class Tool:
    name: str
    description: str
    parameters: dict              # JSON Schema of the arguments
    func: Callable[..., Any]

    def spec(self) -> dict:       # Responses API tool format (flat, no nested "function" key)
        return {"type": "function", "name": self.name, "description": self.description,
                "parameters": self.parameters, "strict": False}


MAX_TOOL_OUTPUT_CHARS = 200_000


def _serialise(result) -> str:
    """Tool results always reach the model as valid JSON; oversize results are replaced, never cut mid-text."""
    text = json.dumps(result, default=str, ensure_ascii=False)
    if len(text) <= MAX_TOOL_OUTPUT_CHARS:
        return text
    return json.dumps({"error": f"tool output too large ({len(text)} characters, limit {MAX_TOOL_OUTPUT_CHARS}); "
                                "request a narrower result"})


def run_agent(system_prompt: str, user_prompt: str, tools: list[Tool],
              max_steps: int = 25, on_event: Callable[[str, Any], None] | None = None,
              stop_when: Callable[[], bool] | None = None, client: Any = None,
              images: list[bytes] | None = None) -> str:
    """
    on_event(kind, payload) lets the UI stream what the agent is doing.
    stop_when() lets the caller end the loop as soon as a goal state is reached.
    client lets tests inject a fake OpenAI client.
    images (PNG bytes) are attached to the first message for vision-capable models.
    """
    emit = on_event or (lambda k, p: None)
    client = client or OpenAI(api_key=config.OPENAI_API_KEY)
    by_name = {t.name: t for t in tools}
    tool_specs = [t.spec() for t in tools]
    extra = {"reasoning": {"effort": config.OPENAI_REASONING_EFFORT}} \
        if config.OPENAI_REASONING_EFFORT else {}
    content: Any = user_prompt
    if images:
        import base64
        content = [{"type": "input_text", "text": user_prompt}] + [
            {"type": "input_image", "image_url": "data:image/png;base64," + base64.b64encode(b).decode()}
            for b in images]
    history: list = [{"role": "user", "content": content}]

    for step in range(1, max_steps + 1):
        resp = client.responses.create(model=config.OPENAI_MODEL, instructions=system_prompt,
                                       input=history, tools=tool_specs, **extra)
        # Send everything back next turn, including reasoning items, so the model keeps its train of thought
        history += resp.output
        calls = [item for item in resp.output if item.type == "function_call"]

        if not calls:                               # LLM is done
            emit("final", resp.output_text)
            return resp.output_text or ""

        for call in calls:
            name = call.name
            try:
                args = json.loads(call.arguments or "{}")
                emit("tool_call", {"step": step, "tool": name, "args": args})
                result = by_name[name].func(**args)
            except Exception as e:                  # errors go back to the LLM so it can self-correct
                result = {"error": f"{type(e).__name__}: {e}"}
                emit("tool_error", {"tool": name, "error": traceback.format_exc(limit=1)})
            emit("tool_result", {"tool": name, "result": result})
            history.append({"type": "function_call_output", "call_id": call.call_id,
                            "output": _serialise(result)})

        if stop_when and stop_when():
            emit("final", "Goal reached.")
            return "Goal reached."

    raise RuntimeError(f"Agent did not finish within {max_steps} steps")

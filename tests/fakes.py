"""Test doubles: a scripted stand-in for the OpenAI Responses API, and temp repositories."""
from __future__ import annotations

import json
import shutil
from types import SimpleNamespace as NS

import config
from repositories.base import Repositories
from repositories.json_store import FileTemplateRepository, JsonDatasourceRepository, JsonRegistryRepository


class FakeOpenAI:
    """Each step is a list of (tool_name, args) calls, or 'done'. Records every tool result."""

    def __init__(self, steps):
        self.steps, self.results, self._seen = list(steps), [], set()
        self.responses = self

    def create(self, **kw):
        for item in kw["input"]:
            if isinstance(item, dict) and item.get("type") == "function_call_output" \
                    and item["call_id"] not in self._seen:
                self._seen.add(item["call_id"])
                self.results.append(json.loads(item["output"]))
        step = self.steps.pop(0) if self.steps else "done"
        if step == "done":
            return NS(output=[NS(type="message")], output_text="Done.")
        self.n = getattr(self, "n", 0) + 1
        calls = [NS(type="function_call", name=n, arguments=json.dumps(a), call_id=f"call{self.n}_{i}")
                 for i, (n, a) in enumerate(step)]
        return NS(output=calls, output_text="")


def temp_repos(name: str) -> Repositories:
    root = config.ROOT / "tests" / "_out" / name
    shutil.rmtree(root, ignore_errors=True)
    (root / "templates").mkdir(parents=True)
    shutil.copy(config.DATA_DIR / "datasources.json", root / "datasources.json")
    (root / "exhibits").mkdir()
    reg = JsonRegistryRepository(root / "template_registry.json")
    return Repositories(reg, FileTemplateRepository(root / "templates"),
                        JsonDatasourceRepository(reg, root / "datasources.json"),
                        JsonRegistryRepository(root / "exhibit_registry.json"),
                        FileTemplateRepository(root / "exhibits"))

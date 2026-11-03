"""
The template-creation workflow as a LangGraph graph.

    load_story -> extract_platinum -> resolve_key -(unchanged?)-> END
                                           |
                                           v
    extract_metadata -> fetch_data -> map_components <-> validate -> human_review -> build_template -> persist
                                                       (retry loop)   (interrupt)     (fidelity gate)

Only map_components uses an LLM. Every node receives injected Services, so storage, JIRA,
datasources and the model are all replaceable without touching this file.
"""
from __future__ import annotations

from langgraph.checkpoint.memory import InMemorySaver
from langgraph.graph import END, START, StateGraph

from template_creation.nodes.analysis import extract_metadata, fetch_data
from template_creation.nodes.build import build_template, persist
from template_creation.nodes.intake import extract_platinum, load_story, resolve_key
from template_creation.nodes.mapping import map_components, validate
from template_creation.nodes.review import human_review
from template_creation.services import Services
from template_creation.state import TERMINAL, TemplateState, WorkflowError


def _bind(fn, svc: Services):
    def node(state):
        try:
            return fn(state, svc)
        except WorkflowError as e:
            svc.emit("error", str(e))
            return {"status": "failed", "error": str(e), "log": [f"{fn.__name__} failed: {e}"]}
    node.__name__ = fn.__name__
    return node


def _then(next_node: str):
    return lambda s: END if s.get("status") in TERMINAL else next_node


def _after_validate(s):
    if s.get("status") in TERMINAL:
        return END
    return "map_components" if s.get("validation_errors") else "human_review"


def _after_review(s):
    if s.get("status") in TERMINAL:
        return END
    return "human_review" if s.get("review_errors") else "build_template"


def build_graph(svc: Services, checkpointer=None):
    g = StateGraph(TemplateState)
    for fn in (load_story, extract_platinum, resolve_key, extract_metadata, fetch_data,
               map_components, validate, human_review, build_template, persist):
        g.add_node(fn.__name__, _bind(fn, svc))

    g.add_edge(START, "load_story")
    chain = ["load_story", "extract_platinum", "resolve_key", "extract_metadata", "fetch_data", "map_components"]
    for a, b in zip(chain, chain[1:]):
        g.add_conditional_edges(a, _then(b), [b, END])
    g.add_conditional_edges("map_components", _then("validate"), ["validate", END])
    g.add_conditional_edges("validate", _after_validate, ["map_components", "human_review", END])
    g.add_conditional_edges("human_review", _after_review, ["human_review", "build_template", END])
    g.add_conditional_edges("build_template", _then("persist"), ["persist", END])
    g.add_edge("persist", END)
    return g.compile(checkpointer=checkpointer or InMemorySaver())

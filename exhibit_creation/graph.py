"""
The Exhibit-creation workflow (Agent 2) as a LangGraph graph. Deterministic, no LLM.

    load_template -> fetch_data -> check_data -(unchanged / data cannot fill template)-> END
                                        |
                                        v
                                  build_exhibit -> persist
"""
from __future__ import annotations

from langgraph.graph import END, START, StateGraph

from exhibit_creation.nodes import build_exhibit, check_data, fetch_data, load_template, persist
from exhibit_creation.services import ExhibitServices
from exhibit_creation.state import TERMINAL, ExhibitError, ExhibitState


def _bind(fn, svc):
    def node(state):
        try:
            return fn(state, svc)
        except ExhibitError as e:
            svc.emit("error", str(e))
            return {"status": "failed", "error": str(e), "log": [f"{fn.__name__} failed: {e}"]}
    node.__name__ = fn.__name__
    return node


def build_graph(svc: ExhibitServices):
    g = StateGraph(ExhibitState)
    chain = [load_template, fetch_data, check_data, build_exhibit, persist]
    for fn in chain:
        g.add_node(fn.__name__, _bind(fn, svc))
    g.add_edge(START, "load_template")
    for a, b in zip(chain, chain[1:]):
        g.add_conditional_edges(a.__name__, lambda s, nxt=b.__name__: END if s.get("status") in TERMINAL else nxt,
                                [b.__name__, END])
    g.add_edge("persist", END)
    return g.compile()

"""
Mapper interface + the LLM implementation.

Code proposes the aliases (core.alias_engine.propose); the agent reviews and completes them:
resolve ambiguous choices, alias what code could not match (numbers inside text, composed text),
decide callouts, bind charts. Its submit tool runs exactly the same validation as the graph.
"""
from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Protocol

from pptx import Presentation

from core import alias_engine as eng
from core import pptx_inspector as insp
from core.alias_schema import AliasDef, AliasDocument, CalloutDecision, ChartBinding, Location
from core.llm_agent import Tool, run_agent
from template_creation import rules

PROMPT = (Path(__file__).parent / "prompts" / "mapping_agent.md").read_text(encoding="utf-8")


@dataclass
class MappingContext:
    platinum_pptx: str
    sample: dict
    markers: dict
    key: dict
    datasource_id: str
    required_params: list[str]
    draft: dict                                          # code's proposal (AliasDocument as dict)
    unresolved: list[dict] = field(default_factory=list)
    previous: dict | None = None
    feedback: list[str] = field(default_factory=list)
    preview_png: bytes | None = None


class Mapper(Protocol):
    def map(self, ctx: MappingContext, on_event: Callable | None = None) -> dict | None: ...


def _definition(path, format=None, case=None, null="", join=None, sample=None) -> AliasDef:
    return AliasDef(path=path, format=format or None, case=case or None, null=null or "", join=join,
                    description=eng.describe(path), sample=sample)


class LlmMapper:
    def __init__(self, max_steps: int = 40, client=None):
        self.max_steps, self.client = max_steps, client

    def map(self, ctx: MappingContext, on_event=None) -> dict | None:
        prs = Presentation(ctx.platinum_pptx)
        doc = AliasDocument.model_validate(ctx.draft)
        previous = AliasDocument.model_validate(ctx.previous) if ctx.previous else None
        state: dict = {"doc": None}

        def summary():
            return {
                "placeholders": [{"id": o.id, "alias": o.alias, "shape_id": o.shape_id, "shape": o.shape_name,
                                  "location": o.location.model_dump(exclude_none=True), "text": o.text,
                                  "path": doc.aliases[o.alias].path, "format": doc.aliases[o.alias].format,
                                  "case": doc.aliases[o.alias].case, "null": doc.aliases[o.alias].null,
                                  "confidence": o.confidence, "other_candidates": o.candidates, "how": o.how}
                                 for o in doc.occurrences],
                "static_labels": [l.model_dump(exclude_none=True) for l in doc.labels],
                "kept_static": [k.model_dump(exclude_none=True) for k in doc.kept_static],
                "unresolved_text_with_numbers": ctx.unresolved,
                "charts_needing_binding": [{"shape_id": s.shape_id, "name": s.name}
                                           for s in insp.iter_shapes(prs.slides[0].shapes)
                                           if getattr(s, "has_chart", False)
                                           and s.shape_id not in {c.shape_id for c in doc.charts}],
            }

        def t_data():
            return ctx.sample

        def t_inspect():
            out = insp.inspect_slide(prs, 0)
            full, part = set(ctx.markers["fully_locked_shape_ids"]), set(ctx.markers["partially_locked_shape_ids"])
            cids = rules.callout_ids(ctx.markers)
            for c in out["components"]:
                if c["shape_id"] in full:
                    c["locked"] = "entire shape is red bold: never changed"
                elif c["shape_id"] in part:
                    c["locked"] = "some runs are red bold (see get_markers)"
                if c["shape_id"] in cids:
                    c["callout_candidate"] = True
            return out

        def t_markers():
            m = ctx.markers
            return {k: m[k] for k in ("callouts", "locked_runs", "fully_locked_shape_ids",
                                      "partially_locked_shape_ids", "locked_cells")}

        def t_previous():
            return ctx.previous or {"note": "No previous version: this is a new template."}

        def t_proposal():
            return summary()

        def t_set_alias(occurrence_id, path, format=None, case=None, null="", join=None):
            vals = eng.resolve(path, ctx.sample)
            o = eng.set_definition(doc, occurrence_id, _definition(path, format, case, null, join,
                                                                   vals[0] if vals else None), previous)
            return {"ok": True, "alias": o.alias, "path_finds": vals[:3]}

        def t_add_alias(shape_id, paragraph, text, path, cell=None, nth=1, format=None, case=None, null="", join=None):
            vals = eng.resolve(path, ctx.sample)
            o = eng.add_occurrence(doc, prs, int(shape_id), Location(paragraph=paragraph, cell=cell), text,
                                   _definition(path, format, case, null, join, vals[0] if vals else None),
                                   nth=nth, previous=previous)
            preview = eng.value_for(doc.aliases[o.alias], ctx.sample)
            return {"ok": True, "id": o.id, "alias": o.alias, "renders_as": preview,
                    "matches_slide_text": preview == text}

        def t_keep_static(occurrence_id, reason):
            eng.keep_static(doc, occurrence_id, reason, previous)
            return {"ok": True}

        def t_bind_chart(shape_id, categories, series):
            doc.charts = [c for c in doc.charts if c.shape_id != shape_id]
            doc.charts.append(ChartBinding(shape_id=shape_id, categories=categories, series=series))
            return {"ok": True}

        def t_decide_callouts(decisions):
            doc.callouts = [CalloutDecision.model_validate(d) for d in decisions]
            design = {d.shape_id for d in doc.callouts if d.decision == "design_element"}
            added = 0
            if design:            # design elements become normal shapes: let code propose their aliases
                all_ids = {s.shape_id for s in insp.iter_shapes(prs.slides[0].shapes)}
                extra = eng.propose(prs, ctx.sample, ctx.markers, exclude_ids=all_ids - design).doc
                for o in extra.occurrences:
                    if not any(x.shape_id == o.shape_id and x.location == o.location for x in doc.occurrences):
                        d = extra.aliases[o.alias]
                        try:
                            eng.add_occurrence(doc, prs, o.shape_id, o.location, o.text, d,
                                               how=o.how + " (design element)", previous=previous)
                            added += 1
                        except ValueError:
                            pass
            return {"ok": True, "design_element_placeholders_added": added}

        def t_submit():
            rules.enrich(doc, ctx.key, ctx.datasource_id, ctx.required_params)
            errors, warnings = rules.validate_all(doc, prs, ctx.sample, ctx.markers)
            if errors:
                return {"status": "invalid", "errors": errors}
            state["doc"] = doc.model_dump(exclude_none=True)
            return {"status": "ok", "placeholders": len(doc.occurrences), "aliases": len(doc.aliases),
                    "warnings": warnings}

        empty = {"type": "object", "properties": {}}
        fmt_props = {"format": {"type": "string"}, "case": {"type": "string",
                     "enum": ["lower", "upper", "lower_first", "upper_first"]},
                     "null": {"type": "string"}, "join": {"type": "string"}}
        tools = [
            Tool("get_data_source_response", "The datasource JSON for this slide.", empty, t_data),
            Tool("inspect_slide", "Every shape: ids, names, text, run texts, table cells, lock/callout flags.",
                 empty, t_inspect),
            Tool("get_markers", "Callout candidates (yellow/green) and locked (red bold) runs and cells.",
                 empty, t_markers),
            Tool("get_previous_mapping", "The current template version's alias document (updates only).",
                 empty, t_previous),
            Tool("get_alias_proposal", "Code's current placeholders, static labels, unresolved text, charts.",
                 empty, t_proposal),
            Tool("set_alias", "Point an existing placeholder at a different JSON path / format.",
                 {"type": "object", "properties": {"occurrence_id": {"type": "string"}, "path": {"type": "string"},
                                                   **fmt_props}, "required": ["occurrence_id", "path"]},
                 t_set_alias),
            Tool("add_alias", "Turn a text fragment the code did not match into a placeholder.",
                 {"type": "object", "properties": {
                     "shape_id": {"type": "integer"}, "paragraph": {"type": "integer"},
                     "cell": {"type": "array", "items": {"type": "integer"}},
                     "text": {"type": "string", "description": "exact text to replace"},
                     "nth": {"type": "integer", "description": "which occurrence of text in that paragraph"},
                     "path": {"type": "string"}, **fmt_props},
                  "required": ["shape_id", "paragraph", "text", "path"]}, t_add_alias),
            Tool("keep_static", "Remove a placeholder and keep the text literal. A reason is mandatory and is "
                 "shown to reviewers.",
                 {"type": "object", "properties": {"occurrence_id": {"type": "string"}, "reason": {"type": "string"}},
                  "required": ["occurrence_id", "reason"]}, t_keep_static),
            Tool("bind_chart", "Bind a chart's categories and series to JSON arrays.",
                 {"type": "object", "properties": {
                     "shape_id": {"type": "integer"}, "categories": {"type": "string"},
                     "series": {"type": "array", "items": {"type": "object", "properties": {
                         "name": {"type": "string"}, "values": {"type": "string"}}, "required": ["name", "values"]}}},
                  "required": ["shape_id", "categories", "series"]}, t_bind_chart),
            Tool("decide_callouts", "Decide every yellow/green callout candidate.",
                 {"type": "object", "properties": {"decisions": {"type": "array", "items": {
                     "type": "object", "properties": {
                         "shape_id": {"type": "integer"},
                         "decision": {"type": "string", "enum": ["instruction", "design_element"]},
                         "instruction_text": {"type": "string"}, "target_shape_id": {"type": "integer"},
                         "how_applied": {"type": "string"}, "needs_human": {"type": "boolean"}},
                     "required": ["shape_id", "decision"]}}}, "required": ["decisions"]}, t_decide_callouts),
            Tool("submit", "Validate and submit the alias document. Returns errors to fix, or status ok.",
                 empty, t_submit),
        ]
        mode = "UPDATE of an existing template" if ctx.previous else "NEW template"
        user = (f"{mode}. Template key: {ctx.key['key']}. Datasource: {ctx.datasource_id}.\n"
                f"Code proposed {len(doc.occurrences)} placeholder(s); "
                f"{sum(o.confidence == 'low' for o in doc.occurrences)} are ambiguous; "
                f"{len(ctx.unresolved)} text(s) still contain unmatched numbers; "
                f"{len(ctx.markers['callouts'])} callout candidate(s).")
        if ctx.feedback:
            user += "\nThe previous attempt failed validation. Fix these:\n- " + "\n- ".join(ctx.feedback)
        run_agent(PROMPT, user, tools, max_steps=self.max_steps, on_event=on_event,
                  stop_when=lambda: state["doc"] is not None, client=self.client,
                  images=[ctx.preview_png] if ctx.preview_png else None)
        return state["doc"]

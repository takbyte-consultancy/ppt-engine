"""
Template-creation rules on top of the alias engine:

  shape sets       which shapes are excluded (undecided/instruction callouts) or removed from the template
  enrich()         code fills in what the LLM must not decide (key fields, datasource, params)
  validate_all()   callout decisions + alias engine validation (paths, coverage, locks, round trip)
  diff_docs()      what changed versus the previous template version (for review on updates)
"""
from __future__ import annotations

from core import alias_engine as eng
from core.alias_schema import AliasDocument


def callout_ids(markers: dict) -> set[int]:
    return {c["shape_id"] for c in markers["callouts"]}


def removed_shape_ids(doc: AliasDocument, markers: dict) -> set[int]:
    """Instruction callouts and their connectors are deleted from the template."""
    by_id = {c["shape_id"]: c for c in markers["callouts"]}
    out: set[int] = set()
    for d in doc.callouts:
        if d.decision == "instruction" and d.shape_id in by_id:
            out.add(d.shape_id)
            out.update(by_id[d.shape_id].get("connector_ids", []))
    return out


def excluded_ids(doc: AliasDocument, markers: dict) -> set[int]:
    """Shapes that never get placeholders: callouts not (yet) decided as design elements, and connectors."""
    design = {d.shape_id for d in doc.callouts if d.decision == "design_element"}
    connectors = set(markers.get("callout_connector_ids", []))
    return (callout_ids(markers) - design) | connectors


def enrich(doc: AliasDocument, key: dict, datasource_id: str, required_params: list[str]) -> AliasDocument:
    doc.template_id = key["template_id"]
    doc.template_key = key["key"]
    doc.key_fields = key["fields"]
    doc.datasource_id = datasource_id
    doc.required_params = list(required_params)
    return doc


def _callout_checks(doc: AliasDocument, markers: dict, shape_ids: set[int]) -> tuple[list[str], list[str]]:
    errors, warnings = [], []
    candidates = {c["shape_id"]: c for c in markers["callouts"]}
    decided = {d.shape_id: d for d in doc.callouts}
    for cid, c in candidates.items():
        if cid not in decided:
            errors.append(f"Callout id {cid} ({c['colour']}: \"{c['text'][:60]}\") has no decision (decide_callouts)")
    for cid in set(decided) - set(candidates):
        errors.append(f"callout decision for shape id {cid}, which is not a callout candidate")
    for d in doc.callouts:
        if d.decision != "instruction":
            continue
        if d.needs_human:
            warnings.append(f"Callout {d.shape_id} needs a human: \"{d.instruction_text[:80]}\"")
        elif d.target_shape_id is None:
            errors.append(f"Instruction callout {d.shape_id} needs target_shape_id (or needs_human=true)")
        elif d.target_shape_id not in shape_ids:
            errors.append(f"Callout {d.shape_id}: target_shape_id {d.target_shape_id} does not exist")
    return errors, warnings


def validate_all(doc: AliasDocument, prs, data, markers: dict) -> tuple[list[str], list[str]]:
    from core.pptx_inspector import iter_shapes
    shape_ids = {s.shape_id for s in iter_shapes(prs.slides[0].shapes)}
    e1, w1 = _callout_checks(doc, markers, shape_ids)
    e2, w2 = eng.validate(doc, prs, data, markers, excluded_ids(doc, markers), removed_shape_ids(doc, markers))
    return e1 + e2, w1 + w2


def diff_docs(previous: dict | None, current: dict) -> dict:
    """Compare by JSON path: which data fields were added, removed, or are now formatted differently."""
    cur = {a["path"]: (n, a) for n, a in current.get("aliases", {}).items()}
    if not previous:
        return {"added": sorted(cur), "removed": [], "changed": []}
    prev = {a["path"]: (n, a) for n, a in previous.get("aliases", {}).items()}
    changed = []
    for p in prev.keys() & cur.keys():
        (pn, pa), (cn, ca) = prev[p], cur[p]
        fields = [f for f in ("format", "case", "null", "join") if pa.get(f) != ca.get(f)]
        if pn != cn:
            fields.append("alias")
        if fields:
            changed.append({"path": p, "alias": cn, "fields": fields})
    return {"added": sorted(cur.keys() - prev.keys()), "removed": sorted(prev.keys() - cur.keys()),
            "changed": changed}

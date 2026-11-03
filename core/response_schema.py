"""
Records the structure of an API response when a template is created, and checks later
responses against it, so a renamed or missing field fails loudly instead of leaving a blank box.
"""
from __future__ import annotations


def infer_schema(v):
    if isinstance(v, dict):
        return {"type": "object", "properties": {k: infer_schema(x) for k, x in v.items()}}
    if isinstance(v, list):
        return {"type": "array", "items": infer_schema(v[0]) if v else {}}
    if isinstance(v, bool):
        return {"type": "boolean"}
    if isinstance(v, (int, float)):
        return {"type": "number"}
    if v is None:
        return {"type": "null"}
    return {"type": "string"}


def check_against_schema(data, schema, path: str = "$") -> list[str]:
    errs: list[str] = []
    t = schema.get("type")
    if not t:
        return errs
    if data is None and t != "null":
        return [f"{path}: is null (expected {t})"]
    if t == "object":
        if not isinstance(data, dict):
            return [f"{path}: expected object"]
        for k, sub in schema["properties"].items():
            if k not in data:
                errs.append(f"{path}.{k}: missing")
            else:
                errs += check_against_schema(data[k], sub, f"{path}.{k}")
    elif t == "array":
        if not isinstance(data, list):
            return [f"{path}: expected array"]
        if data and schema.get("items"):
            errs += check_against_schema(data[0], schema["items"], f"{path}[0]")
    elif t == "number" and not isinstance(data, (int, float)):
        errs.append(f"{path}: expected number, got {type(data).__name__}")
    elif t == "string" and not isinstance(data, str):
        errs.append(f"{path}: expected string, got {type(data).__name__}")
    return errs

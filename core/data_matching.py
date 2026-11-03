"""
Finds, deterministically, which shapes' content already exists in the datasource response.
This is evidence for the validator: a shape whose content is in the data must be mapped, not static.

Evidence strength:
  strong  - exact (whitespace/case-insensitive) or near-exact text match, the JSON names the shape
            via "shape_name", or a 2-D array that matches a table's size and cell text
  weak    - the data value appears inside the shape's text (e.g. "As of 30 Jun 2026" contains a date),
            or an array has the table's size but different text
"""
from __future__ import annotations

import difflib
import re

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

from core.pptx_inspector import iter_shapes


def norm(text) -> str:
    return re.sub(r"\s+", " ", str(text).replace("\x0b", " ")).strip().lower()


def _child_path(parent: str, container: list, index: int) -> str:
    """Prefer a filter by shape_name (stable) over an index (fragile) when the JSON offers one."""
    item = container[index]
    if isinstance(item, dict) and isinstance(item.get("shape_name"), str):
        name = item["shape_name"].replace("'", "\\'")
        return f"{parent}[?(@.shape_name=='{name}')]"
    return f"{parent}[{index}]"


def walk(data, path: str = "$"):
    """Yield (path, value, kind) for string leaves, 2-D arrays and shape_name references."""
    if isinstance(data, dict):
        for k, v in data.items():
            yield from walk(v, f"{path}.{k}")
    elif isinstance(data, list):
        if data and all(isinstance(r, list) for r in data):
            yield path, data, "grid"
        for i, item in enumerate(data):
            child = _child_path(path, data, i)
            if isinstance(item, dict) and isinstance(item.get("shape_name"), str):
                yield child, item, "shape_ref"
            yield from walk(item, child)
    elif isinstance(data, str):
        yield path, data, "text"


def _has_letters(s: str, n: int = 3) -> bool:
    return len(re.findall(r"[a-zA-Z]", s)) >= n


def match_slide_to_data(prs: Presentation, data: dict, original_names: dict[int, str] | None = None) -> dict:
    """{shape_id (str): [{"path", "strength", "how"}...]} for every shape with evidence."""
    original_names = original_names or {}
    leaves = list(walk(data))
    texts = [(p, v, norm(v)) for p, v, k in leaves if k == "text" and _has_letters(v)]
    grids = [(p, v) for p, v, k in leaves if k == "grid"]
    refs = {}
    for p, v, k in leaves:
        if k == "shape_ref":
            refs.setdefault(v["shape_name"], []).append(p)

    shapes = [s for s in iter_shapes(prs.slides[0].shapes) if s.shape_type != MSO_SHAPE_TYPE.GROUP]
    name_count: dict[str, int] = {}
    for s in shapes:
        n = original_names.get(s.shape_id, s.name)
        name_count[n] = name_count.get(n, 0) + 1

    out: dict[str, list] = {}
    for s in shapes:
        ev = []
        name = original_names.get(s.shape_id, s.name)
        for p in refs.get(name, []):
            if name_count[name] == 1:
                ev.append({"path": p, "strength": "strong", "how": f"data refers to shape '{name}' by shape_name"})
            else:   # several shapes share this name: only a text match can tell which one is meant
                ev.append({"path": p, "strength": "weak",
                           "how": f"data refers to '{name}', but {name_count[name]} shapes share that name"})

        if getattr(s, "has_table", False) and s.has_table:
            cells = [[norm(c.text) for c in row.cells] for row in s.table.rows]
            n_rows, n_cols = len(cells), len(s.table.columns)
            for p, g in grids:
                if len(g) != n_rows or max((len(r) for r in g), default=0) != n_cols:
                    continue
                filled = [(r, c) for r in range(n_rows) for c in range(n_cols) if cells[r][c]]
                same = sum(1 for r, c in filled if c < len(g[r]) and norm(g[r][c]) == cells[r][c])
                ratio = same / len(filled) if filled else 0
                strength = "strong" if ratio >= 0.3 else "weak"
                ev.append({"path": p, "strength": strength,
                           "how": f"2-D array with the table's size ({n_rows}x{n_cols}), "
                                  f"{same}/{len(filled)} cells identical -> use table_mode 'grid'"})

        elif getattr(s, "has_text_frame", False) and s.has_text_frame and _has_letters(s.text_frame.text):
            st = norm(s.text_frame.text)
            for p, raw, nv in texts:
                if nv == st:
                    ev.append({"path": p, "strength": "strong", "how": "identical text"})
                elif len(st) > 40 and len(nv) > 40 and difflib.SequenceMatcher(None, st, nv).ratio() >= 0.9:
                    ev.append({"path": p, "strength": "strong", "how": "near-identical text (>=90%)"})
                elif len(nv) >= 8 and nv in st:
                    ev.append({"path": p, "strength": "weak", "how": "data value appears inside the shape text"})
        if ev:
            seen, uniq = set(), []
            for e in sorted(ev, key=lambda e: e["strength"] != "strong"):
                if e["path"] not in seen:
                    seen.add(e["path"]); uniq.append(e)
            out[str(s.shape_id)] = uniq
    return out

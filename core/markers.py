"""
Deterministic detection of the Platinum copy's colour conventions:

  * Yellow / green filled shapes  -> CALLOUT candidates (instructions for the template author).
    The LLM decides whether each is really an instruction or a design element, and what it says.
  * Red + bold text runs          -> LOCKED text: kept exactly as-is (text and red bold formatting)
                                     in the template and in every exhibit.
  * Connectors touching a callout -> removed together with the callout.
"""
from __future__ import annotations

import colorsys

from lxml import etree
from pptx import Presentation
from pptx.enum.dml import MSO_COLOR_TYPE, MSO_FILL, MSO_THEME_COLOR
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.opc.constants import RELATIONSHIP_TYPE as RT

from core.pptx_inspector import iter_shapes

A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"

SCHEME_NAMES = {
    MSO_THEME_COLOR.ACCENT_1: "accent1", MSO_THEME_COLOR.ACCENT_2: "accent2",
    MSO_THEME_COLOR.ACCENT_3: "accent3", MSO_THEME_COLOR.ACCENT_4: "accent4",
    MSO_THEME_COLOR.ACCENT_5: "accent5", MSO_THEME_COLOR.ACCENT_6: "accent6",
    MSO_THEME_COLOR.DARK_1: "dk1", MSO_THEME_COLOR.LIGHT_1: "lt1",
    MSO_THEME_COLOR.DARK_2: "dk2", MSO_THEME_COLOR.LIGHT_2: "lt2",
    MSO_THEME_COLOR.TEXT_1: "dk1", MSO_THEME_COLOR.BACKGROUND_1: "lt1",
    MSO_THEME_COLOR.TEXT_2: "dk2", MSO_THEME_COLOR.BACKGROUND_2: "lt2",
    MSO_THEME_COLOR.HYPERLINK: "hlink", MSO_THEME_COLOR.FOLLOWED_HYPERLINK: "folHlink",
}


# --------------------------------------------------------------------------- #
# Colours
# --------------------------------------------------------------------------- #
def theme_colors(prs: Presentation) -> dict[str, str]:
    """Scheme colour name -> 'RRGGBB' from the first slide master's theme."""
    try:
        theme = prs.slide_masters[0].part.part_related_by(RT.THEME)
        root = etree.fromstring(theme.blob)
    except Exception:
        return {}
    out = {}
    scheme = root.find(f".//{A}clrScheme")
    for child in (scheme if scheme is not None else []):
        name = etree.QName(child).localname
        srgb, sysc = child.find(f"{A}srgbClr"), child.find(f"{A}sysClr")
        if srgb is not None:
            out[name] = srgb.get("val").upper()
        elif sysc is not None and sysc.get("lastClr"):
            out[name] = sysc.get("lastClr").upper()
    return out


def resolve_color(color_format, theme: dict) -> str | None:
    """'RRGGBB' for an RGB or theme colour, None if unset/inherited."""
    try:
        t = color_format.type
        if t == MSO_COLOR_TYPE.RGB:
            return str(color_format.rgb).upper()
        if t == MSO_COLOR_TYPE.SCHEME:
            return theme.get(SCHEME_NAMES.get(color_format.theme_color, ""))
    except Exception:
        pass
    return None


def classify(hex_rgb: str | None) -> str:
    """'red' | 'yellow' | 'green' | 'other' (hue based, so slightly different shades still match)."""
    if not hex_rgb or len(hex_rgb) != 6:
        return "other"
    r, g, b = (int(hex_rgb[i:i + 2], 16) / 255 for i in (0, 2, 4))
    h, s, v = colorsys.rgb_to_hsv(r, g, b)
    h *= 360
    if (h <= 15 or h >= 345) and s >= 0.55 and v >= 0.45:
        return "red"
    if 40 <= h <= 70 and s >= 0.35 and v >= 0.6:
        return "yellow"
    if 75 <= h <= 170 and s >= 0.3 and v >= 0.35:
        return "green"
    return "other"


def shape_fill_hex(shape, theme: dict) -> str | None:
    try:
        if shape.fill.type == MSO_FILL.SOLID:
            return resolve_color(shape.fill.fore_color, theme)
    except Exception:
        pass
    return None


# --------------------------------------------------------------------------- #
# Locked (red + bold) runs
# --------------------------------------------------------------------------- #
def _is_locked_run(run, theme) -> bool:
    try:
        return bool(run.font.bold) and classify(resolve_color(run.font.color, theme)) == "red"
    except Exception:
        return False


def _scan_text_frame(tf, theme, base: dict) -> list[dict]:
    hits = []
    for pi, p in enumerate(tf.paragraphs):
        for ri, r in enumerate(p.runs):
            if r.text.strip() and _is_locked_run(r, theme):
                hits.append({**base, "paragraph": pi, "run": ri, "text": r.text})
    return hits


def _non_empty_runs(tf) -> int:
    return sum(1 for p in tf.paragraphs for r in p.runs if r.text.strip())


# --------------------------------------------------------------------------- #
# Geometry helpers for callout -> target association
# --------------------------------------------------------------------------- #
def _box(s):
    return (s.left or 0, s.top or 0, (s.left or 0) + (s.width or 0), (s.top or 0) + (s.height or 0))


def _overlap(a, b) -> int:
    w = min(a[2], b[2]) - max(a[0], b[0])
    h = min(a[3], b[3]) - max(a[1], b[1])
    return w * h if w > 0 and h > 0 else 0


def _inside(pt, box, margin=114300):          # 0.125 inch tolerance
    return box[0] - margin <= pt[0] <= box[2] + margin and box[1] - margin <= pt[1] <= box[3] + margin


def _is_connector(s) -> bool:
    return s.shape_type in (MSO_SHAPE_TYPE.LINE,) or s.__class__.__name__ == "Connector"


# --------------------------------------------------------------------------- #
# Public API
# --------------------------------------------------------------------------- #
def detect_markers(prs: Presentation, slide_index: int = 0) -> dict:
    theme = theme_colors(prs)
    slide = prs.slides[slide_index]
    shapes = [s for s in iter_shapes(slide.shapes) if s.shape_type != MSO_SHAPE_TYPE.GROUP]
    slide_area = prs.slide_width * prs.slide_height

    # 1) callout candidates: filled yellow/green shapes that carry text (charts/tables/pictures excluded)
    callouts = []
    for s in shapes:
        if getattr(s, "has_chart", False) or getattr(s, "has_table", False) \
                or s.shape_type == MSO_SHAPE_TYPE.PICTURE or _is_connector(s):
            continue
        colour = classify(shape_fill_hex(s, theme))
        text = s.text_frame.text.strip() if getattr(s, "has_text_frame", False) and s.has_text_frame else ""
        if colour in ("yellow", "green") and text:
            callouts.append({"shape_id": s.shape_id, "original_name": s.name, "colour": colour,
                             "text": text, "box": _box(s)})
    callout_ids = {c["shape_id"] for c in callouts}

    # 2) connectors touching callouts, and target candidates for each callout
    targets_pool = [s for s in shapes if s.shape_id not in callout_ids and not _is_connector(s)
                    and (s.width or 0) * (s.height or 0) < 0.8 * slide_area]
    connector_ids: set[int] = set()
    for c in callouts:
        cands, c["connector_ids"] = [], []
        for con in (s for s in shapes if _is_connector(s)):
            try:
                ends = [(con.begin_x, con.begin_y), (con.end_x, con.end_y)]
            except Exception:
                b = _box(con)
                ends = [(b[0], b[1]), (b[2], b[3])]
            if any(_inside(e, c["box"]) for e in ends):
                connector_ids.add(con.shape_id)
                c["connector_ids"].append(con.shape_id)
                far = ends[1] if _inside(ends[0], c["box"]) else ends[0]
                for t in targets_pool:
                    if _inside(far, _box(t), margin=0):
                        cands.append({"shape_id": t.shape_id, "name": t.name, "method": "connector"})
        for t in sorted(targets_pool, key=lambda t: -_overlap(c["box"], _box(t))):
            if _overlap(c["box"], _box(t)) > 0:
                cands.append({"shape_id": t.shape_id, "name": t.name, "method": "overlap"})
        cx, cy = (c["box"][0] + c["box"][2]) / 2, (c["box"][1] + c["box"][3]) / 2
        for t in sorted(targets_pool, key=lambda t: ((_box(t)[0] + _box(t)[2]) / 2 - cx) ** 2
                        + ((_box(t)[1] + _box(t)[3]) / 2 - cy) ** 2)[:2]:
            cands.append({"shape_id": t.shape_id, "name": t.name, "method": "nearest"})
        seen, uniq = set(), []
        for x in cands:
            if x["shape_id"] not in seen:
                seen.add(x["shape_id"]); uniq.append(x)
        c["target_candidates"] = uniq[:4]
        del c["box"]

    # 3) locked runs (text shapes and table cells)
    locked_runs, fully_locked_shapes, locked_cells = [], [], {}
    for s in shapes:
        if s.shape_id in callout_ids:
            continue
        if getattr(s, "has_table", False) and s.has_table:
            for r_i, row in enumerate(s.table.rows):
                for c_i, cell in enumerate(row.cells):
                    hits = _scan_text_frame(cell.text_frame, theme,
                                            {"shape_id": s.shape_id, "original_name": s.name, "cell": [r_i, c_i]})
                    if hits:                 # any locked text in a cell locks the whole cell
                        locked_runs += hits
                        locked_cells.setdefault(s.shape_id, []).append([r_i, c_i])
        elif getattr(s, "has_text_frame", False) and s.has_text_frame:
            hits = _scan_text_frame(s.text_frame, theme, {"shape_id": s.shape_id, "original_name": s.name})
            locked_runs += hits
            if hits and len(hits) == _non_empty_runs(s.text_frame):
                fully_locked_shapes.append(s.shape_id)

    partially_locked = sorted({h["shape_id"] for h in locked_runs if "cell" not in h}
                              - set(fully_locked_shapes))
    return {
        "callouts": callouts,
        "callout_connector_ids": sorted(connector_ids),
        "locked_runs": locked_runs,
        "fully_locked_shape_ids": fully_locked_shapes,
        "partially_locked_shape_ids": partially_locked,
        "locked_cells": {str(k): v for k, v in locked_cells.items()},
        "theme_colors": theme,
    }

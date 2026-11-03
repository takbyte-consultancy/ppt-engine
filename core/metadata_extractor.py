"""
Complete, deterministic snapshot of a slide: locators, geometry, fills, lines, text-frame settings,
paragraph + run formatting, table structure + cell formatting, chart formatting, image hashes.

Saved as metadata.json next to every template. The fidelity checker compares two snapshots to
prove formatting survived templating and filling.
"""
from __future__ import annotations

import hashlib

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Emu

from core.markers import resolve_color, shape_fill_hex, theme_colors

A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
SCHEMA_VERSION = 1


def _safe(fn, default=None):
    try:
        return fn()
    except Exception:
        return default


def _enum(v):
    return None if v is None else str(v).split(".")[-1].split(" ")[0]


def _length(v):
    return None if v is None else int(v)


def _font(font, theme) -> dict:
    return {
        "name": _safe(lambda: font.name),
        "size_pt": _safe(lambda: font.size.pt if font.size else None),
        "bold": _safe(lambda: font.bold),
        "italic": _safe(lambda: font.italic),
        "underline": _safe(lambda: font.underline if isinstance(font.underline, (bool, type(None)))
                           else _enum(font.underline)),
        "color": _safe(lambda: resolve_color(font.color, theme)),
        "color_theme": _safe(lambda: _enum(font.color.theme_color)
                             if _enum(font.color.type) == "SCHEME" else None),
    }


def _paragraphs(tf, theme) -> list[dict]:
    out = []
    for p in tf.paragraphs:
        ppr = p._p.find(f"{A}pPr")
        bullet = None
        if ppr is not None:
            if ppr.find(f"{A}buChar") is not None:
                bullet = "char:" + ppr.find(f"{A}buChar").get("char", "")
            elif ppr.find(f"{A}buAutoNum") is not None:
                bullet = "autonum:" + ppr.find(f"{A}buAutoNum").get("type", "")
            elif ppr.find(f"{A}buNone") is not None:
                bullet = "none"
        end = p._p.find(f"{A}endParaRPr")
        out.append({
            "level": p.level,
            "alignment": _enum(_safe(lambda: p.alignment)),
            "space_before_pt": _safe(lambda: p.space_before.pt if p.space_before is not None else None),
            "space_after_pt": _safe(lambda: p.space_after.pt if p.space_after is not None else None),
            "line_spacing": _safe(lambda: p.line_spacing if not hasattr(p.line_spacing, "pt")
                                  else f"{p.line_spacing.pt}pt"),
            "bullet": bullet,
            "margin_left": None if ppr is None else ppr.get("marL"),
            "indent": None if ppr is None else ppr.get("indent"),
            "end_para_size": None if end is None else end.get("sz"),
            "runs": [{"text": r.text, "font": _font(r.font, theme)} for r in p.runs],
        })
    return out


def _text_frame(tf) -> dict:
    return {
        "word_wrap": _safe(lambda: tf.word_wrap),
        "auto_size": _enum(_safe(lambda: tf.auto_size)),
        "vertical_anchor": _enum(_safe(lambda: tf.vertical_anchor)),
        "margins": [_length(_safe(lambda: getattr(tf, f"margin_{m}"))) for m in ("left", "top", "right", "bottom")],
    }


def _cell_borders(cell) -> dict:
    tcpr = cell._tc.find(f"{A}tcPr")
    out = {}
    if tcpr is None:
        return out
    for side in ("lnL", "lnR", "lnT", "lnB"):
        ln = tcpr.find(f"{A}{side}")
        if ln is not None:
            clr = ln.find(f".//{A}srgbClr")
            scheme = ln.find(f".//{A}schemeClr")
            nofill = ln.find(f"{A}noFill") is not None
            out[side] = {"w": ln.get("w"), "color": clr.get("val") if clr is not None else
                         (f"scheme:{scheme.get('val')}" if scheme is not None else None), "none": nofill}
    return out


def _table(shape, theme) -> dict:
    tbl = shape.table
    cells = []
    for r, row in enumerate(tbl.rows):
        for c, cell in enumerate(row.cells):
            cells.append({
                "r": r, "c": c,
                "merge_origin": cell.is_merge_origin, "spanned": cell.is_spanned,
                "span": [cell.span_height, cell.span_width] if cell.is_merge_origin else None,
                "fill": _safe(lambda: resolve_color(cell.fill.fore_color, theme)
                              if _enum(cell.fill.type) == "SOLID" else _enum(cell.fill.type)),
                "margins": [_length(_safe(lambda: getattr(cell, f"margin_{m}"))) for m in
                            ("left", "top", "right", "bottom")],
                "vertical_anchor": _enum(_safe(lambda: cell.vertical_anchor)),
                "borders": _cell_borders(cell),
                "paragraphs": _paragraphs(cell.text_frame, theme),
            })
    tblpr = tbl._tbl.find(f"{A}tblPr")
    return {
        "n_rows": len(tbl.rows), "n_cols": len(tbl.columns),
        "column_widths": [int(c.width) for c in tbl.columns],
        "row_heights": [int(r.height) for r in tbl.rows],
        "flags": {} if tblpr is None else dict(tblpr.attrib),
        "cells": cells,
    }


def _chart(shape, theme) -> dict:
    ch = shape.chart
    plots = []
    for plot in ch.plots:
        series = []
        for s in plot.series:
            series.append({
                "name": s.name,
                "n_points": len(list(s.values)),
                "fill": _safe(lambda: resolve_color(s.format.fill.fore_color, theme)
                              if _enum(s.format.fill.type) == "SOLID" else None),
                "line": _safe(lambda: resolve_color(s.format.line.color, theme)),
            })
        plots.append({
            "has_data_labels": _safe(lambda: plot.has_data_labels),
            "data_label_number_format": _safe(lambda: plot.data_labels.number_format
                                              if plot.has_data_labels else None),
            "n_categories": len(plot.categories),
            "series": series,
        })
    va = _safe(lambda: ch.value_axis)
    return {
        "chart_type": _enum(ch.chart_type),
        "has_legend": ch.has_legend,
        "legend_position": _enum(_safe(lambda: ch.legend.position)) if ch.has_legend else None,
        "has_title": ch.has_title,
        "title": _safe(lambda: ch.chart_title.text_frame.text) if ch.has_title else None,
        "value_axis": None if va is None else {
            "gridlines": _safe(lambda: va.has_major_gridlines),
            "number_format": _safe(lambda: va.tick_labels.number_format),
            "font_size_pt": _safe(lambda: va.tick_labels.font.size.pt if va.tick_labels.font.size else None),
        },
        "plots": plots,
    }


def _shape(s, z: int, theme) -> dict:
    d = {
        "shape_id": s.shape_id,
        "name": s.name,
        "shape_type": _enum(s.shape_type),
        "auto_shape_type": _safe(lambda: _enum(s.auto_shape_type)),
        "z_order": z,
        "locator": {
            "shape_id": s.shape_id, "name": s.name, "z_order": z,
            "placeholder_idx": _safe(lambda: s.placeholder_format.idx) if s.is_placeholder else None,
            "placeholder_type": _enum(_safe(lambda: s.placeholder_format.type)) if s.is_placeholder else None,
            "xpath": s._element.getroottree().getpath(s._element),
        },
        "geometry_emu": {"left": _length(s.left), "top": _length(s.top),
                         "width": _length(s.width), "height": _length(s.height),
                         "rotation": _safe(lambda: s.rotation)},
        "geometry_in": {k: round(Emu(v or 0).inches, 3) for k, v in
                        (("left", s.left), ("top", s.top), ("width", s.width), ("height", s.height))},
        "fill": _safe(lambda: shape_fill_hex(s, theme)),
        "line": _safe(lambda: {"color": resolve_color(s.line.color, theme),
                               "width": _length(s.line.width)}) if hasattr(s, "line") else None,
    }
    if getattr(s, "has_chart", False):
        d["kind"], d["chart"] = "chart", _chart(s, theme)
    elif getattr(s, "has_table", False) and s.has_table:
        d["kind"], d["table"] = "table", _table(s, theme)
    elif s.shape_type == MSO_SHAPE_TYPE.PICTURE:
        d["kind"] = "picture"
        d["image"] = {"sha1": _safe(lambda: hashlib.sha1(s.image.blob).hexdigest()),
                      "ext": _safe(lambda: s.image.ext),
                      "crop": [_safe(lambda: getattr(s, f"crop_{m}")) for m in ("left", "top", "right", "bottom")]}
    elif s.shape_type == MSO_SHAPE_TYPE.GROUP:
        d["kind"] = "group"
        d["children"] = [c.shape_id for c in s.shapes]
    elif getattr(s, "has_text_frame", False) and s.has_text_frame:
        d["kind"] = "text"
    else:
        d["kind"] = "other"
    if getattr(s, "has_text_frame", False) and s.has_text_frame:
        d["text_frame"] = _text_frame(s.text_frame)
        d["paragraphs"] = _paragraphs(s.text_frame, theme)
    return d


def extract_metadata(prs: Presentation, slide_index: int = 0) -> dict:
    theme = theme_colors(prs)
    slide = prs.slides[slide_index]
    shapes, z = [], 0

    def walk(coll):
        nonlocal z
        for s in coll:
            shapes.append(_shape(s, z, theme)); z += 1
            if s.shape_type == MSO_SHAPE_TYPE.GROUP:
                walk(s.shapes)

    walk(slide.shapes)
    return {
        "schema_version": SCHEMA_VERSION,
        "slide": {"width_emu": int(prs.slide_width), "height_emu": int(prs.slide_height),
                  "layout": slide.slide_layout.name, "master": _safe(lambda: slide.slide_layout.slide_master.name)},
        "theme_colors": theme,
        "shapes": shapes,
    }

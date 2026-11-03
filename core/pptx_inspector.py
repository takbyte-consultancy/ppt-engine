"""
Deterministic PPTX utilities (no LLM here).

inspect_slide()   -> describes every component on a slide so the LLM can "see" it
iter_shapes()     -> walks shapes including those inside groups
keep_only_slide() -> trims a multi-slide Platinum copy down to the one exhibit slide
"""
from __future__ import annotations

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Emu


def iter_shapes(shapes):
    for shp in shapes:
        yield shp
        if shp.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from iter_shapes(shp.shapes)


def _inches(v) -> float:
    return round(Emu(v or 0).inches, 2)


def _describe(shp) -> dict:
    d = {
        "shape_id": shp.shape_id,
        "name": shp.name,
        "position_in": {"left": _inches(shp.left), "top": _inches(shp.top),
                        "width": _inches(shp.width), "height": _inches(shp.height)},
    }
    if getattr(shp, "has_chart", False):
        ch = shp.chart
        plot = ch.plots[0]
        d["kind"] = "chart"
        d["chart_type"] = str(ch.chart_type).split(".")[-1].split(" ")[0]
        d["categories"] = [str(c) for c in plot.categories]
        d["series"] = [{"name": s.name, "values": list(s.values)} for s in plot.series]
    elif getattr(shp, "has_table", False) and shp.has_table:
        tbl = shp.table
        rows = [[cell.text for cell in row.cells] for row in tbl.rows]
        d["kind"] = "table"
        d["n_rows"], d["n_cols"] = len(rows), len(tbl.columns)
        d["cells"] = [[t[:60] for t in r] for r in rows]   # full grid so the LLM can compare to JSON
    elif shp.shape_type == MSO_SHAPE_TYPE.PICTURE:
        d["kind"] = "picture"
    elif shp.shape_type == MSO_SHAPE_TYPE.GROUP:
        d["kind"] = "group"
        d["children"] = [c.name for c in shp.shapes]
    elif getattr(shp, "has_text_frame", False) and shp.has_text_frame and shp.text_frame.text.strip():
        d["kind"] = "text"
        d["text"] = shp.text_frame.text
        paras = shp.text_frame.paragraphs
        if any(len(p.runs) > 1 for p in paras):        # run indexes, for run-level mapping
            d["paragraph_runs"] = [[r.text for r in p.runs] for p in paras]
    else:
        d["kind"] = "decoration"
    return d


def inspect_slide(prs: Presentation, slide_index: int = 0) -> dict:
    slide = prs.slides[slide_index]
    comps = [_describe(s) for s in iter_shapes(slide.shapes)]
    return {
        "slide_size_in": {"width": _inches(prs.slide_width), "height": _inches(prs.slide_height)},
        "n_components": len(comps),
        "components": comps,
    }


def mappable_shape_names(prs: Presentation, slide_index: int = 0) -> set[str]:
    """Every shape that must appear in mapping.json (groups are containers, so skipped)."""
    return {s.name for s in iter_shapes(prs.slides[slide_index].shapes)
            if s.shape_type != MSO_SHAPE_TYPE.GROUP}


def rename_shape(prs: Presentation, shape_id: int, new_name: str, slide_index: int = 0) -> str:
    slide = prs.slides[slide_index]
    taken = {s.name for s in iter_shapes(slide.shapes) if s.shape_id != shape_id}
    if new_name in taken:
        raise ValueError(f"name '{new_name}' already used on this slide")
    for s in iter_shapes(slide.shapes):
        if s.shape_id == shape_id:
            old, s.name = s.name, new_name
            return old
    raise ValueError(f"no shape with shape_id={shape_id}")


def keep_only_slide(prs: Presentation, slide_index: int) -> None:
    """Delete every slide except slide_index (python-pptx has no public delete API)."""
    sld_ids = prs.slides._sldIdLst
    for i, sld_id in reversed(list(enumerate(list(sld_ids)))):
        if i != slide_index:
            prs.part.drop_rel(sld_id.rId)
            sld_ids.remove(sld_id)


def table_dims(prs: Presentation, slide_index: int = 0) -> dict[str, tuple[int, int]]:
    return {s.name: (len(s.table.rows), len(s.table.columns))
            for s in iter_shapes(prs.slides[slide_index].shapes)
            if getattr(s, "has_table", False) and s.has_table}

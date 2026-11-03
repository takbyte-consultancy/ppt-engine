"""
Format-preserving write helpers.

Golden rule: never create new formatting. Text is written into runs that already exist, so fonts,
colours, spacing and bullets of the Platinum copy survive templating (Agent 1) and filling (Agent 2).
"""
from __future__ import annotations

from core.pptx_inspector import iter_shapes


def replace_span(paragraph, start: int, end: int, new_text: str) -> None:
    """
    Replace characters [start, end) of a paragraph's text with new_text.
    The new text goes into the run where the span starts (keeping that run's formatting); the rest of
    the span is trimmed from following runs. A placeholder is therefore never split across runs.
    """
    runs = paragraph.runs
    pos, first = 0, None
    for i, r in enumerate(runs):
        r_start, r_end = pos, pos + len(r.text)
        pos = r_end
        if r_end <= start or r_start >= end:
            if r_start == start == end and first is None:        # insertion at a run boundary
                r.text = new_text + r.text
                return
            continue
        a, b = max(start, r_start) - r_start, min(end, r_end) - r_start
        if first is None:
            first = i
            r.text = r.text[:a] + new_text + r.text[b:]
        else:
            r.text = r.text[:a] + r.text[b:]
    if first is None:                                            # empty paragraph or span at the end
        if runs:
            runs[-1].text = runs[-1].text + new_text
        else:
            paragraph.add_run().text = new_text


def remove_shapes(prs, shape_ids: set[int], slide_index: int = 0) -> list[str]:
    """Delete shapes (instruction callouts and their connectors) from the slide."""
    removed = []
    for s in list(iter_shapes(prs.slides[slide_index].shapes)):
        if s.shape_id in shape_ids:
            s._element.getparent().remove(s._element)
            removed.append(f"{s.name} (id {s.shape_id})")
    return removed

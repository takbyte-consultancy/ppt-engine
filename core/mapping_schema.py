"""
The mapping contract shared by all three agents.

A template = template.pptx (every shape uniquely named) + mapping.json (this schema).
Each component says: "shape X on the slide gets its data from JSONPath Y in the API response".

JSONPath conventions
  - Absolute paths start with "$"           e.g.  $.header.title
  - Table column / chart point paths are RELATIVE to each row (no "$")   e.g.  name, metrics.aum
  - Tables come in two modes:
      records : rows_source -> array of objects, one ColumnMap per column (column-level mapping)
      grid    : source -> 2-D array data[row][col] that matches the table cell-for-cell
                (every cell, including multi-level headers and row labels, is mapped by position)
  - Text 'format' is a Python format string: "{value}", "{value:,.2f}", "As of {value}"
    Named formats are also allowed: currency, percent, number, date
"""
from __future__ import annotations

from typing import Literal, Optional
from jsonpath_ng.ext import parse as jp_parse
from pydantic import BaseModel, Field, model_validator

ComponentType = Literal["text", "table", "chart", "image", "static"]


class ColumnMap(BaseModel):
    header: str = Field(description="Column header text as shown on the slide")
    source: str = Field(description="JSONPath RELATIVE to one row, e.g. 'name'")
    format: Optional[str] = None


class SeriesMap(BaseModel):
    name: str
    values_source: str = Field(description="JSONPath RELATIVE to one row, e.g. 'weight'")


class RunMap(BaseModel):
    """Run-level mapping for text shapes that mix locked (red bold) and dynamic text."""
    paragraph: int
    run: int
    locked: bool = False
    source: Optional[str] = None
    format: Optional[str] = None


class CalloutDecision(BaseModel):
    """What the agent decided about each yellow/green callout candidate."""
    shape_id: int
    decision: Literal["instruction", "design_element"]
    instruction_text: str = ""
    target_shape: Optional[str] = None        # renamed shape the instruction applies to
    how_applied: str = ""                     # what changed in the mapping because of it
    needs_human: bool = False                 # instruction that a mapping cannot express


class Component(BaseModel):
    shape: str = Field(description="Unique shape name on the slide")
    type: ComponentType
    description: str = ""
    override_reason: str = ""                 # why a shape whose content IS in the data stays static
                                              # (or why a cell-for-cell grid isn't used); shown to reviewers
    shape_id: Optional[int] = None            # locator, filled in by the workflow
    original_name: str = ""                   # name in the Platinum copy, filled in by the workflow

    # text / image
    source: Optional[str] = None
    format: Optional[str] = None
    runs: list[RunMap] = []                   # text: per-run mapping when the shape has locked runs

    # table
    table_mode: Literal["records", "grid"] = "records"
    rows_source: Optional[str] = None          # records mode: absolute path to an array
    columns: list[ColumnMap] = []
    header_row: bool = True
    max_rows: Optional[int] = None             # overflow rule for Agent 2
    locked_cells: list[list[int]] = []         # [row, col] cells never written (red bold), set by code

    # chart
    chart_type: Optional[str] = None
    points_source: Optional[str] = None        # absolute path to an array
    categories_source: Optional[str] = None    # RELATIVE path per point
    series: list[SeriesMap] = []

    @model_validator(mode="after")
    def _check_fields(self):
        t = self.type
        if t == "image" and not self.source:
            raise ValueError(f"[{self.shape}] image component needs 'source'")
        if t == "text" and not self.source and not self.runs:
            raise ValueError(f"[{self.shape}] text component needs 'source' or 'runs'")
        for rm in self.runs:
            if not rm.locked and not rm.source:
                raise ValueError(f"[{self.shape}] run p{rm.paragraph}.r{rm.run} needs 'source' or locked=true")
        if t == "table" and self.table_mode == "grid" and not self.source:
            raise ValueError(f"[{self.shape}] grid table needs 'source' (path to a 2-D array)")
        if t == "table" and self.table_mode == "records" and (not self.rows_source or not self.columns):
            raise ValueError(f"[{self.shape}] records table needs 'rows_source' and 'columns'")
        if t == "chart" and (not self.points_source or not self.categories_source or not self.series):
            raise ValueError(f"[{self.shape}] chart needs 'points_source', 'categories_source', 'series'")
        return self


class TemplateMapping(BaseModel):
    template_id: str
    version: int = 1
    template_key: str = ""                      # e.g. FIXED_INCOME::IMNASMA
    key_fields: dict[str, str] = {}             # e.g. {"asset_type": "...", "acronym": "..."}
    datasource_id: str                          # key into the datasource repository
    exhibit_type: str = ""                      # used by Agent 3 to find exhibits
    exhibit_name: str = ""
    required_params: list[str] = []             # copied from the datasource entry
    components: list[Component]
    callouts: list[CalloutDecision] = []


# --------------------------------------------------------------------------- #
# Validation against the real slide and a real (sample) API response
# --------------------------------------------------------------------------- #
def _find(path: str, data):
    return [m.value for m in jp_parse(path).find(data)]


def validate_mapping(mapping: TemplateMapping, shape_names: set[str], sample: dict,
                     table_dims: dict[str, tuple[int, int]] | None = None) -> list[str]:
    """Return a list of human-readable problems. Empty list == valid.
    table_dims: {shape_name: (rows, cols)} lets grid tables be checked against the real table."""
    table_dims = table_dims or {}
    errors: list[str] = []

    mapped = [c.shape for c in mapping.components]
    for dup in {s for s in mapped if mapped.count(s) > 1}:
        errors.append(f"Shape '{dup}' is mapped more than once")
    for missing in shape_names - set(mapped):
        errors.append(f"Shape '{missing}' has no mapping (use type 'static' if it never changes)")

    for c in mapping.components:
        if c.shape not in shape_names:
            errors.append(f"[{c.shape}] no shape with this name exists on the slide")
            continue
        try:
            if c.type == "text" and c.runs:
                for rm in c.runs:
                    if rm.source and not _find(rm.source, sample):
                        errors.append(f"[{c.shape}] run p{rm.paragraph}.r{rm.run}: '{rm.source}' matched nothing")
            elif c.type in ("text", "image"):
                if not _find(c.source, sample):
                    errors.append(f"[{c.shape}] '{c.source}' matched nothing in sample JSON")

            elif c.type == "table" and c.table_mode == "grid":
                found = _find(c.source, sample)
                grid = found[0] if found else None
                if not isinstance(grid, list) or not all(isinstance(r, list) for r in grid):
                    errors.append(f"[{c.shape}] grid source '{c.source}' must resolve to a 2-D array")
                elif c.shape in table_dims:
                    n_rows, n_cols = table_dims[c.shape]
                    widest = max((len(r) for r in grid), default=0)
                    if len(grid) != n_rows or widest != n_cols:
                        errors.append(f"[{c.shape}] JSON grid is {len(grid)}x{widest} but the table is "
                                      f"{n_rows}x{n_cols}")

            elif c.type == "table":
                rows = _find(c.rows_source, sample)
                rows = rows[0] if len(rows) == 1 and isinstance(rows[0], list) else rows
                if not rows:
                    errors.append(f"[{c.shape}] rows_source '{c.rows_source}' matched nothing")
                else:
                    for col in c.columns:
                        if not _find(col.source, rows[0]):
                            errors.append(f"[{c.shape}] column '{col.header}': '{col.source}' not found in first row")

            elif c.type == "chart":
                pts = _find(c.points_source, sample)
                pts = pts[0] if len(pts) == 1 and isinstance(pts[0], list) else pts
                if not pts:
                    errors.append(f"[{c.shape}] points_source '{c.points_source}' matched nothing")
                else:
                    if not _find(c.categories_source, pts[0]):
                        errors.append(f"[{c.shape}] categories_source '{c.categories_source}' not found in first point")
                    for s in c.series:
                        if not _find(s.values_source, pts[0]):
                            errors.append(f"[{c.shape}] series '{s.name}': '{s.values_source}' not found")
        except Exception as e:  # bad JSONPath syntax etc.
            errors.append(f"[{c.shape}] invalid JSONPath: {e}")
    return errors

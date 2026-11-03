"""
alias.json: the contract between a template and its datasource.

The template contains short placeholders ({{A}}, {{B}}, ...) wherever data goes; static text stays
as written. alias.json says, for every alias, which JSON path supplies the value and how to format it,
and records exactly where each placeholder came from (for review and for re-validation).
"""
from __future__ import annotations

from typing import Any, Literal, Optional

from pydantic import BaseModel, Field

CaseRule = Literal["lower", "upper", "lower_first", "upper_first"]


class AliasDef(BaseModel):
    path: str = Field(description="JSONPath into the datasource response")
    format: Optional[str] = None          # python format for numbers, e.g. "{value:.2f}"
    case: Optional[CaseRule] = None       # applied to strings, e.g. lower_first: "Periods" -> "periods"
    null: str = ""                        # what to show when the value is null, e.g. "—"
    join: Optional[str] = None            # when the path returns several values, join them with this
    description: str = ""                 # human explanation (the template only shows the short alias)
    sample: Any = None                    # value seen when the template was created


class Location(BaseModel):
    paragraph: int
    cell: Optional[list[int]] = None      # [row, col] for table cells


class Occurrence(BaseModel):
    id: str
    alias: str
    shape_id: int
    shape_name: str = ""
    location: Location
    start: int                            # character span inside the paragraph text
    end: int
    text: str                             # the Platinum copy text that the placeholder replaces
    how: str = ""                         # why code (or the agent) chose this
    confidence: Literal["high", "medium", "low"] = "high"
    candidates: list[str] = []            # other paths holding the same value


class StaticLabel(BaseModel):
    shape_id: int
    location: Location
    text: str
    key_path: str                         # the JSON key this label names


class KeptStatic(BaseModel):
    shape_id: int
    location: Location
    text: str
    reason: str                           # why data-matching text stays literal (shown to reviewers)


class SeriesBinding(BaseModel):
    name: str
    values: str                           # JSONPath to the list of values


class ChartBinding(BaseModel):
    shape_id: int
    categories: str                       # JSONPath to the list of category labels
    series: list[SeriesBinding]


class CalloutDecision(BaseModel):
    shape_id: int
    decision: Literal["instruction", "design_element"]
    instruction_text: str = ""
    target_shape_id: Optional[int] = None
    how_applied: str = ""
    needs_human: bool = False


class AliasDocument(BaseModel):
    template_id: str = ""
    template_key: str = ""
    key_fields: dict[str, str] = {}
    datasource_id: str = ""
    version: int = 1
    required_params: list[str] = []
    aliases: dict[str, AliasDef] = {}
    occurrences: list[Occurrence] = []
    charts: list[ChartBinding] = []
    labels: list[StaticLabel] = []
    kept_static: list[KeptStatic] = []
    callouts: list[CalloutDecision] = []

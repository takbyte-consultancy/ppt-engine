"""
Excel -> the same dict shape the mapping understands. For every sheet:
  "<sheet>": {"grid": [[cell, ...], ...],          # every cell, for grid-mode tables
              "records": [{"Header": value}, ...]}  # row 1 as headers, for records-mode tables
Values are what Excel displays as stored (formula results, not formulas).
"""
from __future__ import annotations

from openpyxl import load_workbook

from integrations.datasources.base import check_params, local_path


def _clean(v):
    if v is None:
        return ""
    if hasattr(v, "isoformat"):
        return v.isoformat()
    return v


class ExcelFileSource:
    def __init__(self, definition: dict):
        self.definition = definition

    def fetch(self, params=None) -> dict:
        params = check_params(self.definition, params)
        wb = load_workbook(local_path(self.definition["location"], params), data_only=True, read_only=True)
        out = {"sheets": {}}
        for ws in wb.worksheets:
            grid = [[_clean(v) for v in row] for row in ws.iter_rows(values_only=True)]
            while grid and all(v == "" for v in grid[-1]):
                grid.pop()
            headers = [str(h) for h in grid[0]] if grid else []
            records = [dict(zip(headers, r)) for r in grid[1:]] if headers else []
            out["sheets"][ws.title] = {"grid": grid, "records": records}
        wb.close()
        return out

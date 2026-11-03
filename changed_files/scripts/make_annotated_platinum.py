"""Creates mock/platinum_annotated.pptx + mock/annotated.json: a Platinum copy using every marker."""
import json
from pathlib import Path
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.util import Inches, Pt

NAVY, RED = RGBColor(0x1F, 0x2A, 0x44), RGBColor(0xC0, 0x00, 0x00)
ROOT = Path(__file__).parent.parent
prs = Presentation(); prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
s = prs.slides.add_slide(prs.slide_layouts[6])


def tb(x, y, w, h, name):
    t = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h)); t.name = name
    return t


def run(p, text, size, bold=False, color=NAVY):
    r = p.add_run(); r.text = text; r.font.size = Pt(size); r.font.bold = bold; r.font.color.rgb = color
    return r


t = tb(0.5, 0.4, 8, 0.7, "Title 1"); run(t.text_frame.paragraphs[0], "Portfolio Overview — ABC Fund", 26, True)
d = tb(0.5, 1.1, 5, 0.4, "Date 2"); p = d.text_frame.paragraphs[0]
run(p, "As of ", 14, True, RED); run(p, "30 Jun 2026", 14)                       # partially locked
f = tb(0.5, 6.8, 9, 0.4, "Footer 3"); run(f.text_frame.paragraphs[0], "Source: Capital Group. Internal use only.", 10, True, RED)

tbl = s.shapes.add_table(3, 3, Inches(0.5), Inches(2.0), Inches(6), Inches(1.5)); tbl.name = "Returns 4"
grid = [["", "3M", "1Y"], ["Fund", "1.2", "8.4"], ["Benchmark", "1.0", "7.1"]]
for r_i, row in enumerate(grid):
    for c_i, v in enumerate(row):
        cell = tbl.table.cell(r_i, c_i); cell.text = ""
        if v:
            locked = (r_i, c_i) == (2, 0)
            run(cell.text_frame.paragraphs[0], v, 12, locked, RED if locked else NAVY)

kpi = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.5), Inches(2.0), Inches(3), Inches(1.2)); kpi.name = "KPI 5"
kpi.fill.solid(); kpi.fill.fore_color.rgb = RGBColor(0x2E, 0x8B, 0x57)                   # green, but design
run(kpi.text_frame.paragraphs[0], "+8.4% YTD", 24, True, RGBColor(0xFF, 0xFF, 0xFF))

y = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(7.5), Inches(4.6), Inches(4.5), Inches(0.9)); y.name = "Note 6"
y.fill.solid(); y.fill.fore_color.rgb = RGBColor(0xFF, 0xEB, 0x3B)                      # yellow callout
run(y.text_frame.paragraphs[0], "Date must come from the API field as_of_date", 12, color=RGBColor(0, 0, 0))
con = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(7.5), Inches(5.0), Inches(3.0), Inches(1.3))
con.name = "Arrow 7"

g = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(8.8), Inches(0.35), Inches(4.0), Inches(0.8)); g.name = "Note 8"
g.fill.solid(); g.fill.fore_color.rgb = RGBColor(0x92, 0xD0, 0x50)                      # green callout
run(g.text_frame.paragraphs[0], "Title: fund name from data (fund_name)", 12, color=RGBColor(0, 0, 0))

prs.save(ROOT / "mock" / "platinum_annotated.pptx")
(ROOT / "mock" / "annotated.json").write_text(json.dumps({
    "fund_name": "ABC Fund", "as_of_date": "30 Jun 2026", "kpi_text": "+8.4% YTD",
    "returns": {"periods": ["3M", "1Y"], "fund": {"3M": 1.2, "1Y": 8.4}, "benchmark": {"3M": 1.0, "1Y": 7.1}}},
    indent=2), encoding="utf-8")
print("saved mock/platinum_annotated.pptx and mock/annotated.json")

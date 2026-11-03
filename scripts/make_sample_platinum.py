"""Creates mock/platinum_equity.pptx: a realistic Platinum copy for testing without JIRA."""
from pathlib import Path
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches, Pt

NAVY = RGBColor(0x1F, 0x2A, 0x44)
prs = Presentation()
prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
s = prs.slides.add_slide(prs.slide_layouts[6])

bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Inches(0.15))
bar.fill.solid(); bar.fill.fore_color.rgb = NAVY; bar.line.fill.background()

def text(x, y, w, h, t, size, bold=False):
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    p = tb.text_frame.paragraphs[0]; p.text = t
    p.font.size, p.font.bold, p.font.color.rgb = Pt(size), bold, NAVY
    return tb

text(0.5, 0.35, 9, 0.6, "Global Equity Strategy — Portfolio Overview", 26, True)
text(0.5, 0.95, 6, 0.4, "As of 30 Jun 2026", 14)
text(9.6, 0.4, 3.3, 0.9, "AUM: $1,180.2mn", 20, True)

tbl = s.shapes.add_table(6, 4, Inches(0.5), Inches(1.6), Inches(6.3), Inches(2.6)).table
rows = [["Holding", "Ticker", "Weight", "Return"], ["Microsoft Corp", "MSFT", "6.0%", "10.1%"],
        ["Apple Inc", "AAPL", "5.5%", "6.2%"], ["Nestle SA", "NESN", "4.3%", "-0.8%"],
        ["HDFC Bank", "HDFCBANK", "3.5%", "8.0%"], ["Novo Nordisk", "NOVO-B", "3.4%", "3.9%"]]
for r, row in enumerate(rows):
    for c, v in enumerate(row):
        tbl.cell(r, c).text = v

pie = CategoryChartData(); pie.categories = ["Technology", "Financials", "Healthcare", "Consumer", "Other"]
pie.add_series("Weight", (30, 22, 18, 19, 11))
ch = s.shapes.add_chart(XL_CHART_TYPE.PIE, Inches(7.2), Inches(1.5), Inches(5.6), Inches(2.9), pie).chart
ch.has_legend = True; ch.legend.position = XL_LEGEND_POSITION.RIGHT

perf = CategoryChartData(); perf.categories = ["1M", "3M", "YTD", "1Y"]
perf.add_series("Fund", (1.0, 3.5, 7.9, 12.0)); perf.add_series("Benchmark", (0.8, 3.0, 6.8, 10.5))
ch2 = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.5), Inches(4.5), Inches(12.3), Inches(2.3), perf).chart
ch2.has_legend = True; ch2.legend.position = XL_LEGEND_POSITION.BOTTOM

text(0.5, 6.95, 12, 0.35, "For internal use only. Past performance is not indicative of future results.", 9)

out = Path(__file__).parent.parent / "mock" / "platinum_equity.pptx"
prs.save(out); print("saved", out)

#!/usr/bin/env python3
"""
Equity research PDF report template (reportlab platypus).

HOW TO USE THIS FILE:
  1. Copy it into your working directory and rename it for the ticker you're building.
  2. Search for "TICKER-SPECIFIC" — replace every placeholder paragraph/table with real,
     sourced content from your research. Keep the helper functions and section order.
  3. Build, then render 2-3 pages to PNG and actually look at them before delivering:
       pdftoppm -png -r 80 -f 1 -l 1 <output.pdf> preview
     Check the cover, a data-heavy page, and the last page for overflow/orphaned headers.
  4. Delete preview PNGs before delivering the final files.

Section order (mirrors SKILL.md's "Report structure"): cover/exec summary, business model
mechanics, industry/competitive/regulatory landscape, financial analysis, valuation, technical
analysis, catalyst calendar, key risks, sources & limitations.
"""
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, PageBreak, Table, TableStyle,
    HRFlowable, ListFlowable, ListItem
)
from reportlab.pdfgen import canvas as canvas_mod
from reportlab.graphics.shapes import Drawing, String, Line, Rect
from reportlab.graphics.charts.linecharts import HorizontalLineChart
from reportlab.graphics.charts.barcharts import VerticalBarChart
from reportlab.graphics.charts.lineplots import LinePlot
from reportlab.graphics.widgets.markers import makeMarker

# TICKER-SPECIFIC: identity of the report
TICKER = "TICK"
COMPANY_NAME = "Example Company, Inc."
REPORT_DATE = "Month DD, YYYY"
SECTOR = "Sector / Industry"

NAVY = colors.HexColor("#1F3864")
BLUE2 = colors.HexColor("#2E5395")
LIGHTBLUE = colors.HexColor("#D9E2F3")
GRAY = colors.HexColor("#595959")
LIGHTGRAY = colors.HexColor("#F2F2F2")
AMBER = colors.HexColor("#B8860B")

styles = getSampleStyleSheet()
styles.add(ParagraphStyle("ReportTitle", fontName="Helvetica-Bold", fontSize=26, textColor=NAVY, spaceAfter=6, leading=30))
styles.add(ParagraphStyle("ReportSubtitle", fontName="Helvetica", fontSize=13, textColor=GRAY, spaceAfter=4))
styles.add(ParagraphStyle("ReportMeta", fontName="Helvetica", fontSize=9.5, textColor=GRAY, spaceAfter=2))
styles.add(ParagraphStyle("H1", fontName="Helvetica-Bold", fontSize=15, textColor=colors.white, leading=18))
# keepWithNext stops a sub-heading stranding at the foot of a page with its table or chart on the
# next one. Adding charts to this template made that visible immediately — a heading landed alone
# at the bottom of a page and its table opened the next one. One attribute, fixed everywhere.
styles.add(ParagraphStyle("H2", fontName="Helvetica-Bold", fontSize=12, textColor=NAVY, spaceBefore=14, spaceAfter=6, keepWithNext=1))
styles.add(ParagraphStyle("H3", fontName="Helvetica-Bold", fontSize=10.5, textColor=BLUE2, spaceBefore=8, spaceAfter=4, keepWithNext=1))
styles.add(ParagraphStyle("Body", fontName="Helvetica", fontSize=9.5, leading=13.5, spaceAfter=7, alignment=TA_JUSTIFY))
styles.add(ParagraphStyle("BodySmall", fontName="Helvetica", fontSize=8.5, leading=12, spaceAfter=5, alignment=TA_JUSTIFY, textColor=colors.HexColor("#333333")))
styles.add(ParagraphStyle("Note", fontName="Helvetica-Oblique", fontSize=8, leading=11, textColor=GRAY, spaceAfter=6))
styles.add(ParagraphStyle("MyBullet", fontName="Helvetica", fontSize=9.5, leading=13.5, spaceAfter=4))
styles.add(ParagraphStyle("TableCell", fontName="Helvetica", fontSize=8, leading=10))
styles.add(ParagraphStyle("TableCellBold", fontName="Helvetica-Bold", fontSize=8, leading=10))
styles.add(ParagraphStyle("TableHdr", fontName="Helvetica-Bold", fontSize=8, leading=10, textColor=colors.white, alignment=TA_CENTER))
styles.add(ParagraphStyle("StatLabel", fontName="Helvetica", fontSize=8, textColor=GRAY))
styles.add(ParagraphStyle("StatValue", fontName="Helvetica-Bold", fontSize=13, textColor=NAVY))


def section_header(text):
    """Full-width navy section header bar — use one per major report section."""
    t = Table([[Paragraph(text, styles["H1"])]], colWidths=[6.85 * inch])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), NAVY),
        ("LEFTPADDING", (0, 0), (-1, -1), 10),
        ("TOPPADDING", (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
    ]))
    return t


def data_table(header, rows, col_widths, header_bg=BLUE2, align_cols=None, zebra=True):
    """Zebra-striped data table. rows is a list of lists of strings (HTML-lite markup OK)."""
    align_cols = align_cols or {}
    data = [[Paragraph(h, styles["TableHdr"]) for h in header]]
    for r in rows:
        row_out = []
        for i, val in enumerate(r):
            st = styles["TableCellBold"] if (i == 0) else styles["TableCell"]
            row_out.append(Paragraph(str(val), st))
        data.append(row_out)
    t = Table(data, colWidths=col_widths, repeatRows=1)
    style = [
        ("BACKGROUND", (0, 0), (-1, 0), header_bg),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#BFBFBF")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 5), ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 3.5), ("BOTTOMPADDING", (0, 0), (-1, -1), 3.5),
    ]
    if zebra:
        for i in range(1, len(data)):
            if i % 2 == 0:
                style.append(("BACKGROUND", (0, i), (-1, i), LIGHTGRAY))
    for col, al in align_cols.items():
        style.append(("ALIGN", (col, 1), (col, -1), al))
    t.setStyle(TableStyle(style))
    return t


def stat_box(label, value, width=1.62 * inch, color=NAVY):
    """Small KPI box for the cover page snapshot row."""
    t = Table([[Paragraph(label, styles["StatLabel"])],
               [Paragraph(value, ParagraphStyle("v", parent=styles["StatValue"], textColor=color))]],
              colWidths=[width])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), LIGHTBLUE),
        ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#BFBFBF")),
        ("LEFTPADDING", (0, 0), (-1, -1), 8), ("TOPPADDING", (0, 0), (-1, -1), 6), ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
    ]))
    return t


def bullets(items):
    return ListFlowable(
        [ListItem(Paragraph(i, styles["MyBullet"]), leftIndent=10, bulletColor=NAVY) for i in items],
        bulletType="bullet", start="circle", leftIndent=14
    )


def disclaimer_box(text):
    t = Table([[Paragraph(text, styles["BodySmall"])]], colWidths=[6.85 * inch])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#FCE4D6")),
        ("BOX", (0, 0), (-1, -1), 0.75, AMBER),
        ("LEFTPADDING", (0, 0), (-1, -1), 10), ("RIGHTPADDING", (0, 0), (-1, -1), 10),
        ("TOPPADDING", (0, 0), (-1, -1), 8), ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
    ]))
    return t


# ============================================================================
# CHARTS
#   A sell-side report is not a wall of tables — the price history, the revenue/margin
#   trajectory, and where the name sits against its comps are all things a reader takes in
#   visually in a second and would have to reconstruct from a table over a minute.
#   These use reportlab.graphics, which ships with reportlab: no extra dependency, no image
#   files to manage, and the output stays vector (sharp at any zoom, unlike an embedded PNG).
#
#   All three return a Drawing, which is a platypus flowable — append it to `story` like any
#   Paragraph or Table. Keep widths at 6.85*inch to line up with the section bars and tables.
# ============================================================================
CHART_SERIES_COLORS = [NAVY, colors.HexColor("#C55A11"), colors.HexColor("#548235"),
                       colors.HexColor("#7030A0"), GRAY]


def chart_caption(drawing, text, width=6.85 * inch):
    """Chart titles belong ON the chart, not in a separate paragraph that can page-break away."""
    drawing.add(String(0, drawing.height - 10, text,
                       fontName="Helvetica-Bold", fontSize=9.5, fillColor=NAVY))
    return drawing


def line_chart(categories, series, labels, title, y_label_fmt="{:,.0f}",
               width=6.85 * inch, height=2.15 * inch, fill_zero=False):
    """Multi-series line chart — price history, revenue trend, margin trajectory.

    categories: x-axis labels (quarters, years, dates)
    series:     list of lists, one inner list per line, all the same length as categories
    labels:     one legend label per series
    """
    d = Drawing(width, height)
    lc = HorizontalLineChart()
    lc.x, lc.y = 42, 30
    lc.width, lc.height = width - 70, height - 62
    lc.data = series
    lc.categoryAxis.categoryNames = [str(c) for c in categories]
    lc.categoryAxis.labels.fontName = "Helvetica"
    lc.categoryAxis.labels.fontSize = 7
    lc.categoryAxis.labels.angle = 0 if len(categories) <= 9 else 45
    lc.categoryAxis.labels.dy = -4
    if lc.categoryAxis.labels.angle:
        lc.categoryAxis.labels.boxAnchor = "ne"
    lc.valueAxis.labelTextFormat = y_label_fmt.format
    lc.valueAxis.labels.fontName = "Helvetica"
    lc.valueAxis.labels.fontSize = 7
    flat = [v for srs in series for v in srs if v is not None]
    lo, hi = min(flat), max(flat)
    pad = (hi - lo) * 0.12 or abs(hi) * 0.1 or 1
    lc.valueAxis.valueMin = 0 if fill_zero else lo - pad
    lc.valueAxis.valueMax = hi + pad
    lc.valueAxis.visibleGrid = 1
    lc.valueAxis.gridStrokeColor = colors.HexColor("#E0E0E0")
    lc.valueAxis.gridStrokeWidth = 0.4
    for i in range(len(series)):
        lc.lines[i].strokeColor = CHART_SERIES_COLORS[i % len(CHART_SERIES_COLORS)]
        lc.lines[i].strokeWidth = 1.6
        # Markers make single points readable and, more importantly, keep the chart legible
        # for a reader who prints it in black and white.
        lc.lines[i].symbol = makeMarker(["FilledCircle", "FilledSquare", "FilledDiamond",
                                         "FilledTriangle", "Circle"][i % 5])
        lc.lines[i].symbol.size = 3
        lc.lines[i].symbol.fillColor = CHART_SERIES_COLORS[i % len(CHART_SERIES_COLORS)]
    d.add(lc)
    # Inline legend, drawn rather than using reportlab's Legend widget, which is fiddly to place.
    lx = 42
    for i, lab in enumerate(labels):
        col = CHART_SERIES_COLORS[i % len(CHART_SERIES_COLORS)]
        d.add(Line(lx, 12, lx + 12, 12, strokeColor=col, strokeWidth=2))
        d.add(String(lx + 16, 9, lab, fontName="Helvetica", fontSize=7, fillColor=GRAY))
        lx += 26 + 4.6 * len(lab)
    return chart_caption(d, title, width)


def bar_chart(categories, values, title, highlight_index=None, value_fmt="{:,.1f}",
              width=6.85 * inch, height=2.15 * inch, y_label_fmt="{:,.1f}"):
    """Single-series bar chart — comps multiples, peer growth rates, segment mix.

    highlight_index: index of the subject company's bar, drawn in navy while peers stay gray.
    That one touch is what turns a generic bar chart into "where WE sit versus the group."
    """
    d = Drawing(width, height)
    bc = VerticalBarChart()
    bc.x, bc.y = 42, 26
    bc.width, bc.height = width - 70, height - 58
    bc.data = [values]
    bc.categoryAxis.categoryNames = [str(c) for c in categories]
    bc.categoryAxis.labels.fontName = "Helvetica"
    bc.categoryAxis.labels.fontSize = 7
    bc.categoryAxis.labels.dy = -4
    bc.valueAxis.valueMin = 0
    bc.valueAxis.valueMax = max(values) * 1.18
    bc.valueAxis.labelTextFormat = y_label_fmt.format
    bc.valueAxis.labels.fontName = "Helvetica"
    bc.valueAxis.labels.fontSize = 7
    bc.valueAxis.visibleGrid = 1
    bc.valueAxis.gridStrokeColor = colors.HexColor("#E0E0E0")
    bc.valueAxis.gridStrokeWidth = 0.4
    bc.bars[0].fillColor = colors.HexColor("#A6B8D6")
    bc.bars[0].strokeColor = None
    bc.barSpacing = 1.5
    if highlight_index is not None:
        bc.bars[(0, highlight_index)].fillColor = NAVY
    d.add(bc)
    # Value labels on top of each bar — a comps chart is read for the numbers, not the shape.
    n = len(values)
    plot_w = bc.width
    slot = plot_w / n
    for i, v in enumerate(values):
        x = bc.x + slot * (i + 0.5)
        y = bc.y + (v / bc.valueAxis.valueMax) * bc.height + 3
        d.add(String(x, y, value_fmt.format(v), fontName="Helvetica-Bold", fontSize=6.5,
                     fillColor=NAVY if i == highlight_index else GRAY, textAnchor="middle"))
    return chart_caption(d, title, width)


def scenario_chart(labels, values, current_price, title,
                   width=6.85 * inch, height=2.0 * inch):
    """Bear/Base/Bull price targets against the current price.

    The horizontal current-price line is the whole point: it makes "the market is pricing
    something between Bear and Base" a thing the reader sees rather than a sentence they parse.
    """
    d = Drawing(width, height)
    bc = VerticalBarChart()
    bc.x, bc.y = 46, 26
    bc.width, bc.height = width - 78, height - 56
    bc.data = [values]
    bc.categoryAxis.categoryNames = [str(l) for l in labels]
    bc.categoryAxis.labels.fontName = "Helvetica-Bold"
    bc.categoryAxis.labels.fontSize = 7.5
    bc.categoryAxis.labels.dy = -4
    top = max(values + [current_price]) * 1.22
    bc.valueAxis.valueMin, bc.valueAxis.valueMax = 0, top
    bc.valueAxis.labelTextFormat = "${:,.0f}".format
    bc.valueAxis.labels.fontName = "Helvetica"
    bc.valueAxis.labels.fontSize = 7
    bc.bars[0].strokeColor = None
    bc.barSpacing = 3
    for i, v in enumerate(values):
        bc.bars[(0, i)].fillColor = (colors.HexColor("#C00000") if v < current_price
                                     else colors.HexColor("#548235"))
    d.add(bc)
    y_price = bc.y + (current_price / top) * bc.height
    d.add(Line(bc.x, y_price, bc.x + bc.width, y_price,
               strokeColor=NAVY, strokeWidth=1.1, strokeDashArray=[3, 2]))
    # The label sits ON the plot area, so it can land on top of a bar. A white knockout box behind
    # it keeps it readable whichever scenario happens to be tall — without one, navy text on the
    # green Bull bar is close to unreadable.
    label = f"current ${current_price:,.2f}"
    lw = 4.0 * len(label)
    d.add(Rect(bc.x + 3, y_price + 2.5, lw, 9, fillColor=colors.white, strokeColor=None))
    d.add(String(bc.x + 5, y_price + 4.5, label,
                 fontName="Helvetica-Bold", fontSize=7, fillColor=NAVY, textAnchor="start"))
    n = len(values)
    slot = bc.width / n
    for i, v in enumerate(values):
        d.add(String(bc.x + slot * (i + 0.5), bc.y + (v / top) * bc.height + 3,
                     f"${v:,.2f}", fontName="Helvetica-Bold", fontSize=7,
                     fillColor=GRAY, textAnchor="middle"))
    return chart_caption(d, title, width)


DISCLAIMER_TEXT = (
    "<b>IMPORTANT:</b> This document was produced by an AI research assistant from publicly "
    "available secondary sources as of the report date. It is not built on a live market data "
    "terminal feed, has not been reviewed by a licensed analyst, and figures that showed "
    "variance across sources are flagged in the Sources &amp; Limitations section. This is not "
    "investment, legal, or tax advice, and nothing here is a recommendation to buy, hold, or "
    "sell any security. Verify anything decision-critical against primary sources before acting."
)

story = []

# ============================================================================
# COVER PAGE — TICKER-SPECIFIC: snapshot stats and the 3-paragraph thesis.
# ============================================================================
story.append(Spacer(1, 0.4 * inch))
story.append(Paragraph(COMPANY_NAME.upper(), styles["ReportTitle"]))
story.append(Paragraph(f"{TICKER} — Equity Research: Fundamental &amp; Technical Analysis", styles["ReportSubtitle"]))
story.append(Spacer(1, 4))
story.append(HRFlowable(width="100%", thickness=1.4, color=NAVY))
story.append(Spacer(1, 10))
story.append(Paragraph(
    f"Report date: {REPORT_DATE} &nbsp;&nbsp;|&nbsp;&nbsp; Sector: {SECTOR} &nbsp;&nbsp;|&nbsp;&nbsp; "
    "Prepared for informational/educational purposes only", styles["ReportMeta"]))
story.append(Spacer(1, 22))

stats_row1 = Table([[
    stat_box("LAST PRICE", "$0.00"),
    stat_box("MARKET CAP", "$0.0B"),
    stat_box("52-WK RANGE", "$0.00 - $0.00"),
    stat_box("CONSENSUS PT", "$0.00"),
]], colWidths=[1.62 * inch] * 4)
stats_row1.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 3), ("RIGHTPADDING", (0, 0), (-1, -1), 3)]))
story.append(stats_row1)
story.append(Spacer(1, 24))

story.append(Paragraph("EXECUTIVE SUMMARY", styles["H2"]))
# TICKER-SPECIFIC: 3 paragraphs. Para 1 = what the company does and why it matters as an
# investment story. Para 2 = current situation / recent trajectory (what changed recently and
# why the stock has moved). Para 3 = state of the debate (what bulls and bears disagree about,
# and that this report doesn't resolve it).
story.append(Paragraph(
    "[Paragraph 1 — the business in one paragraph: what it does, how it makes money, and the "
    "one sentence that captures why this is or isn't a straightforward investment story.]",
    styles["Body"]))
story.append(Paragraph(
    "[Paragraph 2 — recent trajectory: what's happened to the stock and the business lately, "
    "and the specific events/data that explain it.]", styles["Body"]))
story.append(Paragraph(
    "[Paragraph 3 — the state of the debate: where sell-side/analyst opinion is split and why, "
    "framed neutrally. Explicitly note this report lays out the mechanics and levers without "
    "resolving the debate, and is not a recommendation.]", styles["Body"]))

story.append(Spacer(1, 10))
story.append(Paragraph("WHAT THIS REPORT COVERS", styles["H3"]))
story.append(bullets([
    "Business mechanics — how the company actually makes money and what dominates its cost/revenue structure",
    "Industry &amp; competitive landscape — market share, regulatory environment, emerging disruptors",
    "Financial quality — quarterly trend analysis, margin compression/expansion, balance sheet",
    "Valuation — comps table and a bear/base/bull scenario framework (companion Excel model has live formulas)",
    "Technical analysis — price structure, moving averages, options positioning, short interest",
    "Risks &amp; forward catalysts — the specific, named things that could move this stock next",
]))

story.append(Spacer(1, 14))
story.append(disclaimer_box(DISCLAIMER_TEXT))
story.append(PageBreak())

# ============================================================================
# 1. BUSINESS MODEL & MECHANICS — TICKER-SPECIFIC
# ============================================================================
story.append(section_header("1. BUSINESS MODEL &amp; MECHANICS"))
story.append(Spacer(1, 10))
story.append(Paragraph("1.1 [The core economics, named]", styles["H2"]))
story.append(Paragraph(
    "[Explain the unit economics in enough detail that a reader understands the actual revenue "
    "driver — not \"they sell software,\" but how a dollar of revenue actually gets made, and "
    "what the largest cost or revenue-share relationship is if one dominates.]", styles["Body"]))
story.append(Paragraph("1.2 [The single biggest cost/partner relationship, if one exists]", styles["H2"]))
story.append(Paragraph("[Detail. Is there a dominant distribution partner, supplier, or customer concentration?]", styles["Body"]))
story.append(Paragraph("1.3 [Other revenue lines / growth initiatives]", styles["H2"]))
story.append(bullets(["[Line item 1]", "[Line item 2]", "[Line item 3]"]))
story.append(Paragraph("1.4 Capital structure &amp; ownership", styles["H2"]))
story.append(Paragraph(
    "[Share class structure, insider/founder control, balance sheet summary, dividend/buyback "
    "policy, major shareholders.]", styles["Body"]))
story.append(PageBreak())

# ============================================================================
# 2. INDUSTRY, COMPETITIVE & REGULATORY LANDSCAPE — TICKER-SPECIFIC
# ============================================================================
story.append(section_header("2. INDUSTRY, COMPETITIVE &amp; REGULATORY LANDSCAPE"))
story.append(Spacer(1, 10))
story.append(Paragraph("2.1 Market size and growth", styles["H2"]))
story.append(Paragraph("[Total addressable market, growth trend, credible forecasts with sources.]", styles["Body"]))
story.append(Paragraph("2.2 Competitive position", styles["H2"]))
story.append(data_table(
    ["Competitor", "Metric (e.g. market cap / share)", "Position"],
    [["[Competitor A]", "[value]", "[note]"], ["[Company subject]", "[value]", "[note]"]],
    [2.2 * inch, 2.3 * inch, 2.4 * inch],
))
story.append(Spacer(1, 8))
story.append(Paragraph("2.3 Regulatory environment", styles["H2"]))
story.append(Paragraph(
    "[The regulatory framework governing this business and what's currently in motion — laws, "
    "rulemaking timelines, and whether the effect on the subject company is a clean tailwind, a "
    "clean headwind, or genuinely two-sided.]", styles["Body"]))
story.append(Paragraph("2.4 Emerging disruptors", styles["H2"]))
story.append(bullets(["[Disruptor 1 — live threat or speculative? timeline?]", "[Disruptor 2]"]))
story.append(PageBreak())

# ============================================================================
# 3. FINANCIAL ANALYSIS — TICKER-SPECIFIC
# ============================================================================
story.append(section_header("3. FINANCIAL ANALYSIS"))
story.append(Spacer(1, 10))
story.append(Paragraph("3.1 Quarterly trend", styles["H2"]))
story.append(data_table(
    ["($mm)", "Period A", "Period B", "Period C", "Period D", "Period E"],
    [
        ["Total revenue", "0", "0", "0", "0", "0"],
        ["Net income / (loss)", "0", "0", "0", "0", "0"],
        ["Adjusted EBITDA", "n/a", "0", "0", "0", "0"],
        ["Revenue YoY growth", "n/a", "0%", "0%", "0%", "0%"],
    ],
    [2.15 * inch] + [0.94 * inch] * 5,
    align_cols={1: "CENTER", 2: "CENTER", 3: "CENTER", 4: "CENTER", 5: "CENTER"},
))
story.append(Spacer(1, 8))
# TICKER-SPECIFIC: same quarters and same figures as the table above, as plain numbers.
# The chart is the point of the section — quarter-over-quarter deceleration is close to invisible
# in a row of numbers and unmistakable as a line.
QTRS = ["Q1", "Q2", "Q3", "Q4", "Q1"]
REVENUE_BY_QTR = [1000, 1120, 1190, 1240, 1265]
EBITDA_BY_QTR = [180, 214, 238, 260, 250]
story.append(line_chart(
    QTRS, [REVENUE_BY_QTR, EBITDA_BY_QTR], ["Revenue ($mm)", "Adj. EBITDA ($mm)"],
    "Quarterly revenue and adjusted EBITDA", fill_zero=True))
story.append(Spacer(1, 6))
# TICKER-SPECIFIC: YoY growth and margin for the same quarters, in percent.
GROWTH_BY_QTR = [42.0, 38.5, 33.1, 27.4, 21.6]
MARGIN_BY_QTR = [18.0, 19.1, 20.0, 21.0, 19.8]
story.append(line_chart(
    QTRS, [GROWTH_BY_QTR, MARGIN_BY_QTR], ["Revenue growth YoY (%)", "Adj. EBITDA margin (%)"],
    "Growth and margin trajectory", y_label_fmt="{:,.0f}%"))
story.append(Spacer(1, 8))
story.append(Paragraph(
    "[State the read-through explicitly — is growth decelerating/accelerating, are margins "
    "expanding/compressing, and why. This sentence is the point of the table and the charts "
    "above it: name the inflection and say what caused it.]", styles["Body"]))
story.append(Paragraph("3.2 Guidance", styles["H2"]))
story.append(bullets(["[Latest management guidance, line by line, with source/date.]"]))
story.append(Paragraph("3.3 The key driver — [name it]", styles["H2"]))
story.append(Paragraph(
    "[This should mirror the Driver Sensitivity tab in the companion Excel model — explain why "
    "this one variable dominates the valuation, cite any independent sensitivity estimate found "
    "in research, and note explicitly if the forward path of this driver is contested across "
    "sources rather than resolved.]", styles["Body"]))
story.append(Paragraph("3.4 Balance sheet &amp; capital allocation", styles["H2"]))
story.append(data_table(
    ["Metric", "Value", "Note"],
    [["Cash &amp; equivalents", "$0", ""], ["Total debt", "$0", ""], ["Dividend", "[policy]", ""]],
    [1.7 * inch, 1.1 * inch, 4.05 * inch],
))
story.append(PageBreak())

# ============================================================================
# 4. VALUATION — TICKER-SPECIFIC
# ============================================================================
story.append(section_header("4. VALUATION"))
story.append(Spacer(1, 10))
story.append(Paragraph("4.1 Comparable company multiples", styles["H2"]))
story.append(data_table(
    ["Company", "Price", "Mkt Cap", "EV", "TTM Rev", "EV/Rev (TTM)", "Fwd P/E"],
    [["[Subject company]", "$0", "$0B", "$0B", "$0B", "0x", "0x"],
     ["[Comp A]", "$0", "$0B", "$0B", "$0B", "0x", "0x"]],
    [1.75 * inch] + [0.85 * inch] * 6,
    align_cols={1: "CENTER", 2: "CENTER", 3: "CENTER", 4: "CENTER", 5: "CENTER", 6: "CENTER"},
))
story.append(Spacer(1, 8))
# TICKER-SPECIFIC: same peer set and same multiple as the table above. Put the subject company
# FIRST and pass highlight_index=0 so its bar reads navy against gray peers.
COMP_NAMES = ["[Subject]", "[Comp A]", "[Comp B]", "[Comp C]", "[Comp D]"]
COMP_EV_REV = [12.4, 8.1, 9.6, 6.2, 14.8]
story.append(bar_chart(COMP_NAMES, COMP_EV_REV, "EV / TTM Revenue vs. peer set",
                       highlight_index=0, value_fmt="{:,.1f}x", y_label_fmt="{:,.0f}x"))
story.append(Spacer(1, 8))
story.append(Paragraph("[Note which comps are true comps vs. reference-only, and why. If the "
                       "subject trades at a visible premium or discount in the chart above, say "
                       "what the market is paying for or marking down — the gap is the argument.]",
                       styles["Body"]))
story.append(Paragraph("4.2 Scenario framework (see Excel model for live, adjustable version)", styles["H2"]))
story.append(data_table(
    ["", "Bear", "Base", "Bull"],
    [
        ["Forward revenue", "$0B", "$0B", "$0B"],
        ["Forward EV/Revenue multiple", "0x", "0x", "0x"],
        ["Implied share price", "$0", "$0", "$0"],
        ["Upside/(downside) vs. current", "0%", "0%", "0%"],
    ],
    [1.9 * inch, 1.65 * inch, 1.65 * inch, 1.65 * inch],
    align_cols={1: "CENTER", 2: "CENTER", 3: "CENTER"},
))
story.append(Spacer(1, 8))
# TICKER-SPECIFIC: the same three implied prices as the table above, plus the current price.
# The dashed current-price line is what makes this chart worth the space.
SCENARIO_PRICES = [42.0, 78.0, 118.0]
CURRENT_PRICE = 64.20
story.append(scenario_chart(["Bear", "Base", "Bull"], SCENARIO_PRICES, CURRENT_PRICE,
                            "Scenario price targets vs. current price"))
story.append(Spacer(1, 8))
story.append(Paragraph(
    "This is a simple sanity-check framework, not a full DCF — [point back to Section 3.3's "
    "driver model as the more rigorous basis]. The multiples used here are the report author's "
    "judgment calls — the Excel model's yellow input cells let you substitute your own.",
    styles["Body"]))
story.append(Paragraph("4.3 Sell-side coverage", styles["H2"]))
story.append(data_table(
    ["Date", "Firm", "Rating", "Price Target"],
    [["[date]", "[Firm A]", "[Rating]", "$0"], ["[date]", "[Firm B]", "[Rating]", "$0"]],
    [1.1 * inch, 1.85 * inch, 1.6 * inch, 2.3 * inch],
    align_cols={0: "CENTER", 2: "CENTER"},
))
story.append(Spacer(1, 6))
story.append(Paragraph(
    "[Note the breadth of coverage and the spread between the highest and lowest targets — a "
    "wide spread signals genuine analyst disagreement, worth calling out explicitly.]", styles["Note"]))
story.append(PageBreak())

# ============================================================================
# 5. TECHNICAL ANALYSIS — TICKER-SPECIFIC
# ============================================================================
story.append(section_header("5. TECHNICAL ANALYSIS"))
story.append(Spacer(1, 10))
story.append(Paragraph("5.1 Price structure", styles["H2"]))
# TICKER-SPECIFIC: price history plus moving averages. A technical-analysis section without a
# price chart is the single most conspicuous thing missing from an otherwise credible report —
# every level discussed below is a thing the reader wants to SEE.
PRICE_DATES = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
PRICE_CLOSE = [52.1, 55.4, 61.0, 58.2, 63.7, 71.5, 68.9, 74.2, 70.1, 66.8, 63.0, 64.2]
PRICE_MA50 = [50.8, 52.6, 55.9, 57.4, 59.2, 63.1, 66.0, 69.3, 70.4, 70.0, 67.9, 65.5]
PRICE_MA200 = [48.2, 49.0, 50.1, 51.3, 52.8, 54.9, 56.8, 59.0, 60.7, 62.0, 62.8, 63.1]
story.append(line_chart(
    PRICE_DATES, [PRICE_CLOSE, PRICE_MA50, PRICE_MA200],
    ["Close", "50-day MA", "200-day MA"],
    "Price history with 50-day and 200-day moving averages", y_label_fmt="${:,.0f}"))
story.append(Spacer(1, 8))
story.append(Paragraph(
    "[Describe the price trajectory over the relevant lookback — major moves, current position "
    "relative to 50-day/200-day moving averages, and whether the stock reads as trending or "
    "range-bound/searching for a bottom/top. Point at the chart above rather than restating it: "
    "note the crossover, the failed retest, the level that held.]", styles["Body"]))
story.append(Paragraph("5.2 Key technical levels", styles["H2"]))
story.append(data_table(
    ["Date", "Close", "Levels"],
    [["[date]", "$0", "[support/resistance/indicator readings from recent chart commentary]"]],
    [1.3 * inch, 0.85 * inch, 4.7 * inch],
))
story.append(Paragraph("5.3 Volatility &amp; options positioning", styles["H2"]))
story.append(bullets([
    "<b>Implied volatility:</b> [value and context vs. the stock's own historical range]",
    "<b>Put/call skew:</b> [value and what it implies about flow positioning]",
    "<b>Short interest:</b> [% of float and days-to-cover]",
    "<b>Ownership:</b> [institutional/insider % and notable holders or recent insider activity]",
]))
story.append(Paragraph("5.4 Relative performance", styles["H2"]))
story.append(Paragraph("[Performance vs. closest peer(s) and vs. a relevant index, with the framing analysts use for the comparison.]", styles["Body"]))
story.append(PageBreak())

# ============================================================================
# 6. CATALYST CALENDAR — TICKER-SPECIFIC
# ============================================================================
story.append(section_header("6. CATALYST CALENDAR"))
story.append(Spacer(1, 10))
story.append(Paragraph("6.1 Historical catalysts", styles["H2"]))
story.append(data_table(
    ["Date", "Event", "Reaction"],
    [["[date]", "[event]", "[stock reaction]"]],
    [1.15 * inch, 3.85 * inch, 1.85 * inch],
))
story.append(Spacer(1, 10))
story.append(Paragraph("6.2 Forward catalysts", styles["H2"]))
story.append(data_table(
    ["Date", "Event", "Why it matters"],
    [["[date]", "[event]", "[why it matters]"]],
    [1.15 * inch, 3.0 * inch, 2.7 * inch],
))
story.append(PageBreak())

# ============================================================================
# 7. KEY RISKS — TICKER-SPECIFIC
# ============================================================================
story.append(section_header("7. KEY RISKS"))
story.append(Spacer(1, 10))
story.append(bullets([
    "<b>[Risk name — the dominant one first].</b> [Specific mechanism, not generic boilerplate.]",
    "<b>[Risk name].</b> [Specific mechanism.]",
    "<b>[Risk name].</b> [Specific mechanism.]",
    "<b>Valuation risk.</b> [Where bulls/bears actually disagree and why — frame as an open question.]",
]))
story.append(PageBreak())

# ============================================================================
# 8. SOURCES & LIMITATIONS — always include, always specific
# ============================================================================
story.append(section_header("8. SOURCES &amp; LIMITATIONS"))
story.append(Spacer(1, 10))
story.append(Paragraph(
    f"This report was compiled {REPORT_DATE} from secondary public sources: [list categories — "
    "company press releases/IR site, SEC filings referenced via financial media, data "
    "aggregators, sell-side coverage summaries]. It was not built from a live Bloomberg or "
    "FactSet terminal.", styles["Body"]))
story.append(Paragraph("Data-quality flags — verify before relying on these specifically:", styles["H3"]))
story.append(bullets([
    "[Every figure that showed meaningful variance across sources — market cap, YTD return, exact dates, etc.]",
    "[Every date/figure that is an estimate rather than a confirmed, disclosed fact.]",
    "[Any individual investor's back-of-envelope estimate used as a cross-check, labeled as such — not sell-side or company guidance.]",
]))
story.append(Spacer(1, 10))
story.append(disclaimer_box(
    "<b>This is not investment advice.</b> This report is an educational/informational document "
    "produced with AI research assistance, not a licensed analyst's opinion, and should not be "
    "the sole basis for any investment decision. Do your own due diligence, verify load-bearing "
    "figures against primary sources, and consider consulting a licensed financial advisor "
    "before making any investment decision."
))


def add_page_number(c: canvas_mod.Canvas, doc):
    c.saveState()
    c.setFont("Helvetica", 8)
    c.setFillColor(GRAY)
    c.drawString(0.75 * inch, 0.5 * inch, f"{COMPANY_NAME} ({TICKER}) — Equity Research Report — {REPORT_DATE}")
    c.drawRightString(7.75 * inch, 0.5 * inch, f"Page {doc.page}")
    c.restoreState()


OUTPUT_PATH = f"{TICKER}_Equity_Research_Report.pdf"  # TICKER-SPECIFIC: real output filename
doc = SimpleDocTemplate(
    OUTPUT_PATH, pagesize=letter,
    topMargin=0.65 * inch, bottomMargin=0.75 * inch, leftMargin=0.75 * inch, rightMargin=0.75 * inch,
    title=f"{COMPANY_NAME} ({TICKER}) - Equity Research Report", author="AI Research Assistant",
)
doc.build(story, onFirstPage=add_page_number, onLaterPages=add_page_number)
print(f"Saved {OUTPUT_PATH} — now render a few pages to PNG and visually check before delivering.")

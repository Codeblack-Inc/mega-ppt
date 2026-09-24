# /// script
# requires-python = ">=3.10"
# dependencies = ["python-pptx>=1.0"]
# ///
"""deck.json → .pptx. Usage: uv run build_deck.py deck.json -o out.pptx

Layouts are plain functions registered in LAYOUTS. Spec: references/layouts.md.
"""
import argparse
import json
import sys

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

THEME = {
    "font": "Pretendard",
    "paper": "#FFFFFF",
    "ink": "#1B1F2A",
    "muted": "#6B7280",
    "rule": "#E5E7EB",
    "soft": "#F3F4F6",
    "accent": "#1E4FD8",
}
W, H = 13.333, 7.5  # 16:9 inches
M = 0.6  # side margin
BODY_TOP, BODY_BOTTOM = 1.75, 6.75

T = dict(THEME)


def rgb(hex_):
    return RGBColor.from_string(hex_.lstrip("#"))


def _font(run, size, color, bold=False):
    f = run.font
    f.size = Pt(size)
    f.bold = bold
    f.color.rgb = rgb(T[color] if color in T else color)
    f.name = T["font"]
    rpr = run._r.get_or_add_rPr()  # python-pptx only sets latin; Korean needs <a:ea>
    ea = rpr.find(qn("a:ea"))
    if ea is None:
        ea = rpr.makeelement(qn("a:ea"), {})
        rpr.append(ea)
    ea.set("typeface", T["font"])


def text(slide, x, y, w, h, lines, size=16, color="ink", bold=False,
         align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, bullet=False, gap=6):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for i, line in enumerate([lines] if isinstance(lines, str) else lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_after = Pt(gap)
        _font(p.add_run(), size, color, bold)
        p.runs[0].text = ("•  " if bullet else "") + str(line)
    return tb


def rect(slide, x, y, w, h, fill, shape=MSO_SHAPE.RECTANGLE):
    s = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    s.fill.solid()
    s.fill.fore_color.rgb = rgb(T[fill] if fill in T else fill)
    s.line.fill.background()
    s._element.remove(s._element.find(qn("p:style")))  # drop theme shadow/effects
    return s


def header(slide, d):
    """Action title + optional kicker. Every content slide gets this."""
    rect(slide, M, 0.42, 0.45, 0.06, "accent")
    if d.get("kicker"):
        text(slide, M + 0.6, 0.3, 8, 0.3, d["kicker"], 11, "muted", bold=True)
    text(slide, M, 0.62, W - 2 * M, 0.95, d.get("title", ""), 26, "ink", bold=True,
         anchor=MSO_ANCHOR.MIDDLE)


def footer(slide, d, n):
    if d.get("source"):
        text(slide, M, 7.0, 9, 0.3, "출처: " + d["source"], 9, "muted")
    text(slide, W - M - 1, 7.0, 1, 0.3, str(n), 9, "muted", align=PP_ALIGN.RIGHT)


# ---------------------------------------------------------------- layouts

def cover(s, d):
    rect(s, 0, 0, 0.25, H, "accent")
    text(s, 1.1, 2.3, 10.5, 1.8, d["title"], 40, "ink", bold=True, anchor=MSO_ANCHOR.BOTTOM)
    if d.get("subtitle"):
        text(s, 1.1, 4.3, 10.5, 0.8, d["subtitle"], 18, "muted")
    meta = " · ".join(x for x in (d.get("author"), d.get("date")) if x)
    if meta:
        text(s, 1.1, 6.4, 10.5, 0.4, meta, 12, "muted")


def section(s, d):
    rect(s, 0, 0, W, H, "ink")
    if d.get("no"):
        text(s, 1.1, 2.2, 4, 1, d["no"], 54, "accent", bold=True)
    text(s, 1.1, 3.3, 11, 1.2, d["title"], 34, "paper", bold=True)
    if d.get("subtitle"):
        text(s, 1.1, 4.5, 11, 0.8, d["subtitle"], 16, "rule")


def bullets(s, d):
    header(s, d)
    text(s, M, BODY_TOP + 0.1, W - 2 * M, BODY_BOTTOM - BODY_TOP, d["bullets"], 18,
         bullet=True, gap=14)


def two_column(s, d):
    header(s, d)
    cw = (W - 2 * M - 0.6) / 2
    for i, col in enumerate((d["left"], d["right"])):
        x = M + i * (cw + 0.6)
        hl = d.get("highlight") == i
        rect(s, x, BODY_TOP, cw, 0.06, "accent" if hl else "rule")
        text(s, x, BODY_TOP + 0.25, cw, 0.5, col.get("heading", ""), 18,
             "accent" if hl else "ink", bold=True)
        text(s, x, BODY_TOP + 0.95, cw, 3.8, col.get("bullets", []), 15, bullet=True, gap=10)


def kpi(s, d):
    header(s, d)
    items = d["items"]
    cw = (W - 2 * M - 0.4 * (len(items) - 1)) / len(items)
    for i, it in enumerate(items):
        x = M + i * (cw + 0.4)
        hl = d.get("highlight") == i
        rect(s, x, BODY_TOP + 0.4, cw, 2.9, "accent" if hl else "soft")
        fg = "paper" if hl else "ink"
        text(s, x + 0.3, BODY_TOP + 0.7, cw - 0.6, 1.3, it["value"], 44, fg, bold=True)
        text(s, x + 0.3, BODY_TOP + 2.0, cw - 0.6, 1.1, it["label"], 14,
             "paper" if hl else "muted")
    if d.get("takeaway"):
        text(s, M, 5.8, W - 2 * M, 0.8, d["takeaway"], 16, "ink")


CHART_TYPES = {
    "bar": XL_CHART_TYPE.COLUMN_CLUSTERED,
    "hbar": XL_CHART_TYPE.BAR_CLUSTERED,
    "stacked": XL_CHART_TYPE.COLUMN_STACKED,
    "line": XL_CHART_TYPE.LINE_MARKERS,
    "pie": XL_CHART_TYPE.DOUGHNUT,
}


def chart(s, d):
    header(s, d)
    c = d["chart"]
    data = CategoryChartData()
    data.categories = c["categories"]
    for ser in c["series"]:
        data.add_series(ser["name"], ser["values"])
    cw = W - 2 * M - (3.6 if d.get("takeaway") else 0)
    gf = s.shapes.add_chart(CHART_TYPES[c.get("type", "bar")], Inches(M), Inches(BODY_TOP),
                            Inches(cw), Inches(BODY_BOTTOM - BODY_TOP - 0.2), data)
    ch = gf.chart
    ch.font.size, ch.font.name = Pt(12), T["font"]
    ch.font.color.rgb = rgb(T["muted"])
    kind = c.get("type", "bar")
    multi = len(c["series"]) > 1 or kind == "pie"
    ch.has_title = bool(c.get("title"))
    if c.get("title"):
        ch.chart_title.text_frame.text = c["title"]
    ch.has_legend = multi
    if multi:
        ch.legend.position, ch.legend.include_in_layout = XL_LEGEND_POSITION.BOTTOM, False
    greys = ["#9CA3AF", "#D1D5DB", "#4B5563", "#E5E7EB"]
    hl = c.get("highlight")  # category index to accent (single-series) — the "look here"
    plot = ch.plots[0]
    if kind == "pie":
        for i, pt in enumerate(plot.series[0].points):
            pt.format.fill.solid()
            pt.format.fill.fore_color.rgb = rgb(T["accent"] if i == (hl or 0) else greys[i % 4])
    else:
        if kind != "line":
            plot.gap_width = 60
        va = ch.value_axis
        va.major_gridlines.format.line.color.rgb = rgb(T["rule"])
        va.format.line.fill.background()
        if c.get("labels", True):  # labels carry the numbers; axis is noise
            va.visible = False
            va.has_major_gridlines = False
        for i, ser in enumerate(plot.series):
            col = T["accent"] if i == 0 else greys[(i - 1) % 4]
            if kind == "line":
                ser.format.line.color.rgb = rgb(col)
                ser.format.line.width = Pt(2.5)
                ser.smooth = False
                continue
            ser.format.fill.solid()
            ser.format.fill.fore_color.rgb = rgb(col if hl is None or len(plot.series) > 1 else greys[0])
            if hl is not None and len(plot.series) == 1:
                pt = ser.points[hl]
                pt.format.fill.solid()
                pt.format.fill.fore_color.rgb = rgb(T["accent"])
        if c.get("labels", True):
            plot.has_data_labels = True
            plot.data_labels.font.size = Pt(11)
            if c.get("number_format"):
                plot.data_labels.number_format = c["number_format"]
                plot.data_labels.number_format_is_linked = False
    if d.get("takeaway"):
        x = W - M - 3.3
        rect(s, x, BODY_TOP, 0.06, 2.4, "accent")
        text(s, x + 0.3, BODY_TOP, 3.0, 4.5, d["takeaway"], 16, "ink")


def process(s, d):
    header(s, d)
    steps = d["steps"]
    n = len(steps)
    gap = 0.35
    cw = (W - 2 * M - gap * (n - 1)) / n
    for i, st in enumerate(steps):
        x = M + i * (cw + gap)
        hl = d.get("highlight") == i
        shp = rect(s, x, BODY_TOP + 0.3, cw, 0.9, "accent" if hl else "ink",
                   MSO_SHAPE.CHEVRON if 0 < i else MSO_SHAPE.PENTAGON)
        shp.adjustments[0] = 0.25
        text(s, x + 0.35, BODY_TOP + 0.3, cw - 0.7, 0.9, f"{i + 1:02d}", 16, "paper",
             bold=True, anchor=MSO_ANCHOR.MIDDLE)
        text(s, x, BODY_TOP + 1.5, cw, 0.6, st["title"], 17, "accent" if hl else "ink", bold=True)
        text(s, x, BODY_TOP + 2.15, cw, 2.6, st.get("body", ""), 13, "muted")


def table(s, d):
    header(s, d)
    rows = [d["columns"]] + d["rows"]
    nr, nc = len(rows), len(d["columns"])
    rh = min(0.6, (BODY_BOTTOM - BODY_TOP) / nr)
    tbl = s.shapes.add_table(nr, nc, Inches(M), Inches(BODY_TOP), Inches(W - 2 * M),
                             Inches(rh * nr)).table
    hl = d.get("highlight")  # row index (0 = first data row)
    for r, row in enumerate(rows):
        for c, val in enumerate(row):
            cell = tbl.cell(r, c)
            cell.fill.solid()
            is_hl = r > 0 and hl == r - 1
            cell.fill.fore_color.rgb = rgb(T["ink"] if r == 0 else T["soft"] if is_hl else T["paper"])
            tf = cell.text_frame
            tf.paragraphs[0].text = ""
            run = tf.paragraphs[0].add_run()
            _font(run, 13, "paper" if r == 0 else "accent" if is_hl else "ink", bold=r == 0 or is_hl)
            run.text = str(val)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE


def quote(s, d):
    rect(s, M + 0.5, 2.0, 0.08, 2.6, "accent")
    text(s, M + 1.0, 1.9, W - 2 * M - 1.5, 2.8, d["quote"], 30, "ink", bold=True,
         anchor=MSO_ANCHOR.MIDDLE)
    if d.get("by"):
        text(s, M + 1.0, 5.0, 10, 0.5, "— " + d["by"], 14, "muted")


def closing(s, d):
    rect(s, 0, 0, W, H, "ink")
    text(s, 1.1, 2.6, 11, 1.4, d.get("title", "감사합니다"), 40, "paper", bold=True,
         anchor=MSO_ANCHOR.BOTTOM)
    if d.get("subtitle"):
        text(s, 1.1, 4.2, 11, 1, d["subtitle"], 16, "rule")


LAYOUTS = {
    "cover": cover, "section": section, "bullets": bullets, "two-column": two_column,
    "kpi": kpi, "chart": chart, "process": process, "table": table, "quote": quote,
    "closing": closing,
}
NO_CHROME = {"cover", "section", "closing"}


def build(deck, out):
    T.clear()
    T.update(THEME, **deck.get("theme", {}))
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(W), Inches(H)
    for n, d in enumerate(deck["slides"], 1):
        lay = d.get("layout")
        if lay not in LAYOUTS:
            sys.exit(f"slide {n}: unknown layout {lay!r}. Choose from: {', '.join(LAYOUTS)}")
        s = prs.slides.add_slide(prs.slide_layouts[6])
        rect(s, 0, 0, W, H, "paper")
        LAYOUTS[lay](s, d)
        if lay not in NO_CHROME:
            footer(s, d, n)
        if d.get("notes"):
            s.notes_slide.notes_text_frame.text = d["notes"]
    prs.save(out)
    return len(prs.slides)


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("deck")
    ap.add_argument("-o", "--out", default="deck.pptx")
    a = ap.parse_args()
    with open(a.deck, encoding="utf-8") as f:
        n = build(json.load(f), a.out)
    print(f"wrote {a.out} ({n} slides)")

# /// script
# requires-python = ">=3.10"
# dependencies = ["python-pptx>=1.0", "pymupdf"]
# ///
"""deck.json → .pptx. Usage: uv run build_deck.py deck.json -o out.pptx

A content slide = header (kicker · headline · lead) + body grid of panels.
Panels are plain functions in PANELS; full-bleed slides in FULL. Spec: references/layouts.md.
"""
import argparse
import io
import json
import math
import re
import sys
from pathlib import Path

from lxml import etree
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
    "ink": "#14213D",
    "text": "#252C3A",
    "muted": "#5B6475",
    "rule": "#D5DAE3",
    "soft": "#F3F5F9",
    "accent": "#2152E0",
    "title_style": "line",  # panel titles: "line" | "bar"
}
W, H = 13.333, 7.5
M = 0.5          # side margin
GAP = 0.2        # grid gutter
BODY_BOTTOM = 6.85
BODY_SIZE = 10.5
MIN_SIZE = 8

T = dict(THEME)
CTX = {"slide": 0, "base": Path(".")}
WARN = []

ALIGN = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}
ANCHOR = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}
EMPH = re.compile(r"\*\*(.+?)\*\*")


# ------------------------------------------------------------------ primitives

def col(c):
    return T.get(c, c)


def rgb(c):
    return RGBColor.from_string(col(c).lstrip("#"))


def tint(c, a):
    """Mix colour c with white; a=0.9 → 90% white."""
    h = col(c).lstrip("#")
    return "#%02X%02X%02X" % tuple(round(v + (255 - v) * a) for v in
                                   (int(h[i:i + 2], 16) for i in (0, 2, 4)))


def warn(msg):
    WARN.append(f"slide {CTX['slide']}: {msg}")


def plain(t):
    return EMPH.sub(r"\1", str(t))


def _cw(ch):
    """Approx glyph advance in em. ponytail: width table, swap for real font metrics if fit misjudges."""
    o = ord(ch)
    if ch == " ":
        return 0.28
    if 0xAC00 <= o <= 0xD7A3 or 0x1100 <= o <= 0x11FF or 0x3130 <= o <= 0x318F \
            or 0x2E80 <= o <= 0x9FFF or 0xFF00 <= o <= 0xFFEF:
        return 0.94
    if ch.isupper() or ch in "mwMW%@&":
        return 0.66
    if ch.isdigit():
        return 0.57
    if ch in ".,:;!|il'\"()[]-·/":
        return 0.3
    return 0.53


def text_w(t, size, bold=False):
    return sum(_cw(c) for c in plain(t)) * size / 72 * (1.05 if bold else 1)


def n_lines(t, size, width, bold=False):
    maxw = width / (size / 72) / (1.05 if bold else 1)
    total = 0
    for para in plain(t).split("\n"):
        lines, cur = 1, 0.0
        for word in re.split(r"(\s+)", para):
            wl = sum(_cw(c) for c in word)
            if cur + wl > maxw and cur > 0 and not word.isspace():
                lines += 1
                cur = wl
            else:
                cur += wl
            while cur > maxw:  # word longer than the line
                lines += 1
                cur -= maxw
        total += lines
    return total


def _paras(content, bullets):
    items = [content] if isinstance(content, (str, int, float)) else content
    out = []
    for it in items:
        s = str(it)
        if bullets and s.startswith("- "):
            out.append((s[2:], 2))
        else:
            out.append((s, 1 if bullets else 0))
    return out


def _indent(size, lvl):
    return 0 if lvl == 0 else size / 72 * (1.1 if lvl == 1 else 2.3)


def text_height(content, size, w, bullets=False, bold=False, line=1.15, gap=4):
    ps = _paras(content, bullets)
    h = sum(n_lines(t, size - (1 if lvl == 2 else 0), w - _indent(size, lvl), bold)
            * size * line * 1.2 / 72 for t, lvl in ps)
    return h + gap / 72 * (len(ps) - 1)


def fit_size(content, size, w, h, bullets=False, bold=False, min_size=MIN_SIZE, line=1.15,
             gap=4, name="text", underfill=True):
    while size > min_size and text_height(content, size, w, bullets, bold, line, gap) > h:
        size -= 0.5
    need = text_height(content, size, w, bullets, bold, line, gap)
    if need > h * 1.04:
        warn(f"{name!r} overflows its box even at {size}pt — shorten the text")
    elif underfill and h > 1.5 and need < h * 0.3:
        warn(f"{name!r} fills only {need / h:.0%} of its box — add content or lower the row 'h'")
    return size


def _set_font(run, size, color, bold):
    f = run.font
    f.size = Pt(size)
    f.bold = bold
    f.color.rgb = rgb(color)
    f.name = T["font"]
    rpr = run._r.get_or_add_rPr()  # python-pptx sets only latin; Hangul needs <a:ea>
    ea = rpr.find(qn("a:ea"))
    if ea is None:
        ea = etree.SubElement(rpr, qn("a:ea"))
    ea.set("typeface", T["font"])


def _bullet(p, lvl, size, color):
    pPr = p._p.get_or_add_pPr()
    ind = int(size / 72 * 914400 * 1.1)
    pPr.set("marL", str(ind * lvl))
    pPr.set("indent", str(-ind))
    clr = etree.SubElement(pPr, qn("a:buClr"))
    etree.SubElement(clr, qn("a:srgbClr"), val=col(color).lstrip("#"))
    etree.SubElement(pPr, qn("a:buFont"), typeface="Arial")
    etree.SubElement(pPr, qn("a:buChar"), char="•" if lvl == 1 else "–")


def text(s, box, content, size=BODY_SIZE, color="text", bold=False, align="l", anchor="t",
         bullets=False, fit=True, min_size=MIN_SIZE, line=1.15, gap=4, emph="accent",
         name="text", wrap=True):
    """Text box with **emphasis** markup, hanging bullets ('- ' = level 2) and shrink-to-fit."""
    x, y, w, h = box
    if fit:
        size = fit_size(content, size, w, h, bullets, bold, min_size, line, gap, name,
                        underfill=anchor == "t")
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = ANCHOR[anchor]
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for i, (t, lvl) in enumerate(_paras(content, bullets)):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = ALIGN[align]
        p.line_spacing = line
        p.space_after = Pt(gap)
        sz = size - (1 if lvl == 2 else 0)
        if lvl:
            _bullet(p, lvl, size, "accent" if lvl == 1 else "muted")
        for j, seg in enumerate(t.split("\n")):
            if j:
                p.add_line_break()
            for k, part in enumerate(EMPH.split(seg)):
                if part:
                    r = p.add_run()
                    r.text = part
                    _set_font(r, sz, emph if k % 2 else color, bold or bool(k % 2))
    return tb


def rect(s, x, y, w, h, fill=None, line=None, shape=MSO_SHAPE.RECTANGLE, lw=0.75):
    sh = s.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    sh._element.remove(sh._element.find(qn("p:style")))  # no theme shadow/effects
    if fill:
        sh.fill.solid()
        sh.fill.fore_color.rgb = rgb(fill)
    else:
        sh.fill.background()
    if line:
        sh.line.color.rgb = rgb(line)
        sh.line.width = Pt(lw)
    else:
        sh.line.fill.background()
    return sh


def line(s, x1, y1, x2, y2, color="rule", lw=0.75, arrow=False, dash=False, arrow_start=False):
    c = s.shapes.add_connector(1, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    c._element.remove(c._element.find(qn("p:style")))
    c.line.color.rgb = rgb(color)
    c.line.width = Pt(lw)
    if dash:
        c.line.dash_style = 4
    ln = c.line._get_or_add_ln()
    if arrow_start:  # schema order: headEnd before tailEnd
        etree.SubElement(ln, qn("a:headEnd"), type="triangle", w="med", len="med")
    if arrow:
        etree.SubElement(ln, qn("a:tailEnd"), type="triangle", w="med", len="med")
    return c


def hline(s, x, y, w, color="rule", lw=0.75):
    return line(s, x, y, x + w, y, color, lw)


def shape_text(sh, t, size, color, bold=True, align="c", box=None):
    """Text inside a shape. box=(w, h) usable area → shrink to fit (never wraps mid-word)."""
    if box:
        words = max((wd for ln in plain(t).split("\n") for wd in ln.split()), key=len, default="")
        while size > 7 and (text_w(words, size, bold) > box[0] * 0.85 or  # margin for fallback fonts
                            text_height(t, size, box[0], bold=bold, line=1.0) > box[1]):
            size -= 0.5
    tf = sh.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(0.05)
    tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = ALIGN[align]
    for j, seg in enumerate(str(t).split("\n")):
        if j:
            p.add_line_break()
        r = p.add_run()
        r.text = plain(seg)
        _set_font(r, size, color, bold)


# ------------------------------------------------------------------ panels
# Each panel: f(slide, (x, y, w, h), spec). Optional spec["title"] is drawn by panel().

def p_bullets(s, b, p):
    x, y, w, h = b
    text(s, b, p["items"], p.get("size", BODY_SIZE), bullets=p.get("style") != "plain",
         gap=p.get("gap", 5), name=p.get("title", "bullets"))


def p_callout(s, b, p):
    x, y, w, h = b
    dark = p.get("style") == "dark"
    rect(s, x, y, w, h, "ink" if dark else tint("accent", 0.9))
    if not dark:
        rect(s, x, y, 0.06, h, "accent")
    text(s, (x + 0.25, y + 0.08, w - 0.4, h - 0.16), p["text"], p.get("size", 12.5),
         "paper" if dark else "ink", bold=True, anchor="m", align=p.get("align", "l"),
         emph="#9DB6FF" if dark else "accent", name="callout")


def p_table(s, b, p):
    x, y, w, h = b
    cols, rows = p.get("columns"), p["rows"]
    data = ([cols] if cols else []) + rows
    nr, nc = len(data), max(len(r) for r in data)
    rel = p.get("widths") or [1] * nc
    cws = [w * r / sum(rel) for r in rel]
    size = p.get("size", 9.5)
    rowhead = p.get("rowhead", False)

    def heights(sz):
        return [max(n_lines(c, sz, cw - 0.16, bold=(r == 0 and cols) or (rowhead and i == 0))
                    for i, (c, cw) in enumerate(zip(row, cws))) * sz * 1.25 / 72 + 0.14
                for r, row in enumerate(data)]
    rh = heights(size)
    while sum(rh) > h and size > 7.5:
        size -= 0.5
        rh = heights(size)
    if sum(rh) > h * 1.04:
        warn(f"table {p.get('title', '')!r} too tall ({sum(rh):.1f}in > {h:.1f}in) — cut rows")
    if p.get("stretch", True) and sum(rh) < h:  # fill the panel so grids line up
        k = min(h / sum(rh), 0.6 * nr / sum(rh))  # but keep rows ≤ ~0.6in on average
        rh = [r * max(k, 1) for r in rh]
    gf = s.shapes.add_table(nr, nc, Inches(x), Inches(y), Inches(w), Inches(sum(rh)))
    tbl = gf.table
    tblPr = tbl._tbl.tblPr
    tblPr.set("firstRow", "0")
    tblPr.set("bandRow", "0")
    sid = tblPr.find(qn("a:tableStyleId"))
    if sid is None:
        sid = etree.SubElement(tblPr, qn("a:tableStyleId"))
    sid.text = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"  # "No Style, No Grid"
    for i, cw in enumerate(cws):
        tbl.columns[i].width = Inches(cw)
    hl = p.get("highlight")
    total = p.get("total", False)
    body = data[1:] if cols else data
    left = [any(len(plain(r[c])) > 14 for r in body if c < len(r)) for c in range(nc)]
    for r, row in enumerate(data):
        tbl.rows[r].height = Inches(rh[r])
        head = bool(cols) and r == 0
        is_hl = hl is not None and r - (1 if cols else 0) == hl
        for c in range(nc):
            val = str(row[c]) if c < len(row) else ""
            cell = tbl.cell(r, c)
            rh_cell = rowhead and c == 0 and not head
            last = r == nr - 1
            is_total = total and last
            fill = ("ink" if head else tint("accent", 0.9) if is_hl
                    else "soft" if rh_cell or is_total else "paper")
            _cell_borders(cell, T=("ink" if (r == 0 and not cols) or is_total else None),
                          B=("ink" if last else "rule"),
                          B_w=1.25 if last else 0.5,
                          R=("rule" if rh_cell else None))
            cell.fill.solid()
            cell.fill.fore_color.rgb = rgb(fill)
            cell.margin_left = cell.margin_right = Inches(0.08)
            cell.margin_top = cell.margin_bottom = Inches(0.04)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf = cell.text_frame
            tf.word_wrap = True
            para = tf.paragraphs[0]
            para.alignment = (PP_ALIGN.LEFT if (rowhead and c == 0) or (left[c] and not head)
                              else PP_ALIGN.CENTER)
            color = ("paper" if head else "accent" if is_hl
                     else "ink" if rh_cell or is_total else "text")
            mark = MARKS.get(plain(val).strip()) if not head else None
            if mark:
                para.alignment = PP_ALIGN.CENTER
                run = para.add_run()
                run.text = plain(val).strip()
                _set_font(run, size + 8, mark, False)
                continue
            for k, part in enumerate(EMPH.split(val)):
                if part:
                    run = para.add_run()
                    run.text = part
                    _set_font(run, size, "accent" if k % 2 and not head else color,
                              head or rh_cell or is_hl or is_total or bool(k % 2))
    for r1, c1, r2, c2 in p.get("merge", []):  # data coords, header row = 0 when columns given
        tbl.cell(r1, c1).merge(tbl.cell(r2, c2))


MARKS = {"●": "accent", "◐": "accent", "○": "muted", "✓": "accent", "✗": "muted",
         "△": "muted", "O": "accent", "X": "muted"}


def _cell_borders(cell, L=None, R=None, T=None, B=None, B_w=0.5):
    tcPr = cell._tc.get_or_add_tcPr()
    for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
        old = tcPr.find(qn(tag))
        if old is not None:
            tcPr.remove(old)
    for i, (tag, c, wpt) in enumerate((("a:lnL", L, 0.5), ("a:lnR", R, 0.5),
                                       ("a:lnT", T, 1.25), ("a:lnB", B, B_w))):
        ln = etree.Element(qn(tag), w=str(int(wpt * 12700)))
        if c:
            sf = etree.SubElement(ln, qn("a:solidFill"))
            etree.SubElement(sf, qn("a:srgbClr"), val=col(c).lstrip("#"))
        else:
            etree.SubElement(ln, qn("a:noFill"))
        tcPr.insert(i, ln)  # schema order: lnL lnR lnT lnB, then fill


CHART_TYPES = {
    "bar": XL_CHART_TYPE.COLUMN_CLUSTERED, "hbar": XL_CHART_TYPE.BAR_CLUSTERED,
    "stacked": XL_CHART_TYPE.COLUMN_STACKED, "line": XL_CHART_TYPE.LINE_MARKERS,
    "pie": XL_CHART_TYPE.DOUGHNUT,
}
GREYS = ["#A9B1BF", "#CDD2DB", "#6B7485", "#E3E6EC"]


def p_chart(s, b, p):
    x, y, w, h = b
    if p.get("unit"):
        text(s, (x, y, w, 0.2), f"(단위: {p['unit']})", 8, "muted", align="r", fit=False)
        y, h = y + 0.2, h - 0.2
    kind = p.get("kind", "bar")
    if kind == "waterfall":
        return _waterfall(s, (x, y, w, h), p)
    if kind == "combo":
        return _combo(s, (x, y, w, h), p)
    data = CategoryChartData()
    data.categories = p["categories"]
    for ser in p["series"]:
        data.add_series(ser["name"], ser["values"])
    ch = s.shapes.add_chart(CHART_TYPES[kind], Inches(x), Inches(y), Inches(w), Inches(h),
                            data).chart
    ch.font.size, ch.font.name = Pt(9), T["font"]
    ch.font.color.rgb = rgb("muted")
    ch.has_title = False
    multi = len(p["series"]) > 1 or kind == "pie"
    ch.has_legend = multi
    if multi:
        ch.legend.position, ch.legend.include_in_layout = XL_LEGEND_POSITION.BOTTOM, False
        ch.legend.font.size = Pt(9)
    hl = p.get("highlight")
    plot = ch.plots[0]
    if kind == "pie":
        plot.has_data_labels = True
        plot.data_labels.number_format = p.get("number_format", "0%")
        plot.data_labels.number_format_is_linked = False
        pie_greys = ["#AEB6C4", "#CDD2DB", "#E3E6EC", "#98A1B1"]
        for i, pt in enumerate(plot.series[0].points):
            on = i == (hl or 0)
            pt.format.fill.solid()
            pt.format.fill.fore_color.rgb = rgb("accent" if on else pie_greys[i % 4])
            f = pt.data_label.font
            f.size, f.bold = Pt(10), True
            f.color.rgb = rgb("paper" if on else "ink")
            dl = pt.data_label._dLbl  # per-point label drops the plot numFmt; restate it
            dl.insert(dl.index(dl.find(qn("c:txPr"))), etree.Element(
                qn("c:numFmt"), formatCode=p.get("number_format", "0%"), sourceLinked="0"))
        return
    if kind != "line":
        plot.gap_width = 70
        if kind == "stacked":
            plot.overlap = 100
    va = ch.value_axis
    va.major_gridlines.format.line.color.rgb = rgb("rule")
    va.format.line.fill.background()
    va.tick_labels.font.size = Pt(8)
    ch.category_axis.format.line.color.rgb = rgb("rule")
    ch.category_axis.tick_labels.font.size = Pt(9)
    if kind == "hbar":
        ch.category_axis.reverse_order = True  # first category on top, like the source table
    multiline = kind == "line" and len(p["series"]) > 1
    labels = p.get("labels", True) and not multiline
    if labels:
        va.visible = False
        va.has_major_gridlines = False
    single = len(plot.series) == 1
    for i, ser in enumerate(plot.series):
        c = "accent" if i == 0 else GREYS[(i - 1) % 4]
        if kind == "line":
            ser.format.line.color.rgb = rgb(c)
            ser.format.line.width = Pt(2.25)
            ser.smooth = False
            ser.marker.format.fill.solid()
            ser.marker.format.fill.fore_color.rgb = rgb(c)
            ser.marker.format.line.color.rgb = rgb(c)
            if multiline:  # label only the end point; overlapping labels are unreadable
                dl = ser.points[len(p["categories"]) - 1].data_label
                dl.has_text_frame = False
                dl.show_value = True
                dl.font.size, dl.font.bold = Pt(9), True
                dl.font.color.rgb = rgb(c)
            continue
        ser.format.fill.solid()
        ser.format.fill.fore_color.rgb = rgb(GREYS[0] if single and hl is not None else c)
        if single and hl is not None:
            pt = ser.points[hl]
            pt.format.fill.solid()
            pt.format.fill.fore_color.rgb = rgb("accent")
    if labels:
        plot.has_data_labels = True
        plot.data_labels.font.size = Pt(9)
        plot.data_labels.font.color.rgb = rgb("text")
        if p.get("number_format"):
            plot.data_labels.number_format = p["number_format"]
            plot.data_labels.number_format_is_linked = False


def _combo(s, b, p):
    """Bars on the primary axis + series with "type": "line" on a hidden secondary axis."""
    from pptx.oxml import parse_xml
    from pptx.oxml.ns import nsdecls
    x, y, w, h = b
    bars = [sr for sr in p["series"] if sr.get("type", "bar") != "line"]
    lines = [sr for sr in p["series"] if sr.get("type") == "line"]
    data = CategoryChartData()
    data.categories = p["categories"]
    for sr in bars + lines:
        data.add_series(sr["name"], sr["values"])
    ch = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(x), Inches(y), Inches(w),
                            Inches(h), data).chart
    plot_area = ch._chartSpace.chart.plotArea
    bar = plot_area.find(qn("c:barChart"))
    lc = parse_xml(f'<c:lineChart {nsdecls("c")}><c:grouping val="standard"/>'
                   '<c:varyColors val="0"/></c:lineChart>')
    for ser in bar.findall(qn("c:ser"))[len(bars):]:
        bar.remove(ser)
        inv = ser.find(qn("c:invertIfNegative"))
        if inv is not None:
            ser.remove(inv)
        etree.SubElement(ser, qn("c:smooth"), val="0")
        lc.append(ser)
    for tag, v in (("c:marker", "1"), ("c:axId", "50000"), ("c:axId", "50001")):
        etree.SubElement(lc, qn(tag), val=v)
    bar.addnext(lc)
    axes = [e for e in plot_area if e.tag in (qn("c:valAx"), qn("c:catAx"))]
    axes[-1].addnext(parse_xml(
        f'<c:catAx {nsdecls("c")}><c:axId val="50001"/><c:scaling><c:orientation val="minMax"/>'
        '</c:scaling><c:delete val="1"/><c:axPos val="b"/><c:majorTickMark val="none"/>'
        '<c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="50000"/>'
        '<c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/>'
        '<c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>'))
    axes[-1].addnext(parse_xml(
        f'<c:valAx {nsdecls("c")}><c:axId val="50000"/><c:scaling><c:orientation val="minMax"/>'
        '</c:scaling><c:delete val="0"/><c:axPos val="r"/><c:numFmt formatCode="General" '
        'sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/>'
        '<c:tickLblPos val="none"/><c:spPr><a:ln><a:noFill/></a:ln></c:spPr>'
        '<c:crossAx val="50001"/><c:crosses val="max"/><c:crossBetween val="between"/></c:valAx>'
        .replace("<c:valAx ", f'<c:valAx {nsdecls("a")} ')))
    ch.font.size, ch.font.name = Pt(9), T["font"]
    ch.font.color.rgb = rgb("muted")
    ch.has_title = False
    ch.has_legend = True
    ch.legend.position, ch.legend.include_in_layout = XL_LEGEND_POSITION.BOTTOM, False
    ch.legend.font.size = Pt(9)
    for va in plot_area.findall(qn("c:valAx")):  # by id: python-pptx may pick the secondary
        secondary = va.find(qn("c:axId")).get("val") == "50000"
        va.find(qn("c:delete")).set("val", "0" if secondary else "1")
        grid = va.find(qn("c:majorGridlines"))
        if grid is not None:
            va.remove(grid)
    ch.category_axis.format.line.color.rgb = rgb("rule")
    ch.category_axis.tick_labels.font.size = Pt(9)
    bar_plot, line_plot = ch.plots[0], ch.plots[1]
    bar_plot.gap_width = 60
    for i, (ser, spec) in enumerate(zip(bar_plot.series, bars)):
        ser.format.fill.solid()
        ser.format.fill.fore_color.rgb = rgb(tint("ink", 0.35) if i == 0 else GREYS[i % 4])
    for ser, spec in zip(line_plot.series, lines):
        ser.format.line.color.rgb = rgb("accent")
        ser.format.line.width = Pt(2.5)
        ser.marker.format.fill.solid()
        ser.marker.format.fill.fore_color.rgb = rgb("accent")
        ser.marker.format.line.color.rgb = rgb("paper")
    for plot, specs, color in ((bar_plot, bars, "text"), (line_plot, lines, "accent")):
        for ser, spec in zip(plot.series, specs):
            dl = ser.data_labels
            dl.show_value = True
            dl.font.size, dl.font.bold = Pt(9), plot is line_plot
            dl.font.color.rgb = rgb(color)
            if spec.get("number_format"):
                dl.number_format, dl.number_format_is_linked = spec["number_format"], False
            if plot is line_plot:
                from pptx.enum.chart import XL_LABEL_POSITION
                dl.position = XL_LABEL_POSITION.ABOVE


def _waterfall(s, b, p):
    """Bridge chart as stacked columns with an invisible base. p["totals"] = absolute bars."""
    x, y, w, h = b
    vals, totals = p["series"][0]["values"], set(p.get("totals", []))
    base, up, down, run = [], [], [], 0
    for i, v in enumerate(vals):
        if i in totals:
            base.append(0), up.append(v), down.append(0)
            run = v
        elif v >= 0:
            base.append(run), up.append(v), down.append(0)
            run += v
        else:
            run += v
            base.append(run), up.append(0), down.append(-v)
    data = CategoryChartData()
    data.categories = p["categories"]
    for name, vs in (("base", base), ("증가", up), ("감소", down)):
        data.add_series(name, vs)
    ch = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_STACKED, Inches(x), Inches(y), Inches(w),
                            Inches(h), data).chart
    ch.font.size, ch.font.name = Pt(9), T["font"]
    ch.font.color.rgb = rgb("muted")
    ch.has_title = ch.has_legend = False
    plot = ch.plots[0]
    plot.gap_width, plot.overlap = 50, 100
    ch.value_axis.visible = False
    ch.value_axis.has_major_gridlines = False
    ch.category_axis.format.line.color.rgb = rgb("rule")
    ch.category_axis.tick_labels.font.size = Pt(9)
    nf = p.get("number_format", "#,##0")
    b_ser, u_ser, d_ser = plot.series
    b_ser.format.fill.background()
    for ser, c, fmt in ((u_ser, tint("ink", 0.4), f"{nf};-{nf};;"),
                        (d_ser, "#E07A6E", f'"-"{nf};;;')):
        ser.format.fill.solid()
        ser.format.fill.fore_color.rgb = rgb(c)
        dl = ser.data_labels
        dl.show_value = True
        dl.number_format, dl.number_format_is_linked = fmt, False
        dl.font.size, dl.font.bold = Pt(9), True
        dl.font.color.rgb = rgb("text")
    for i in totals:
        pt = u_ser.points[i]
        pt.format.fill.solid()
        pt.format.fill.fore_color.rgb = rgb("accent")


def p_kpi(s, b, p):
    x, y, w, h = b
    items = p["items"]
    n = len(items)
    vertical = p.get("direction") == "column"
    g = 0.12
    cw = w if vertical else (w - g * (n - 1)) / n
    chh = (h - g * (n - 1)) / n if vertical else h
    for i, it in enumerate(items):
        cx, cy = (x, y + i * (chh + g)) if vertical else (x + i * (cw + g), y)
        hl = p.get("highlight") == i
        rect(s, cx, cy, cw, chh, "soft")
        rect(s, cx, cy, cw, 0.05, "accent" if hl else "rule")
        vs = 30
        while vs > 14 and (text_w(it["value"], vs, True) > (cw - 0.35) * 0.85 or vs / 72 * 1.3 > chh * 0.5):
            vs -= 1
        text(s, (cx + 0.15, cy + 0.15, cw - 0.3, chh * 0.5), it["value"], vs,
             "accent" if hl else "ink", bold=True, anchor="b", fit=False, wrap=False)
        text(s, (cx + 0.15, cy + 0.2 + chh * 0.5, cw - 0.3, chh * 0.5 - 0.3), it["label"], 10,
             "muted", name="kpi label")


def p_cards(s, b, p):
    x, y, w, h = b
    items = p["items"]
    ncol = p.get("cols", len(items))
    nrow = math.ceil(len(items) / ncol)
    g = 0.15
    cw = (w - g * (ncol - 1)) / ncol
    chh = (h - g * (nrow - 1)) / nrow
    numbered = p.get("numbered", True)
    for i, it in enumerate(items):
        cx, cy = x + (i % ncol) * (cw + g), y + (i // ncol) * (chh + g)
        hl = p.get("highlight") == i
        th = min(0.5, max(0.36, text_height(it["title"], 11, cw - 0.6, bold=True) + 0.14))
        rect(s, cx, cy, cw, chh, "paper", "accent" if hl else "rule", lw=1 if hl else 0.75)
        rect(s, cx, cy, cw, th, "accent" if hl else "soft")
        tx = cx + 0.12
        if it.get("icon"):
            icon(s, it["icon"], cx + 0.11, cy + th / 2 - 0.13, 0.26, "paper" if hl else "accent")
            tx = cx + 0.46
        elif numbered:
            badge = rect(s, cx + 0.1, cy + th / 2 - 0.12, 0.24, 0.24, "paper" if hl else "accent",
                         shape=MSO_SHAPE.OVAL)
            shape_text(badge, i + 1, 8.5, "accent" if hl else "paper")
            tx = cx + 0.42
        text(s, (tx, cy, cx + cw - tx - 0.08, th), it["title"], 11, "paper" if hl else "ink",
             bold=True, anchor="m", name="card title")
        body = it.get("items") or it.get("text")
        if body:
            text(s, (cx + 0.14, cy + th + 0.12, cw - 0.28, chh - th - 0.2), body,
                 p.get("size", 10), bullets="items" in it, gap=3, name=f"card {it['title']}")


def p_process(s, b, p):
    x, y, w, h = b
    steps = p["steps"]
    n = len(steps)
    g = 0.08
    cw = (w - g * (n - 1)) / n
    ah = 0.5
    for i, st in enumerate(steps):
        cx = x + i * (cw + g)
        hl = p.get("highlight") == i
        sh = rect(s, cx, y, cw, ah, "accent" if hl else "ink",
                  shape=MSO_SHAPE.PENTAGON if i == 0 else MSO_SHAPE.CHEVRON)
        sh.adjustments[0] = 0.3
        shape_text(sh, st["title"], 11, "paper")
        if st.get("label"):
            text(s, (cx + 0.2, y + ah + 0.1, cw - 0.3, 0.25), st["label"], 9, "accent", bold=True,
                 fit=False)
        body = st.get("items") or st.get("text")
        if body:
            top = y + ah + (0.38 if st.get("label") else 0.15)
            text(s, (cx + 0.2, top, cw - 0.3, y + h - top), body, p.get("size", 10),
                 bullets="items" in st, gap=3, name=f"step {st['title']}")


def p_timeline(s, b, p):
    x, y, w, h = b
    periods, tasks = p["periods"], p["tasks"]
    lw = p.get("label_width", 1.8)
    gw = (w - lw) / len(periods)
    hh = 0.32
    rect(s, x, y, w, hh, "ink")
    text(s, (x + 0.1, y, lw - 0.1, hh), p.get("label", "구분"), 9.5, "paper", bold=True,
         anchor="m", fit=False)
    for i, per in enumerate(periods):
        text(s, (x + lw + i * gw, y, gw, hh), per, 9, "paper", bold=True, align="c", anchor="m",
             fit=False)
    rh = min(0.42, (h - hh) / len(tasks))
    for i in range(len(periods) + 1):
        c = s.shapes.add_connector(1, Inches(x + lw + i * gw), Inches(y + hh),
                                   Inches(x + lw + i * gw), Inches(y + hh + rh * len(tasks)))
        c._element.remove(c._element.find(qn("p:style")))
        c.line.color.rgb = rgb("rule")
        c.line.width = Pt(0.5)
    for r, t in enumerate(tasks):
        ry = y + hh + r * rh
        if r % 2 == 0:
            rect(s, x, ry, lw, rh, "soft")
        hline(s, x, ry + rh, w)
        text(s, (x + 0.1, ry, lw - 0.15, rh), t["name"], 9.5, "ink", bold=bool(t.get("group")),
             anchor="m", name="task")
        spans = t["span"] if isinstance(t["span"][0], list) else [t["span"]]
        for a, z in spans:  # 1-based inclusive period indices
            bx = x + lw + (a - 1) * gw + 0.05
            bar = rect(s, bx, ry + rh * 0.28, (z - a + 1) * gw - 0.1, rh * 0.44,
                       "accent" if t.get("highlight") else tint("ink", 0.35),
                       shape=MSO_SHAPE.ROUNDED_RECTANGLE)
            bar.adjustments[0] = 0.5
            if t.get("note"):
                shape_text(bar, t["note"], 8, "paper", bold=False)


def p_cycle(s, b, p):
    x, y, w, h = b
    nodes = p["nodes"]
    n = len(nodes)
    cx, cy = x + w / 2, y + h / 2
    R = min(w, h) / 2 - 0.45
    nr = min(0.48, math.pi * R / n * 0.8)
    ring = rect(s, cx - R, cy - R, 2 * R, 2 * R, None, "rule", MSO_SHAPE.OVAL, lw=1.5)
    ring.line.dash_style = 4  # dash
    c = rect(s, cx - R * 0.5, cy - R * 0.5, R, R, "accent", shape=MSO_SHAPE.OVAL)
    text(s, (cx - R * 0.4, cy - R * 0.3, R * 0.8, R * 0.6), p.get("center", ""), 12, "paper",
         bold=True, align="c", anchor="m", line=1.0, gap=0, min_size=8, name="cycle center")
    for i, nd in enumerate(nodes):
        a = -math.pi / 2 + 2 * math.pi * i / n
        nx, ny = cx + R * math.cos(a), cy + R * math.sin(a)
        hl = p.get("highlight") == i
        o = rect(s, nx - nr, ny - nr, 2 * nr, 2 * nr, "accent" if hl else "paper",
                 "accent", MSO_SHAPE.OVAL, lw=1.25)
        tw = nr * 1.6  # separate text box: ellipse text insets vary by renderer
        text(s, (nx - tw / 2, ny - nr * 0.6, tw, nr * 1.2), nd, 10, "paper" if hl else "ink",
             bold=True, align="c", anchor="m", line=1.0, gap=0, min_size=7, name="cycle node")


def p_stack(s, b, p):
    x, y, w, h = b
    layers = p["layers"]
    g = 0.08
    lh = (h - g * (len(layers) - 1)) / len(layers)
    lw = p.get("label_width", 1.3)
    for i, ly in enumerate(layers):
        ly_y = y + i * (lh + g)
        hl = ly.get("highlight")
        lab = rect(s, x, ly_y, lw, lh, "accent" if hl else "ink")
        shape_text(lab, ly["label"], 10, "paper", box=(lw - 0.12, lh - 0.1))
        rect(s, x + lw, ly_y, w - lw, lh, tint("accent", 0.92) if hl else "soft")
        items = ly["items"]
        ig = 0.1
        iw = (w - lw - 0.2 - ig * (len(items) - 1)) / len(items)
        for j, it in enumerate(items):
            bx = rect(s, x + lw + 0.1 + j * (iw + ig), ly_y + 0.1, iw, lh - 0.2, "paper",
                      "accent" if hl else "rule")
            shape_text(bx, it, 9.5, "ink", bold=False, box=(iw - 0.12, lh - 0.24))


def p_label(s, b, p):
    x, y, w, h = b
    lab = rect(s, x, y, w, h, "accent" if p.get("highlight") else "ink")
    t = p["text"]
    shape_text(lab, "\n".join(t) if h > w * 2 and " " not in t else t, 10.5, "paper")


def p_arrow(s, b, p):
    x, y, w, h = b
    down = p.get("direction") == "down"
    aw, ah = (0.5, 0.26) if down else (0.26, 0.5)
    a = rect(s, x + (w - aw) / 2, y + (h - ah) / 2, aw, ah, tint("accent", 0.35),
             shape=MSO_SHAPE.DOWN_ARROW if down else MSO_SHAPE.RIGHT_ARROW)
    return a


def _rotated_label(s, cx, cy, length, t):
    tb = text(s, (cx - length / 2, cy - 0.13, length, 0.26), t, 9.5, "muted", bold=True,
              align="c", anchor="m", fit=False)
    tb.rotation = 270


def p_matrix(s, b, p):
    """2×2: quadrants (SWOT 등) or positioning map with points."""
    x, y, w, h = b
    ax = 0.35 if p.get("y_label") else 0
    ay = 0.32 if p.get("x_label") else 0
    gx, gy, gw, gh = x + ax, y, w - ax, h - ay
    if p.get("y_label"):
        _rotated_label(s, x + 0.15, gy + gh / 2, gh, p["y_label"])
    if p.get("x_label"):
        text(s, (gx, y + h - 0.26, gw, 0.26), p["x_label"], 9.5, "muted", bold=True, align="c",
             anchor="m", fit=False)
    if p.get("points"):
        return _positioning(s, (gx, gy, gw, gh), p)
    g = 0.12
    qw, qh = (gw - g) / 2, (gh - g) / 2
    letters = p.get("letters")
    for i, q in enumerate(p["quadrants"][:4]):
        qx, qy = gx + (i % 2) * (qw + g), gy + (i // 2) * (qh + g)
        hl = p.get("highlight") == i
        rect(s, qx, qy, qw, qh, tint("accent", 0.92) if hl else "soft")
        if hl:
            rect(s, qx, qy, qw, 0.05, "accent")
        tx = qx + 0.18
        if letters:
            lb = rect(s, qx + 0.15, qy + 0.15, 0.46, 0.46, "accent" if hl else "ink")
            shape_text(lb, letters[i], 18, "paper")
            tx = qx + 0.75
        text(s, (tx, qy + 0.15, qx + qw - tx - 0.15, 0.46), q["title"], 12,
             "accent" if hl else "ink", bold=True, anchor="m", name="quadrant title")
        body = q.get("items") or q.get("text")
        if body:
            text(s, (qx + 0.2, qy + 0.75, qw - 0.4, qh - 0.9), body, p.get("size", 10),
                 bullets="items" in q, gap=3, name=q["title"])


def _positioning(s, b, p):
    x, y, w, h = b
    hl = p.get("highlight")
    for i in range(4):
        rect(s, x + (i % 2) * w / 2, y + (i // 2) * h / 2, w / 2, h / 2,
             tint("accent", 0.92) if hl == i else "soft" if i in (0, 3) else "paper")
    line(s, x + w / 2, y, x + w / 2, y + h, "rule", 1, dash=True)
    line(s, x, y + h / 2, x + w, y + h / 2, "rule", 1, dash=True)
    line(s, x, y + h, x + w, y + h, "muted", 1.25, arrow=True)
    line(s, x, y + h, x, y, "muted", 1.25, arrow=True)
    for i, q in enumerate(p.get("quadrant_labels", [])[:4]):
        qx, qy = x + (i % 2) * w / 2, y + (i // 2) * h / 2
        text(s, (qx + 0.12, qy + 0.08, w / 2 - 0.24, 0.26), q, 9.5,
             "accent" if hl == i else "muted", bold=True, align="r" if i % 2 else "l",
             fit=False)
    for pt in p["points"]:
        on = pt.get("highlight")
        r = 0.13 if on else 0.09
        cx, cy = x + pt["x"] * w, y + (1 - pt["y"]) * h
        rect(s, cx - r, cy - r, 2 * r, 2 * r, "accent" if on else tint("ink", 0.35),
             shape=MSO_SHAPE.OVAL)
        right = pt["x"] < 0.75
        lx = cx + r + 0.06 if right else cx - r - 0.06 - 1.9
        text(s, (lx, cy - 0.13, 1.9, 0.26), pt["label"], 10 if on else 9.5,
             "accent" if on else "ink", bold=bool(on), align="l" if right else "r", anchor="m",
             fit=False)


def p_pyramid(s, b, p, funnel=False):
    x, y, w, h = b
    levels = p["levels"]
    n = len(levels)
    has_text = any(lv.get("text") for lv in levels)
    sw = w * 0.5 if has_text else w
    g = 0.06
    lh = (h - g * (n - 1)) / n
    for i, lv in enumerate(levels):
        t = i / (n - 1) if n > 1 else 1
        frac = 1 - 0.6 * t if funnel else 0.32 + 0.68 * t
        bw = sw * frac
        bx, by = x + (sw - bw) / 2, y + i * (lh + g)
        hl = p.get("highlight") == i
        shade = min(0.7, (i if not funnel else i) * 0.16)
        fill = "accent" if hl else tint("ink", shade)
        sh = rect(s, bx, by, bw, lh, fill)
        label = lv["title"] + (f"\n{lv['value']}" if lv.get("value") else "")
        shape_text(sh, label, 11.5, "paper" if hl or shade < 0.45 else "ink",
                   box=(bw - 0.1, lh - 0.04))
        if lv.get("text"):
            line(s, bx + bw + 0.06, by + lh / 2, x + sw + 0.18, by + lh / 2, "rule", 0.75,
                 dash=True)
            text(s, (x + sw + 0.28, by, w - sw - 0.28, lh), lv["text"], 10, anchor="m",
                 name=lv["title"])


def p_funnel(s, b, p):
    p_pyramid(s, b, p, funnel=True)


def _node(s, x, y, w, h, d, fill, fg, border=None):
    rect(s, x, y, w, h, fill, border)
    name, role = (d, None) if isinstance(d, str) else (d["name"], d.get("role"))
    if role:
        text(s, (x + 0.06, y + 0.04, w - 0.12, h * 0.55 - 0.04), name, 10.5, fg, bold=True,
             align="c", anchor="b", min_size=7.5, line=1.0, name="node")
        text(s, (x + 0.06, y + h * 0.55, w - 0.12, h * 0.45 - 0.04), role, 8.5,
             fg if fg != "ink" else "muted", align="c", anchor="t", min_size=7, line=1.0,
             name="node role")
    else:
        text(s, (x + 0.06, y, w - 0.12, h), name, 10.5, fg, bold=True, align="c", anchor="m",
             min_size=7.5, line=1.0, name="node")


def p_tree(s, b, p):
    """Org chart / 추진체계: root → children (row) → grandchildren (stacked list)."""
    x, y, w, h = b
    root = p["root"]
    kids = root.get("children", [])
    nh = p.get("node_h", 0.58)
    rw = min(2.8, w * 0.3)
    _node(s, x + (w - rw) / 2, y, rw, nh, root, "ink", "paper")
    for j, sd in enumerate(root.get("side", [])):  # advisory boxes beside the root
        sx = x + (w + rw) / 2 + 0.5 + j * (min(2.2, w * 0.2) + 0.15)
        sww = min(2.2, w * 0.2)
        line(s, x + (w + rw) / 2, y + nh / 2, sx, y + nh / 2, "muted", 1, dash=True)
        _node(s, sx, y + 0.06, sww, nh - 0.12, sd, "paper", "ink", "muted")
    if not kids:
        return
    n = len(kids)
    g = 0.15
    cw = (w - g * (n - 1)) / n
    bus = y + nh + 0.22
    ky = y + nh + 0.44
    cxs = [x + i * (cw + g) + cw / 2 for i in range(n)]
    line(s, x + w / 2, y + nh, x + w / 2, bus, "ink", 1)
    if n > 1:
        line(s, cxs[0], bus, cxs[-1], bus, "ink", 1)
    for i, k in enumerate(kids):
        kx = x + i * (cw + g)
        hl = p.get("highlight") == i
        line(s, cxs[i], bus, cxs[i], ky, "ink", 1)
        _node(s, kx, ky, cw, nh, k, "accent" if hl else tint("ink", 0.12), "paper")
        gk = [] if isinstance(k, str) else k.get("children", [])
        if gk:
            ly = ky + nh + 0.12
            gg = 0.07
            each = min(0.46, (y + h - ly - gg * (len(gk) - 1)) / len(gk))
            for j, c in enumerate(gk):
                cy = ly + j * (each + gg)
                line(s, kx + 0.08, cy + each / 2, kx + 0.2, cy + each / 2, "rule", 1)
                _node(s, kx + 0.2, cy, cw - 0.2, each, c, "paper", "ink",
                      "accent" if hl else "rule")
            line(s, kx + 0.08, ky + nh, kx + 0.08, ly + (len(gk) - 1) * (each + gg) + each / 2,
                 "rule", 1)
        elif not isinstance(k, str) and k.get("items"):
            text(s, (kx + 0.05, ky + nh + 0.12, cw - 0.1, y + h - ky - nh - 0.12), k["items"],
                 9.5, bullets=True, gap=2, name=k["name"])


def p_house(s, b, p):
    """전략 체계도: vision roof → goals → pillars → foundation."""
    x, y, w, h = b
    rh = 0.8
    roof = rect(s, x, y, w, rh, "ink", shape=MSO_SHAPE.TRAPEZOID)
    roof.adjustments[0] = w * 0.1 / rh
    text(s, (x + w * 0.15, y + 0.08, w * 0.7, 0.22), p.get("vision_label", "VISION"), 9,
         "#9DB6FF", bold=True, align="c", fit=False)
    text(s, (x + w * 0.15, y + 0.28, w * 0.7, rh - 0.32), p["vision"], 14, "paper", bold=True,
         align="c", anchor="m", min_size=10, emph="#9DB6FF", name="vision")
    cy = y + rh + 0.1
    goals = p.get("goals", [])
    if goals:
        gh = 0.5
        rect(s, x, cy, 1.1, gh, "accent")
        text(s, (x, cy, 1.1, gh), p.get("goals_label", "목표"), 10.5, "paper", bold=True,
             align="c", anchor="m", fit=False)
        gw = (w - 1.2 - 0.08 * (len(goals) - 1)) / len(goals)
        for i, gl in enumerate(goals):
            gx = x + 1.2 + i * (gw + 0.08)
            rect(s, gx, cy, gw, gh, tint("accent", 0.9))
            text(s, (gx + 0.1, cy, gw - 0.2, gh), gl, 10.5, "ink", bold=True, align="c",
                 anchor="m", min_size=8, name="goal")
        cy += gh + 0.1
    base = p.get("base", [])
    bh = 0.4
    ph = y + h - cy - (len(base) * (bh + 0.06) + 0.04 if base else 0)
    pillars = p["pillars"]
    n = len(pillars)
    pg = 0.12
    pw = (w - pg * (n - 1)) / n
    for i, pl in enumerate(pillars):
        px = x + i * (pw + pg)
        hl = p.get("highlight") == i
        rect(s, px, cy, pw, ph, "soft")
        rect(s, px, cy, pw, 0.44, "accent" if hl else tint("ink", 0.12))
        text(s, (px + 0.08, cy, pw - 0.16, 0.44), pl["title"], 11, "paper", bold=True,
             align="c", anchor="m", min_size=8, name="pillar")
        text(s, (px + 0.14, cy + 0.56, pw - 0.28, ph - 0.66), pl["items"], p.get("size", 9.5),
             bullets=True, gap=2, name=pl["title"])
    by = cy + ph + 0.06
    for i, bs in enumerate(base):
        rect(s, x, by, w, bh, tint("ink", 0.8) if i == 0 else "soft")
        text(s, (x + 0.2, by, w - 0.4, bh), bs, 10.5, "ink", bold=True, align="c", anchor="m",
             min_size=8, name="base")
        by += bh + 0.06


def p_milestones(s, b, p):
    x, y, w, h = b
    ev = p["events"]
    n = len(ev)
    if h > 3.6:  # a taller band only spreads the labels apart; centre a fixed band instead
        y, h = y + (h - 3.6) / 2, 3.6
    my = y + h / 2
    line(s, x, my, x + w, my, "ink", 1.5)
    seg = w / n
    tw = min(2 * seg - 0.2, 2.6)
    for i, e in enumerate(ev):
        cx = x + seg * i + seg / 2
        hl = e.get("highlight")
        r = 0.12 if hl else 0.085
        rect(s, cx - r, my - r, 2 * r, 2 * r, "accent" if hl else "paper", "accent",
             MSO_SHAPE.OVAL, lw=1.75)
        up = i % 2 == 0
        line(s, cx, my - r if up else my + r, cx, my - 0.38 if up else my + 0.38, "rule", 1)
        bh = h / 2 - 0.45
        tx = cx - tw / 2
        body = e.get("text")
        th = text_height(e["title"], 11, tw, bold=True, line=1.05)
        bth = min(text_height(body, 9.5, tw), bh - th - 0.35) if body else 0
        total = 0.3 + th + (bth + 0.05 if body else 0)
        ty = (my - 0.42 - total) if up else my + 0.42
        text(s, (tx, ty, tw, 0.28), e["date"], 12, "accent", bold=True, align="c", fit=False)
        text(s, (tx, ty + 0.3, tw, th + 0.02), e["title"], 11, "ink", bold=True, align="c",
             line=1.05, name="milestone")
        if body:
            text(s, (tx, ty + 0.3 + th + 0.05, tw, bth + 0.02), body, 9.5, "muted", align="c",
                 name="milestone text")


def p_profiles(s, b, p):
    x, y, w, h = b
    people = p["people"]
    ncol = p.get("cols", len(people))
    nrow = math.ceil(len(people) / ncol)
    g = 0.15
    cw, chh = (w - g * (ncol - 1)) / ncol, (h - g * (nrow - 1)) / nrow
    for i, pr in enumerate(people):
        cx, cy = x + (i % ncol) * (cw + g), y + (i // ncol) * (chh + g)
        hl = p.get("highlight") == i
        rect(s, cx, cy, cw, chh, "paper", "accent" if hl else "rule", lw=1 if hl else 0.75)
        d = min(0.85, chh * 0.35, cw * 0.28)
        av = rect(s, cx + 0.16, cy + 0.16, d, d, tint("accent", 0.85) if hl else "soft",
                  shape=MSO_SHAPE.OVAL)
        shape_text(av, plain(pr["name"])[0], d * 30, "accent" if hl else "muted")
        tx = cx + 0.16 + d + 0.14
        tw = cx + cw - tx - 0.12
        text(s, (tx, cy + 0.14, tw, 0.3), pr["name"], 12.5, "ink", bold=True, anchor="m",
             min_size=9, name="name")
        text(s, (tx, cy + 0.44, tw, 0.24), pr.get("role", ""), 10, "accent", bold=True,
             min_size=8, name="role")
        if pr.get("org"):
            text(s, (tx, cy + 0.68, tw, 0.24), pr["org"], 9, "muted", min_size=7.5, name="org")
        top = cy + 0.16 + max(d, 0.78) + 0.12
        hline(s, cx + 0.16, top, cw - 0.32)
        if pr.get("items"):
            text(s, (cx + 0.16, top + 0.1, cw - 0.32, cy + chh - top - 0.2), pr["items"],
                 p.get("size", 9.5), bullets=True, gap=2, name=pr["name"])


def p_numbered(s, b, p):
    x, y, w, h = b
    items = p["items"]
    ncol = p.get("cols", 1)
    per = math.ceil(len(items) / ncol)
    g = 0.3
    cw = (w - g * (ncol - 1)) / ncol
    rh = h / per
    for i, it in enumerate(items):
        cx, cy = x + (i // per) * (cw + g), y + (i % per) * rh
        hl = p.get("highlight") == i
        text(s, (cx, cy + 0.04, 0.75, 0.5), f"{i + 1:02d}", 24, "accent" if hl else
             tint("ink", 0.55), bold=True, fit=False)
        text(s, (cx + 0.8, cy + 0.08, cw - 0.8, 0.32), it["title"], 12.5,
             "accent" if hl else "ink", bold=True, anchor="m", min_size=9, name="item title")
        if it.get("text"):
            text(s, (cx + 0.8, cy + 0.44, cw - 0.8, rh - 0.54), it["text"], 10, "muted",
                 name=it["title"])
        if i % per != per - 1 and i != len(items) - 1:
            hline(s, cx, cy + rh - 0.04, cw)


def p_flow(s, b, p):
    """Left→right nodes with labelled arrows (BM, 서비스 흐름, 데이터 흐름)."""
    x, y, w, h = b
    nodes, edges = p["nodes"], p.get("edges", [])
    n = len(nodes)
    ag = min(1.5, w * 0.13)
    nw = (w - ag * (n - 1)) / n
    my = y + h / 2
    for i, nd in enumerate(nodes):
        nx = x + i * (nw + ag)
        hl = p.get("highlight") == i
        body = nd.get("items") or nd.get("text")
        bh = h if body else min(h, 1.1)
        by = my - bh / 2
        rect(s, nx, by, nw, bh, "paper", "accent" if hl else "ink",
             lw=1.25 if hl else 0.75)
        hh = 0.46 if body else bh
        rect(s, nx, by, nw, hh, "accent" if hl else "ink")
        text(s, (nx + 0.08, by, nw - 0.16, hh), nd["title"], 11.5, "paper", bold=True,
             align="c", anchor="m", min_size=8, name="flow node")
        if body:
            text(s, (nx + 0.14, by + hh + 0.12, nw - 0.28, bh - hh - 0.2), body, 9.5,
                 bullets="items" in nd, gap=2, name=nd["title"])
        if i < n - 1:
            e = edges[i] if i < len(edges) else ""
            fwd, back = (e, None) if isinstance(e, str) else (e.get("label", ""), e.get("back"))
            ax1, ax2 = nx + nw + 0.06, nx + nw + ag - 0.06
            oy = 0.16 if back else 0
            line(s, ax1, my - oy, ax2, my - oy, "accent", 1.5, arrow=True)
            if fwd:
                text(s, (ax1, my - oy - 0.5, ax2 - ax1, 0.45), fwd, 8.5, "accent", bold=True,
                     align="c", anchor="b", line=1.0, min_size=7, name="edge")
            if back:
                line(s, ax2, my + oy, ax1, my + oy, "muted", 1.25, arrow=True)
                text(s, (ax1, my + oy + 0.05, ax2 - ax1, 0.45), back, 8.5, "muted", bold=True,
                     align="c", line=1.0, min_size=7, name="edge")


def p_progress(s, b, p):
    x, y, w, h = b
    items = p["items"]
    mx = p.get("max", 100)
    rh = min(0.55, h / len(items))
    lw = p.get("label_width", w * 0.32)
    vw = 0.75
    for i, it in enumerate(items):
        ry = y + i * rh
        hl = p.get("highlight") == i
        text(s, (x, ry, lw - 0.12, rh), it["label"], 10.5, "ink", bold=hl, anchor="m",
             min_size=8, name="progress label")
        tw = w - lw - vw
        rect(s, x + lw, ry + rh * 0.3, tw, rh * 0.4, "soft")
        rect(s, x + lw, ry + rh * 0.3, tw * min(1, it["value"] / mx), rh * 0.4,
             "accent" if hl else tint("ink", 0.35))
        text(s, (x + w - vw + 0.1, ry, vw - 0.1, rh), it.get("display", f"{it['value']}%"),
             11, "accent" if hl else "ink", bold=True, align="r", anchor="m", fit=False)


def p_roadmap(s, b, p):
    """Tracks (rows) × phases (columns) with item chips."""
    x, y, w, h = b
    phases, tracks = p["phases"], p["tracks"]
    lw = p.get("label_width", 1.5)
    pw = (w - lw) / len(phases)
    hh = 0.46
    hl = p.get("highlight")
    if hl is not None:
        rect(s, x + lw + hl * pw, y + hh, pw, h - hh, tint("accent", 0.94))
    for i, ph in enumerate(phases):
        sh = rect(s, x + lw + i * pw, y, pw - 0.02, hh, "accent" if hl == i else "ink",
                  shape=MSO_SHAPE.PENTAGON if i == 0 else MSO_SHAPE.CHEVRON)
        sh.adjustments[0] = 0.25
        shape_text(sh, ph, 10.5, "paper", box=(pw - 0.5, hh))
    th = (h - hh - 0.1) / len(tracks)
    for r, t in enumerate(tracks):
        ty = y + hh + 0.1 + r * th
        rect(s, x, ty, lw - 0.1, th - 0.08, "soft")
        text(s, (x + 0.08, ty, lw - 0.26, th - 0.08), t["name"], 10.5, "ink", bold=True,
             align="c", anchor="m", min_size=8, name="track")
        if r < len(tracks) - 1:
            hline(s, x, ty + th - 0.04, w)
        for i, cell in enumerate(t["cells"]):
            items = [cell] if isinstance(cell, str) else cell
            items = [c for c in items if c]
            if not items:
                continue
            cg = 0.05
            ch = min(0.4, (th - 0.2 - cg * (len(items) - 1)) / len(items))
            for j, it in enumerate(items):
                cy = ty + 0.04 + j * (ch + cg)
                on = it.startswith("**")
                chip = rect(s, x + lw + i * pw + 0.1, cy, pw - 0.2, ch,
                            "accent" if on else "paper", None if on else "rule",
                            MSO_SHAPE.ROUNDED_RECTANGLE)
                chip.adjustments[0] = 0.2
                shape_text(chip, plain(it), 9, "paper" if on else "ink", bold=on,
                           box=(pw - 0.3, ch))


# ------------------------------------------------------------------ icons & images

ICONS_PATH = Path(__file__).resolve().parent.parent / "assets/icons.json"
_ICON = {}


def icon(s, name, x, y, size, color="accent"):
    """Lucide icon (assets/icons.json) rendered to a transparent PNG."""
    import pymupdf
    if "set" not in _ICON:
        _ICON["set"] = json.loads(ICONS_PATH.read_text(encoding="utf-8"))
    icons = _ICON["set"]
    if name not in icons:
        warn(f"unknown icon {name!r} — search the names in assets/icons.json")
        name = "circle-help"
    key = (name, col(color))
    if key not in _ICON:
        svg = ('<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" '
               f'fill="none" stroke="{col(color)}" stroke-width="2" stroke-linecap="round" '
               f'stroke-linejoin="round">{icons[name]}</svg>')
        pix = pymupdf.open(stream=svg.encode(), filetype="svg")[0].get_pixmap(
            matrix=pymupdf.Matrix(8, 8), alpha=True)
        _ICON[key] = pix.tobytes("png")
    return s.shapes.add_picture(io.BytesIO(_ICON[key]), Inches(x), Inches(y), Inches(size),
                                Inches(size))


def alpha(sh, a):
    """Make a solid-filled shape translucent (a = opacity 0..1)."""
    clr = sh._element.spPr.find(qn("a:solidFill")).find(qn("a:srgbClr"))
    etree.SubElement(clr, qn("a:alpha"), val=str(int(a * 100000)))
    return sh


def _src(p):
    src = p.get("src") and (CTX["base"] / p["src"])
    return src if src and src.exists() else None


def picture(s, src, x, y, w, h, crop="center"):
    """Fill (x, y, w, h) with the image, cropping the overflow. crop: center | top."""
    pic = s.shapes.add_picture(str(src), Inches(x), Inches(y), Inches(w), Inches(h))
    iw, ih = pic.image.size
    box, img = w / h, iw / ih
    if img > box:  # too wide → trim sides
        cut = (1 - box / img) / 2
        pic.crop_left = pic.crop_right = cut
    elif img < box:  # too tall → trim bottom (screenshots) or both ends (photos)
        cut = 1 - img / box
        if crop == "top":
            pic.crop_bottom = cut
        else:
            pic.crop_top = pic.crop_bottom = cut / 2
    return pic


def placeholder(s, x, y, w, h, alt, dark=False):
    ph = rect(s, x, y, w, h, "#2A3550" if dark else "soft", None if dark else "rule")
    if not dark:
        ph.line.dash_style = 4
    d = min(0.5, w * 0.2, h * 0.3)
    icon(s, "image", x + w / 2 - d / 2, y + h / 2 - d * 0.9, d, "muted")
    text(s, (x + 0.1, y + h / 2 + d * 0.25, w - 0.2, 0.5), alt or "이미지를 넣어 주세요", 9,
         "muted", align="c", fit=False)


FRAME_RATIO = {"browser": None, "laptop": 16 / 10, "tablet": 4 / 3, "phone": 9 / 19.5}


def _frame(s, b, kind, src, alt):
    """Draw a device frame around the screenshot; returns the screen rect."""
    x, y, w, h = b
    if kind == "browser":
        rect(s, x, y, w, h, "paper", "rule", lw=1)
        rect(s, x, y, w, 0.3, "soft")
        for i, c in enumerate(("#E8A39B", "#E9CF8C", "#A9D3A1")):
            rect(s, x + 0.14 + i * 0.16, y + 0.1, 0.1, 0.1, c, shape=MSO_SHAPE.OVAL)
        url = rect(s, x + 0.7, y + 0.06, min(3.2, w - 0.9), 0.18, "paper",
                   shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        url.adjustments[0] = 0.5
        scr = (x + 0.01, y + 0.3, w - 0.02, h - 0.31)
    elif kind == "laptop":
        base_h = h * 0.07
        sw = min(w * 0.88, (h - base_h) * FRAME_RATIO["laptop"])
        sh_ = sw / FRAME_RATIO["laptop"]
        sx = x + (w - sw) / 2
        sy = y + (h - base_h - sh_) / 2
        body = rect(s, sx, sy, sw, sh_, "#1B2233", shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        body.adjustments[0] = 0.03
        base = rect(s, sx - sw * 0.07, sy + sh_, sw * 1.14, base_h, "#C9CED8",
                    shape=MSO_SHAPE.TRAPEZOID)
        base.rotation = 180
        bz = sw * 0.025
        scr = (sx + bz, sy + bz, sw - 2 * bz, sh_ - 2 * bz * 1.3)
    else:  # phone / tablet
        r = FRAME_RATIO[kind]
        fh = min(h, w / r)
        fw = fh * r
        fx, fy = x + (w - fw) / 2, y + (h - fh) / 2
        body = rect(s, fx, fy, fw, fh, "#1B2233", shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        body.adjustments[0] = 0.12 if kind == "phone" else 0.05
        bz = fw * (0.045 if kind == "phone" else 0.05)
        scr = (fx + bz, fy + bz * (1.6 if kind == "phone" else 1), fw - 2 * bz,
               fh - bz * (3.2 if kind == "phone" else 2))
    if src:
        picture(s, src, *scr, crop="top")
    else:
        placeholder(s, *scr, alt, dark=kind != "browser")
    if kind == "phone":
        notch = rect(s, scr[0] + scr[2] * 0.35, scr[1] + 0.04, scr[2] * 0.3, 0.07, "#1B2233",
                     shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        notch.adjustments[0] = 0.5
    return scr


def p_image(s, b, p):
    """Image (screenshot, photo, diagram) with optional device frame and numbered callouts."""
    x, y, w, h = b
    cap = p.get("caption")
    callouts = p.get("callouts", [])
    legend = callouts and any(c.get("text") for c in callouts)
    lw = w * p.get("legend_width", 0.34) if legend else 0
    iw, ih = w - lw - (0.25 if legend else 0), h - (0.32 if cap else 0)
    src = _src(p)
    frame = p.get("frame")
    if frame:
        scr = _frame(s, (x, y, iw, ih), frame, src, p.get("alt"))
    elif src:
        if p.get("fit") == "contain":
            pic = s.shapes.add_picture(str(src), Inches(x), Inches(y))
            k = min(iw / (pic.width / 914400), ih / (pic.height / 914400))
            pic.width, pic.height = int(pic.width * k), int(pic.height * k)
            pic.left = int(Inches(x) + (Inches(iw) - pic.width) / 2)
            pic.top = int(Inches(y) + (Inches(ih) - pic.height) / 2)
            scr = (pic.left / 914400, pic.top / 914400, pic.width / 914400, pic.height / 914400)
        else:
            picture(s, src, x, y, iw, ih, p.get("crop", "center"))
            scr = (x, y, iw, ih)
        if p.get("border", True):
            rect(s, *scr, None, "rule")
    else:
        placeholder(s, x, y, iw, ih, p.get("alt"))
        scr = (x, y, iw, ih)
    for i, c in enumerate(callouts):
        d = 0.3
        m = rect(s, scr[0] + c["x"] * scr[2] - d / 2, scr[1] + c["y"] * scr[3] - d / 2, d, d,
                 "accent", "paper", MSO_SHAPE.OVAL, lw=1.5)
        shape_text(m, i + 1, 10, "paper")
        if c.get("box"):  # highlight a region: [w, h] relative to the image
            bw, bh = c["box"]
            rr = rect(s, scr[0] + c["x"] * scr[2] - bw * scr[2] / 2,
                      scr[1] + c["y"] * scr[3] - bh * scr[3] / 2, bw * scr[2], bh * scr[3],
                      None, "accent", shape=MSO_SHAPE.ROUNDED_RECTANGLE, lw=2)
            rr.adjustments[0] = 0.05
            s.shapes._spTree.remove(m._element)
            s.shapes._spTree.append(m._element)  # marker above its box
    if legend:
        lx = x + iw + 0.25
        rows = [c for c in callouts if c.get("text")]
        rh = min(1.1, ih / len(rows))
        for i, c in enumerate(callouts):
            if not c.get("text"):
                continue
            ly = y + i * rh
            m = rect(s, lx, ly + 0.02, 0.26, 0.26, "accent", shape=MSO_SHAPE.OVAL)
            shape_text(m, i + 1, 9, "paper")
            title, body = (c["text"], None) if "\n" not in c["text"] else c["text"].split("\n", 1)
            text(s, (lx + 0.38, ly, lw - 0.38, 0.3), title, 11, "ink", bold=True, anchor="m",
                 min_size=8.5, name="callout title")
            if body:
                text(s, (lx + 0.38, ly + 0.32, lw - 0.38, rh - 0.4), body, 9.5, "muted",
                     name="callout text")
    if cap:
        text(s, (x, y + ih + 0.08, iw, 0.24), f"< {cap} >", 9, "muted", align="c", fit=False)


def p_images(s, b, p):
    """Grid of photos / screenshots with captions."""
    x, y, w, h = b
    items = p["items"]
    ncol = p.get("cols", min(len(items), 4))
    nrow = math.ceil(len(items) / ncol)
    g = p.get("gap", 0.12)
    cw, chh = (w - g * (ncol - 1)) / ncol, (h - g * (nrow - 1)) / nrow
    for i, it in enumerate(items):
        cx, cy = x + (i % ncol) * (cw + g), y + (i // ncol) * (chh + g)
        cap = it.get("caption")
        ih = chh - (0.3 if cap else 0)
        src = _src(it)
        if src:
            picture(s, src, cx, cy, cw, ih, it.get("crop", p.get("crop", "center")))
            rect(s, cx, cy, cw, ih, None, "rule")
        else:
            placeholder(s, cx, cy, cw, ih, it.get("alt"))
        if it.get("tag"):
            tg = rect(s, cx + 0.08, cy + 0.08, text_w(it["tag"], 8.5, True) + 0.2, 0.24,
                      "accent" if p.get("highlight") == i else "ink")
            shape_text(tg, it["tag"], 8.5, "paper")
        if cap:
            text(s, (cx, cy + ih + 0.05, cw, 0.25), cap, 9, "text", align="c", anchor="m",
                 min_size=7.5, name="image caption")


# ------------------------------------------------------------------ more panels

def p_group(s, b, p):
    """Nested grid: a panel whose body is another grid of panels."""
    grid(s, p["body"], b, p.get("gap", 0.15))


def p_icons(s, b, p):
    x, y, w, h = b
    items = p["items"]
    n = len(items)
    ncol = p.get("cols", n if n <= 4 else math.ceil(n / 2))
    nrow = math.ceil(n / ncol)
    g = 0.18
    cw, chh = (w - g * (ncol - 1)) / ncol, (h - g * (nrow - 1)) / nrow
    left = p.get("align") == "left"
    card = p.get("style") == "card"
    for i, it in enumerate(items):
        cx, cy = x + (i % ncol) * (cw + g), y + (i // ncol) * (chh + g)
        hl = p.get("highlight") == i
        if card:
            rect(s, cx, cy, cw, chh, "paper", "accent" if hl else "rule")
            cx, cy, cww, chh_ = cx + 0.16, cy + 0.16, cw - 0.32, chh - 0.32
        else:
            cww, chh_ = cw, chh
        d = min(0.7, chh_ * 0.36, cww * 0.4)
        ox = cx if left else cx + (cww - d) / 2
        rect(s, ox, cy, d, d, "accent" if hl else tint("accent", 0.9), shape=MSO_SHAPE.OVAL)
        icon(s, it["icon"], ox + d * 0.23, cy + d * 0.23, d * 0.54, "paper" if hl else "accent")
        if left:
            tx, tw = cx + d + 0.14, cww - d - 0.14
            text(s, (tx, cy, tw, d), it["title"], 12, "accent" if hl else "ink", bold=True,
                 anchor="m", min_size=9, name="icon title")
            if it.get("text"):
                text(s, (cx, cy + d + 0.1, cww, chh_ - d - 0.1), it["text"], 10, "muted",
                     name=it["title"])
        else:
            text(s, (cx, cy + d + 0.1, cww, 0.32), it["title"], 12, "accent" if hl else "ink",
                 bold=True, align="c", anchor="m", min_size=9, name="icon title")
            if it.get("text"):
                text(s, (cx + 0.05, cy + d + 0.46, cww - 0.1, chh_ - d - 0.46), it["text"], 10,
                     "muted", align="c", name=it["title"])


def p_venn(s, b, p):
    x, y, w, h = b
    sets = p["sets"]
    n = len(sets)
    cols = ["accent", tint("ink", 0.2), "#8A94A6"]
    cx, cy = x + w / 2, y + h / 2
    if n == 2:
        r = min(h / 2, w / 3.3)
        centers = [(cx - 0.62 * r, cy), (cx + 0.62 * r, cy)]
        mid = (cx, cy)
    else:
        r = min(h / 3.05, w / 3.25)
        cy += 0.04 * r
        centers = [(cx - 0.58 * r, cy - 0.45 * r), (cx + 0.58 * r, cy - 0.45 * r),
                   (cx, cy + 0.55 * r)]
        mid = (cx, cy - 0.12 * r)
    for (sx, sy), c in zip(centers, cols):
        alpha(rect(s, sx - r, sy - r, 2 * r, 2 * r, c, shape=MSO_SHAPE.OVAL), 0.16)
        rect(s, sx - r, sy - r, 2 * r, 2 * r, None, c, MSO_SHAPE.OVAL, lw=1.5)
    for (sx, sy), c, st in zip(centers, cols, sets):
        dx, dy = sx - mid[0], sy - mid[1]
        k = math.hypot(dx, dy) or 1
        lx, ly = sx + dx / k * r * 0.42, sy + dy / k * r * 0.42
        bw = r * (0.85 if n == 2 else 0.95)
        items = st.get("items", [])
        th = 0.32 + (text_height(items, 9.5, bw, line=1.05, gap=1) if items else 0)
        text(s, (lx - bw / 2, ly - th / 2, bw, 0.3), st["title"], 12.5,
             "accent" if c == "accent" else "ink", bold=True, align="c", min_size=9,
             name="venn title")
        if items:
            text(s, (lx - bw / 2, ly - th / 2 + 0.32, bw, th - 0.3), items, 9.5, "text",
                 align="c", line=1.05, gap=1, min_size=7.5, name=st["title"])
    if p.get("center"):
        cw = r * (0.6 if n == 2 else 0.62)
        text(s, (mid[0] - cw / 2, mid[1] - 0.35, cw, 0.7), p["center"], 11, "ink", bold=True,
             align="c", anchor="m", line=1.05, min_size=8, name="venn center")


BMC = [("partners", "핵심 파트너", "handshake"), ("activities", "핵심 활동", "zap"),
       ("resources", "핵심 자원", "boxes"), ("value", "가치 제안", "gem"),
       ("relationships", "고객 관계", "heart-handshake"), ("channels", "채널", "megaphone"),
       ("segments", "고객 세그먼트", "users"), ("costs", "비용 구조", "receipt"),
       ("revenue", "수익원", "banknote")]


def p_canvas(s, b, p):
    """Business Model Canvas (9 blocks)."""
    x, y, w, h = b
    g = 0.06
    cw = (w - 4 * g) / 5
    th = (h - g) * 0.66
    bh = h - g - th
    half = (th - g) / 2
    col_x = [x + i * (cw + g) for i in range(5)]
    pos = {"partners": (col_x[0], y, cw, th), "activities": (col_x[1], y, cw, half),
           "resources": (col_x[1], y + half + g, cw, half), "value": (col_x[2], y, cw, th),
           "relationships": (col_x[3], y, cw, half),
           "channels": (col_x[3], y + half + g, cw, half), "segments": (col_x[4], y, cw, th),
           "costs": (x, y + th + g, (w - g) / 2, bh),
           "revenue": (x + (w + g) / 2, y + th + g, (w - g) / 2, bh)}
    hl = p.get("highlight", [])
    hl = [hl] if isinstance(hl, str) else hl
    for key, label, ic in BMC:
        bx, by, bw, bhh = pos[key]
        on = key in hl
        rect(s, bx, by, bw, bhh, tint("accent", 0.9) if on else "soft")
        if on:
            rect(s, bx, by, bw, 0.05, "accent")
        icon(s, ic, bx + 0.12, by + 0.13, 0.22, "accent" if on else "ink")
        text(s, (bx + 0.4, by + 0.1, bw - 0.5, 0.28), label, 10.5, "accent" if on else "ink",
             bold=True, anchor="m", min_size=8, name=label)
        if p.get(key):
            text(s, (bx + 0.12, by + 0.5, bw - 0.24, bhh - 0.6), p[key], p.get("size", 9.5),
                 bullets=True, gap=2, name=label)


STATUS = {"done": ("circle-check", "accent", "완료"), "partial": ("circle-dot-dashed", "accent", "진행"),
          "todo": ("circle", "muted", "미착수"), "fail": ("circle-x", "#D64545", "미흡")}


def p_checklist(s, b, p):
    """Checklist (status icons) or scorecard (rating dots, optional weights)."""
    x, y, w, h = b
    items = p["items"]
    rating = any("score" in it for it in items)
    weighted = any("weight" in it for it in items)
    cols = p.get("columns") or (["평가 항목"] + (["배점"] if weighted else []) +
                                (["점수", "근거"] if rating else ["상태", "비고"]))
    rel = p.get("widths") or ([3.2] + ([0.8] if weighted else []) + ([1.8, 3] if rating else [1.1, 3]))
    ws = [w * r / sum(rel) for r in rel]
    xs = [x + sum(ws[:i]) for i in range(len(ws))]
    hh = 0.36
    rect(s, x, y, w, hh, "ink")
    for cx, cw, c in zip(xs, ws, cols):
        text(s, (cx + 0.1, y, cw - 0.2, hh), c, 10, "paper", bold=True, align="c", anchor="m",
             fit=False)
    summary = p.get("summary")
    if weighted and rating and summary is None:
        tot = sum(it.get("weight", 0) * it.get("score", 0) / it.get("max", 5) for it in items)
        summary = f"종합 점수 **{tot:.1f}** / {sum(it.get('weight', 0) for it in items)}"
    sh = 0.42 if summary else 0
    rh = min(0.55, (h - hh - sh) / len(items))
    for r, it in enumerate(items):
        ry = y + hh + r * rh
        hl = p.get("highlight") == r
        if hl:
            rect(s, x, ry, w, rh, tint("accent", 0.92))
        hline(s, x, ry + rh, w)
        text(s, (xs[0] + 0.12, ry, ws[0] - 0.2, rh), it["label"], 10, "ink", bold=hl,
             anchor="m", min_size=8, name="check label")
        c = 1
        if weighted:
            text(s, (xs[1], ry, ws[1], rh), str(it.get("weight", "")), 10, "text", align="c",
                 anchor="m", fit=False)
            c = 2
        mx, mw = xs[c], ws[c]
        if rating:
            mxv = it.get("max", 5)
            d = min(0.16, (mw - 0.7) / mxv - 0.05)
            sx = mx + (mw - (mxv * (d + 0.05) + 0.45)) / 2
            for k in range(mxv):
                full = k + 1 <= it["score"]
                rect(s, sx + k * (d + 0.05), ry + rh / 2 - d / 2, d, d,
                     "accent" if full else "paper", "accent" if full else "rule",
                     MSO_SHAPE.OVAL)
            text(s, (sx + mxv * (d + 0.05) + 0.05, ry, 0.45, rh), f"{it['score']}", 10.5,
                 "accent", bold=True, anchor="m", fit=False)
        else:
            st = it.get("status", "todo")
            st = "done" if st is True else "todo" if st is False else st
            ic, cl, lab = STATUS[st]
            icon(s, ic, mx + mw / 2 - 0.42, ry + rh / 2 - 0.11, 0.22, cl)
            text(s, (mx + mw / 2 - 0.15, ry, 0.7, rh), it.get("status_label", lab), 9.5, cl,
                 bold=True, anchor="m", fit=False)
        if it.get("note"):
            text(s, (xs[c + 1] + 0.12, ry, ws[c + 1] - 0.2, rh), it["note"], 9.5, "muted",
                 anchor="m", min_size=7.5, name="check note")
    if summary:
        sy = y + hh + len(items) * rh + 0.06
        rect(s, x, sy, w, sh - 0.06, "soft")
        text(s, (x + 0.15, sy, w - 0.3, sh - 0.06), summary, 11, "ink", bold=True, align="r",
             anchor="m", fit=False)


KOREA_TILES = {"서울": (1, 0), "강원": (2, 0), "인천": (0, 1), "경기": (1, 1), "충북": (2, 1),
               "경북": (3, 1), "충남": (0, 2), "세종": (1, 2), "대전": (2, 2), "대구": (3, 2),
               "전북": (0, 3), "광주": (1, 3), "경남": (2, 3), "울산": (3, 3), "전남": (1, 4),
               "부산": (2, 4), "제주": (0, 5)}


def p_map(s, b, p):
    """South Korea tile map (17 시·도) coloured by value."""
    x, y, w, h = b
    vals = p.get("values", {})
    fmt = p.get("format", "{:,}")
    hl = set(p.get("highlight", []))
    lg = 0.5
    t = min(w / 4, (h - lg) / 6)
    g = t * 0.06
    ox, oy = x + (w - 4 * t) / 2, y + (h - lg - 6 * t) / 2
    nums = [v for v in vals.values() if isinstance(v, (int, float))]
    lo, hi = (min(nums), max(nums)) if nums else (0, 1)
    for name, (c, r) in KOREA_TILES.items():
        v = vals.get(name)
        k = (v - lo) / ((hi - lo) or 1) if isinstance(v, (int, float)) else None
        fill = "soft" if k is None else tint("accent", 0.9 - 0.8 * k)
        tl = rect(s, ox + c * t + g / 2, oy + r * t + g / 2, t - g, t - g, fill,
                  "ink" if name in hl else None, MSO_SHAPE.ROUNDED_RECTANGLE,
                  lw=2.25 if name in hl else 0.75)
        tl.adjustments[0] = 0.12
        fg = "paper" if k is not None and k > 0.5 else "ink"
        text(s, (ox + c * t, oy + r * t + t * 0.18, t, t * 0.32), name, min(11, t * 16), fg,
             bold=True, align="c", anchor="m", fit=False)
        if v is not None:
            text(s, (ox + c * t, oy + r * t + t * 0.5, t, t * 0.3),
                 fmt.format(v) if isinstance(v, (int, float)) else str(v), min(9.5, t * 13), fg,
                 align="c", anchor="m", fit=False)
    if nums:
        ly = y + h - lg + 0.12
        steps = 5
        sw = min(0.45, w / 12)
        lx = x + (w - steps * sw) / 2
        text(s, (lx - 1.3, ly, 1.2, 0.2), fmt.format(lo), 8.5, "muted", align="r", fit=False)
        for i in range(steps):
            rect(s, lx + i * sw, ly + 0.03, sw, 0.14, tint("accent", 0.9 - 0.8 * i / (steps - 1)))
        text(s, (lx + steps * sw + 0.1, ly, 1.6, 0.2),
             f"{fmt.format(hi)} {p.get('unit', '')}".strip(), 8.5, "muted", fit=False)


def _edge_point(n, side):
    x, y, w, h = n
    return {"l": (x, y + h / 2), "r": (x + w, y + h / 2), "t": (x + w / 2, y),
            "b": (x + w / 2, y + h)}[side]


def p_diagram(s, b, p):
    """Free-form diagram on a cols×rows grid: zones, nodes and routed arrows. lanes → swimlane."""
    x, y, w, h = b
    lanes = p.get("lanes")
    lw = p.get("lane_width", 1.0) if lanes else 0
    rows = len(lanes) if lanes else p.get("rows", 3)
    cols = p.get("cols", 4)
    gx, gw = x + lw, w - lw
    cw, rh = gw / cols, h / rows
    if lanes:
        for i, ln in enumerate(lanes):
            ly = y + i * rh
            rect(s, gx, ly, gw, rh, "soft" if i % 2 == 0 else "paper")
            lab = rect(s, x, ly, lw - 0.04, rh - 0.03, "ink")
            shape_text(lab, ln, 10, "paper", box=(lw - 0.15, rh - 0.1))
            hline(s, x, ly + rh - 0.015, w, "rule", 0.5)
    for z in p.get("zones", []):
        zx, zy = gx + z["col"] * cw + 0.04, y + z["row"] * rh + 0.04
        zw, zh = z.get("w", 1) * cw - 0.08, z.get("h", 1) * rh - 0.08
        zr = rect(s, zx, zy, zw, zh, tint("accent", 0.94) if z.get("highlight") else "soft",
                  "accent" if z.get("highlight") else "rule", lw=1)
        zr.line.dash_style = 4
        text(s, (zx + 0.1, zy + 0.05, zw - 0.2, 0.22), z["label"], 9, "accent"
             if z.get("highlight") else "muted", bold=True, fit=False)
    boxes = {}
    pad_x, pad_top, pad_bot = 0.14, 0.3 if p.get("zones") else 0.14, 0.14
    for nd in p["nodes"]:
        nx = gx + nd["col"] * cw + pad_x
        ny = y + nd["row"] * rh + pad_top
        nw = nd.get("w", 1) * cw - 2 * pad_x
        nh = nd.get("h", 1) * rh - pad_top - pad_bot
        if nd.get("height"):  # fixed node height, centred in its cells
            ny += (nh - nd["height"]) / 2
            nh = nd["height"]
        style = nd.get("style", "paper")
        fill, fg, border = {"accent": ("accent", "paper", None), "ink": ("ink", "paper", None),
                            "soft": ("soft", "ink", "rule")}.get(style, ("paper", "ink", "ink"))
        bx = rect(s, nx, ny, nw, nh, fill, border, MSO_SHAPE.ROUNDED_RECTANGLE, lw=1)
        bx.adjustments[0] = 0.08
        tx = nx + 0.1
        if nd.get("icon"):
            d = min(0.34, nh * 0.5)
            icon(s, nd["icon"], nx + 0.12, ny + nh / 2 - d / 2, d,
                 "paper" if style in ("accent", "ink") else "accent")
            tx = nx + 0.18 + d
        if nd.get("sub"):
            text(s, (tx, ny + 0.04, nx + nw - tx - 0.08, nh * 0.55 - 0.04), nd["label"], 10.5,
                 fg, bold=True, align="l" if nd.get("icon") else "c", anchor="b", line=1.0,
                 min_size=7.5, name="node")
            text(s, (tx, ny + nh * 0.55, nx + nw - tx - 0.08, nh * 0.45 - 0.04), nd["sub"], 8.5,
                 fg if style in ("accent", "ink") else "muted", align="l" if nd.get("icon")
                 else "c", line=1.0, min_size=7, name="node sub")
        else:
            text(s, (tx, ny, nx + nw - tx - 0.08, nh), nd["label"], 10.5, fg, bold=True,
                 align="l" if nd.get("icon") else "c", anchor="m", line=1.0, min_size=7.5,
                 name="node")
        boxes[nd["id"]] = (nx, ny, nw, nh)
    for e in p.get("edges", []):
        a, z = boxes[e["from"]], boxes[e["to"]]
        c = "accent" if e.get("highlight") else "muted"
        dash, both = e.get("style") == "dashed", e.get("both", False)
        ov_y = min(a[1] + a[3], z[1] + z[3]) - max(a[1], z[1])
        ov_x = min(a[0] + a[2], z[0] + z[2]) - max(a[0], z[0])
        if ov_y > 0.15:  # side by side → straight horizontal
            yy = max(a[1], z[1]) + ov_y / 2
            x1, x2 = (a[0] + a[2], z[0]) if a[0] < z[0] else (a[0], z[0] + z[2])
            pts = [(x1, yy), (x2, yy)]
        elif ov_x > 0.15:  # stacked → straight vertical
            xx = max(a[0], z[0]) + ov_x / 2
            y1, y2 = (a[1] + a[3], z[1]) if a[1] < z[1] else (a[1], z[1] + z[3])
            pts = [(xx, y1), (xx, y2)]
        else:  # elbow: leave sideways, enter from top/bottom
            p1 = _edge_point(a, "r" if z[0] > a[0] else "l")
            p3 = _edge_point(z, "t" if z[1] > a[1] else "b")
            pts = [p1, (p3[0], p1[1]), p3]
        for i in range(len(pts) - 1):
            line(s, *pts[i], *pts[i + 1], c, 1.5 if e.get("highlight") else 1.25,
                 arrow=i == len(pts) - 2, dash=dash, arrow_start=both and i == 0)
        if e.get("label"):
            (ax, ay), (bx2, by2) = pts[0], pts[1]
            mx, my = (ax + bx2) / 2, (ay + by2) / 2
            lw_ = text_w(e["label"], 8.5, True) + 0.16
            lb = rect(s, mx - lw_ / 2, my - 0.12, lw_, 0.24, "paper")
            shape_text(lb, e["label"], 8.5, c)


PANELS = {
    "bullets": p_bullets, "callout": p_callout, "table": p_table, "chart": p_chart,
    "kpi": p_kpi, "cards": p_cards, "process": p_process, "timeline": p_timeline,
    "cycle": p_cycle, "stack": p_stack, "image": p_image, "label": p_label, "arrow": p_arrow,
    "matrix": p_matrix, "pyramid": p_pyramid, "funnel": p_funnel, "tree": p_tree,
    "house": p_house, "milestones": p_milestones, "profiles": p_profiles,
    "numbered": p_numbered, "flow": p_flow, "progress": p_progress, "roadmap": p_roadmap,
    "group": p_group, "icons": p_icons, "venn": p_venn, "canvas": p_canvas,
    "checklist": p_checklist, "map": p_map, "diagram": p_diagram, "images": p_images,
}
FIXED_W = {"arrow": 0.35, "label": 0.5}


def panel(s, b, p):
    x, y, w, h = b
    if p.get("boxed"):
        rect(s, x, y, w, h, "soft")
        x, y, w, h = x + 0.15, y + 0.12, w - 0.3, h - 0.24
    if p.get("title"):
        bar = p.get("title_style", T["title_style"]) == "bar"
        if bar:
            rect(s, x, y, w, 0.34, "ink")
            text(s, (x + 0.12, y, w - 0.24, 0.34), p["title"], 11, "paper", bold=True,
                 anchor="m", align="c", name="panel title")
        else:
            rect(s, x, y + 0.08, 0.06, 0.2, "accent")
            text(s, (x + 0.14, y, w - 0.14, 0.34), p["title"], 11.5, "ink", bold=True,
                 anchor="m", name="panel title")
            hline(s, x, y + 0.38, w)
        y, h = y + 0.5, h - 0.5
    if p["type"] not in PANELS:
        sys.exit(f"slide {CTX['slide']}: unknown panel type {p['type']!r}. "
                 f"Choose from: {', '.join(PANELS)}")
    PANELS[p["type"]](s, (x, y, w, h), p)


def grid(s, body, box, gap=GAP):
    """body = [row, ...]; row = [panel, ...] or {"h": weight, "cols": [...]}; panel "w" = weight."""
    x, y, w, h = box
    rows = [r if isinstance(r, dict) else {"cols": r} for r in body]
    hw = [r.get("h", 1) for r in rows]
    avail = h - gap * (len(rows) - 1)
    cy = y
    for r, weight in zip(rows, hw):
        rh = avail * weight / sum(hw)
        cols = r["cols"]
        fixed = [FIXED_W[c["type"]] if "w" not in c and c["type"] in FIXED_W else None
                 for c in cols]
        flex = w - gap * (len(cols) - 1) - sum(f for f in fixed if f)
        tot = sum(c.get("w", 1) for c, f in zip(cols, fixed) if f is None) or 1
        cx = x
        for c, f in zip(cols, fixed):
            cw = f if f else flex * c.get("w", 1) / tot
            panel(s, (cx, cy, cw, rh), c)
            cx += cw + gap
        cy += rh + gap


# ------------------------------------------------------------------ slide chrome

def header(s, d, deck):
    sections = deck.get("sections", [])
    sec = d.get("section")
    if isinstance(sec, int) and sec < len(sections):
        sec = sections[sec]
    if sections:  # nav tabs, top-right
        tx = W - M
        for name in reversed(sections):
            cur = name == sec
            tw = text_w(name, 9, cur) + 0.05
            tx -= tw
            text(s, (tx, 0.28, tw, 0.22), name, 9, "accent" if cur else "muted", bold=cur,
                 fit=False)
            if cur:
                rect(s, tx, 0.52, tw - 0.05, 0.03, "accent")
            tx -= 0.28
    kicker = d.get("kicker") or sec
    if kicker:
        text(s, (M, 0.28, 6, 0.22), kicker, 10, "accent", bold=True, fit=False)
    y = 0.6
    hs = fit_size(d["title"], 20, W - 2 * M, 0.75, bold=True, min_size=16, line=1.05,
                  name="headline")
    hh = text_height(d["title"], hs, W - 2 * M, bold=True, line=1.05)
    text(s, (M, y, W - 2 * M, hh + 0.05), d["title"], hs, "ink", bold=True, line=1.05, fit=False)
    y += hh + 0.1
    if d.get("lead"):
        lh = text_height(d["lead"], 11, W - 2 * M)
        text(s, (M, y, W - 2 * M, lh + 0.04), d["lead"], 11, "muted", name="lead")
        y += lh + 0.08
    hline(s, M, y + 0.06, W - 2 * M, "ink", 1)
    return y + 0.28


def footer(s, d, deck, n):
    bottom = BODY_BOTTOM
    if d.get("source"):
        text(s, (M, 6.93, W - 2 * M - 1, 0.2), "출처: " + d["source"], 8, "muted", fit=False)
    hline(s, M, 7.17, W - 2 * M, "rule", 0.5)
    if deck.get("footer"):
        text(s, (M, 7.2, 8, 0.2), deck["footer"], 8, "muted", fit=False)
    text(s, (W - M - 1, 7.2, 1, 0.2), str(n), 8, "muted", bold=True, align="r", fit=False)
    return bottom


# ------------------------------------------------------------------ full-bleed slides

def f_cover(s, d, deck):
    rect(s, 0, 0, W, H, "paper")
    src = _src({"src": d.get("image")})
    if src:  # photo panel on the right
        picture(s, src, W - 4.6, 0, 4.6, H)
        rect(s, W - 4.6, 0, 0.08, H, "accent")
    else:
        rect(s, W - 2.6, 0, 2.6, H, tint("accent", 0.92))
        rect(s, W - 2.6, 0, 0.08, H, "accent")
        rect(s, W - 1.6, H - 1.6, 0.6, 0.6, "accent")
    if d.get("kicker"):
        text(s, (0.9, 1.7, 9, 0.35), d["kicker"], 13, "accent", bold=True)
    text(s, (0.9, 2.1, 9.4, 1.9), d["title"], 34, "ink", bold=True, anchor="b", line=1.05,
         min_size=24, name="cover title")
    hline(s, 0.9, 4.25, 1.2, "ink", 2)
    if d.get("subtitle"):
        text(s, (0.9, 4.45, 9.2, 0.8), d["subtitle"], 15, "muted")
    meta = [m for m in (d.get("org"), d.get("author"), d.get("date")) if m]
    if meta:
        text(s, (0.9, 6.2, 9.2, 0.7), meta, 11, "text", gap=2)


def f_divider(s, d, deck):
    rect(s, 0, 0, W, H, "ink")
    src = _src({"src": d.get("image")})
    if src:  # full-bleed photo under a dark veil
        picture(s, src, 0, 0, W, H)
        alpha(rect(s, 0, 0, W, H, "ink"), 0.72)
    if d.get("no"):
        text(s, (0.9, 2.0, 4, 1.0), d["no"], 54, "accent", bold=True, fit=False)
    text(s, (0.9, 3.1, 11.5, 1.0), d["title"], 32, "paper", bold=True, name="divider title")
    if d.get("items"):
        text(s, (0.9, 4.3, 11, 2.2), d["items"], 13, tint("ink", 0.6), gap=6, fit=False)


def f_toc(s, d, deck):
    items = d.get("items") or deck.get("sections", [])
    rect(s, 0, 0, 4.2, H, "ink")
    text(s, (0.9, 2.6, 3, 0.8), d.get("title", "목차"), 34, "paper", bold=True, fit=False)
    text(s, (0.9, 3.45, 3, 0.4), "CONTENTS", 12, "#9DB6FF", bold=True, fit=False)
    ncol = 1 if len(items) <= 6 else 2
    per = math.ceil(len(items) / ncol)
    cw = (W - 5.0 - 0.8 - 0.4 * (ncol - 1)) / ncol
    rh = min(0.95, 5.6 / per)
    top = (H - rh * per) / 2
    for i, it in enumerate(items):
        cx = 5.0 + (i // per) * (cw + 0.4)
        cy = top + (i % per) * rh
        text(s, (cx, cy, 0.9, rh), f"{i + 1:02d}", 24, "accent", bold=True, anchor="m",
             fit=False)
        text(s, (cx + 1.0, cy, cw - 1.0, rh), it, 17, "ink", bold=True, anchor="m",
             name="toc item")
        hline(s, cx, cy + rh, cw)


def f_closing(s, d, deck):
    rect(s, 0, 0, W, H, "ink")
    text(s, (0.9, 2.6, 11, 1.2), d.get("title", "감사합니다"), 36, "paper", bold=True,
         anchor="b", fit=False)
    if d.get("subtitle"):
        text(s, (0.9, 4.0, 11, 1.2), d["subtitle"], 14, tint("ink", 0.6))


def f_statement(s, d, deck):
    rect(s, 0, 0, W, H, tint("accent", 0.95))
    rect(s, 0.9, 2.45, 0.09, 2.2, "accent")
    if d.get("kicker"):
        text(s, (1.25, 1.9, 10, 0.4), d["kicker"], 13, "accent", bold=True, fit=False)
    text(s, (1.25, 2.4, 11, 2.3), d["title"], 30, "ink", bold=True, anchor="m", line=1.15,
         min_size=20, name="statement")
    if d.get("subtitle"):
        text(s, (1.25, 4.95, 11, 1.2), d["subtitle"], 14, "muted", name="statement sub")


def f_photo(s, d, deck):
    """Full-bleed image with a text veil — hero shots, site photos, product screens."""
    src = _src({"src": d.get("image")})
    if src:
        picture(s, src, 0, 0, W, H, d.get("crop", "center"))
    else:
        placeholder(s, 0, 0, W, H, d.get("alt"), dark=True)
    side = d.get("align", "left")
    pw = W * 0.42
    px = 0 if side == "left" else W - pw
    alpha(rect(s, px, 0, pw, H, "ink"), 0.86)
    tx = px + 0.7
    if d.get("kicker"):
        text(s, (tx, 1.6, pw - 1.2, 0.35), d["kicker"], 12, "#9DB6FF", bold=True, fit=False)
    text(s, (tx, 2.05, pw - 1.2, 2.4), d["title"], 28, "paper", bold=True, line=1.1,
         min_size=18, emph="#9DB6FF", name="photo title")
    if d.get("subtitle"):
        text(s, (tx, 4.6, pw - 1.2, 1.6), d["subtitle"], 13, tint("ink", 0.7), name="photo sub")
    if d.get("caption"):
        text(s, (tx, H - 0.8, pw - 1.2, 0.3), d["caption"], 9, tint("ink", 0.55), fit=False)


FULL = {"cover": f_cover, "divider": f_divider, "toc": f_toc, "statement": f_statement,
        "photo": f_photo, "closing": f_closing}


# ------------------------------------------------------------------ build

def build(deck, out, base=Path(".")):
    T.clear()
    T.update(THEME, **deck.get("theme", {}))
    WARN.clear()
    CTX["base"] = base
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(W), Inches(H)
    for n, d in enumerate(deck["slides"], 1):
        CTX["slide"] = n
        s = prs.slides.add_slide(prs.slide_layouts[6])
        lay = d.get("layout", "content")
        if lay in FULL:
            FULL[lay](s, d, deck)
        elif lay == "content":
            rect(s, 0, 0, W, H, "paper")
            top = header(s, d, deck)
            bottom = footer(s, d, deck, n) - (0.1 if d.get("source") else 0)
            grid(s, d.get("body", []), (M, top, W - 2 * M, bottom - top))
        else:
            sys.exit(f"slide {n}: unknown layout {lay!r}. Use content or one of {', '.join(FULL)}")
        if d.get("notes"):
            s.notes_slide.notes_text_frame.text = d["notes"]
    prs.save(out)
    return len(prs.slides)


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("deck")
    ap.add_argument("-o", "--out", default="deck.pptx")
    a = ap.parse_args()
    path = Path(a.deck)
    n = build(json.loads(path.read_text(encoding="utf-8")), a.out, path.parent)
    for w_ in WARN:
        print("WARN", w_, file=sys.stderr)
    print(f"wrote {a.out} ({n} slides, {len(WARN)} warnings)")

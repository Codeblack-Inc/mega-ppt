# /// script
# requires-python = ">=3.10"
# dependencies = ["python-pptx>=1.0"]
# ///
"""deck.json → .pptx. Usage: uv run build_deck.py deck.json -o out.pptx

A content slide = header (kicker · headline · lead) + body grid of panels.
Panels are plain functions in PANELS; full-bleed slides in FULL. Spec: references/layouts.md.
"""
import argparse
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
             gap=4, name="text"):
    while size > min_size and text_height(content, size, w, bullets, bold, line, gap) > h:
        size -= 0.5
    need = text_height(content, size, w, bullets, bold, line, gap)
    if need > h * 1.04:
        warn(f"{name!r} overflows its box even at {size}pt — shorten the text")
    elif h > 1.5 and need < h * 0.3:
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
        size = fit_size(content, size, w, h, bullets, bold, min_size, line, gap, name)
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


def hline(s, x, y, w, color="rule", lw=0.75):
    c = s.shapes.add_connector(1, Inches(x), Inches(y), Inches(x + w), Inches(y))
    c._element.remove(c._element.find(qn("p:style")))
    c.line.color.rgb = rgb(color)
    c.line.width = Pt(lw)
    return c


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
            fill = ("ink" if head else tint("accent", 0.9) if is_hl
                    else "soft" if rh_cell else "paper")
            last = r == nr - 1
            _cell_borders(cell, T=("ink" if r == 0 and not cols else None),
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
            color = "paper" if head else "accent" if is_hl else "ink" if rh_cell else "text"
            for k, part in enumerate(EMPH.split(val)):
                if part:
                    run = para.add_run()
                    run.text = part
                    _set_font(run, size, "accent" if k % 2 and not head else color,
                              head or rh_cell or is_hl or bool(k % 2))


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
        if numbered:
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


def p_image(s, b, p):
    x, y, w, h = b
    cap = p.get("caption")
    ih = h - (0.3 if cap else 0)
    src = p.get("src") and (CTX["base"] / p["src"])
    if src and src.exists():
        pic = s.shapes.add_picture(str(src), Inches(x), Inches(y), width=Inches(w))
        if pic.height > Inches(ih):
            ratio = Inches(ih) / pic.height
            pic.height, pic.width = Inches(ih), int(pic.width * ratio)
        pic.left = int(Inches(x) + (Inches(w) - pic.width) / 2)
    else:
        ph = rect(s, x, y, w, ih, "soft", "rule")
        ph.line.dash_style = 4
        shape_text(ph, "[이미지] " + p.get("alt", "이미지를 넣어 주세요"), 9, "muted", bold=False)
    if cap:
        text(s, (x, y + ih + 0.06, w, 0.24), f"< {cap} >", 9, "muted", align="c", fit=False)


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


PANELS = {
    "bullets": p_bullets, "callout": p_callout, "table": p_table, "chart": p_chart,
    "kpi": p_kpi, "cards": p_cards, "process": p_process, "timeline": p_timeline,
    "cycle": p_cycle, "stack": p_stack, "image": p_image, "label": p_label, "arrow": p_arrow,
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


def grid(s, body, box):
    """body = [row, ...]; row = [panel, ...] or {"h": weight, "cols": [...]}; panel "w" = weight."""
    x, y, w, h = box
    rows = [r if isinstance(r, dict) else {"cols": r} for r in body]
    hw = [r.get("h", 1) for r in rows]
    avail = h - GAP * (len(rows) - 1)
    cy = y
    for r, weight in zip(rows, hw):
        rh = avail * weight / sum(hw)
        cols = r["cols"]
        fixed = [FIXED_W[c["type"]] if "w" not in c and c["type"] in FIXED_W else None
                 for c in cols]
        flex = w - GAP * (len(cols) - 1) - sum(f for f in fixed if f)
        tot = sum(c.get("w", 1) for c, f in zip(cols, fixed) if f is None) or 1
        cx = x
        for c, f in zip(cols, fixed):
            cw = f if f else flex * c.get("w", 1) / tot
            panel(s, (cx, cy, cw, rh), c)
            cx += cw + GAP
        cy += rh + GAP


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


FULL = {"cover": f_cover, "divider": f_divider, "toc": f_toc, "closing": f_closing}


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

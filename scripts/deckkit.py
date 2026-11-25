# -*- coding: utf-8 -*-
"""NEWX PPT 绘制引擎：1280x720 画布，切分树布局、扁平渐变组件、原生表格与图表。
颜色由 palette.apply() 在运行时写入本模块的全局变量；组件一律在运行时读取，不要在参数默认值里引用颜色。"""
import math
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION

PX = 9525
MAIN, ACC, INK, MUTED = "0E5A78", "DD7A41", "1F2A30", "6B7780"
BG, LINE, SUB = "F7F8F9", "D5DCE0", "3C4A52"
ONDARK, TRACK, STRIPE, CARD_BORDER, CORNER = "CFE3EC", "DCE5EA", "F5F8F9", "F2F2F2", "square"
# 非重点数据用的灰（带一点主色色相）：重点 = 主色，其余 = 灰，不另设"强调色"做重点
NORM, GRAYS = "B8BEC2", ["7F878C", "B8BEC2", "D3D8DB", "E3E7E9"]
FONT = "Source Han Sans CN"  # Source Han Sans CN；对外交付时 public_output 统一换成 Noto Sans CJK SC

BODY_X, BODY_Y, BODY_W, BODY_H = 48, 116, 1184, 558

# 格内文字间距四档（由小到大，差别不大但要分得出）：
#   行间距 < 段内间距（导语与说明之间）< 段间距（条目之间）< 标题与正文间距（格内小标题到内容）
LS_LEAD, LS_BODY = 1.25, 1.42     # 导语行距、正文行距（倍数）
SP_INNER = 6                      # 段内：导语与说明之间（px）
SP_ITEM_MIN, SP_ITEM_MAX = 18, 26  # 段间：条目之间（px）
BAR_ROW = 44                      # 条形图每条占的高度（条约 22px＋间隙），图按条数取高
SP_HEAD = 20                      # 标题与正文：小标题文字底到内容顶（px），大于条目间距下限


def rgb(h):
    return RGBColor.from_string(h)


def mix(a, b, k):
    A = [int(a[i:i + 2], 16) for i in (0, 2, 4)]
    B = [int(b[i:i + 2], 16) for i in (0, 2, 4)]
    return "".join(f"{round(x + (y - x) * k):02X}" for x, y in zip(A, B))


DARK_G = (MAIN, mix(MAIN, "FFFFFF", .28), 315)
HEAD_G = (MAIN, mix(MAIN, "FFFFFF", .30), 0)
CELL_G = (mix(MAIN, "FFFFFF", .90), mix(MAIN, "FFFFFF", .975), 0)


def E(v):
    return Emu(int(round(v * PX)))


def flat(s):
    st = s._element.find(qn("p:style"))
    if st is not None:
        s._element.remove(st)


def _paint(s, fill):
    if isinstance(fill, tuple):
        s.fill.gradient()
        s.fill.gradient_angle = fill[2] if len(fill) > 2 else 0
        st = s.fill.gradient_stops
        st[0].color.rgb = rgb(fill[0]); st[0].position = 0.0
        st[1].color.rgb = rgb(fill[1]); st[1].position = 1.0
    else:
        s.fill.solid(); s.fill.fore_color.rgb = rgb(fill)
    s.line.fill.background()
    flat(s)
    return s


def rect(sl, x, y, w, h, fill):
    s = sl.shapes.add_shape(MSO_SHAPE.RECTANGLE, E(x), E(y), E(max(w, 1)), E(max(h, 1)))
    return _paint(s, fill)


def hline(sl, x, y, w, c=None, pt=0.75):
    c = c or LINE
    l = sl.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, E(x), E(y), E(x + w), E(y))
    l.line.color.rgb = rgb(c); l.line.width = Pt(pt); flat(l)


def vline(sl, x, y, h, c=None, pt=0.75):
    c = c or LINE
    l = sl.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, E(x), E(y), E(x), E(y + h))
    l.line.color.rgb = rgb(c); l.line.width = Pt(pt); flat(l)


# ---------------- 文本 ----------------
def est_lines(text, width, size):
    """粗略估算换行数：中文按 1.0 em，西文数字按 0.56 em。"""
    if not text:
        return 1
    units = 0.0
    for ch in text:
        o = ord(ch)
        if o < 128:
            units += 0.56 if (ch.isalnum() or ch in ".,%-+/") else 0.36
        else:
            units += 1.0
    per = max(width / size, 1.0)
    lines = 0
    for part in text.split("\n"):
        u = sum((0.56 if (ord(c) < 128 and (c.isalnum() or c in ".,%-+/")) else
                 (0.36 if ord(c) < 128 else 1.0)) for c in part)
        lines += max(1, math.ceil(u / per))
    return lines


def _units(t):
    return sum((0.56 if (ch.isalnum() or ch in ".,%-+/") else 0.36) if ord(ch) < 128 else 1.0 for ch in t)


_PUNCT_AFTER = "，、；：:,;—|｜ "
_CONJ = "与和及暨并"
_ORG = ("中心", "集团", "公司", "分行", "支行", "组", "部", "局", "处", "厅", "委", "院", "馆", "站", "所", "室", "科")
_NO_BEFORE = "，、。；：！？）》」』%％年月日号个项次元亿万"


def _break_score(t, i):
    """在 t[i] 前断行的代价：越小越像自然断句；None 表示不能在这里断。"""
    a, b = t[i - 1], t[i]
    if (a.isascii() and a.isalnum()) and (b.isascii() and b.isalnum()):
        return None  # 不拆数字、英文
    if b in _NO_BEFORE or a in "（《「『":
        return None
    if a in _PUNCT_AFTER:
        return 0
    if b.isdigit() and not a.isdigit():
        return 0  # 年份、数字前
    if b in _CONJ:
        return 1
    if any(t[:i].endswith(o) for o in _ORG):
        return 1  # 单位名之后
    return 5


def break_title(title, width, size, min_size, max_lines=2):
    """标题合理换行：先看一行能否放下；放不下时在标点、数字（年份）、连词、单位名之后断成两行，
    两行都要放得下，并尽量均衡；找不到自然断点才逐字断。返回 (行列表, 字号)。标题里写了换行符就照写的断。"""
    if "\n" in title:
        lines = [l for l in title.split("\n") if l]
        while size > min_size and any(_units(l) * size > width for l in lines):
            size -= 1
        return lines, size
    fallback = None
    for sz in range(int(size), int(min_size) - 1, -1):
        if _units(title) * sz <= width:
            return [title], sz
        if max_lines < 2:
            continue
        best = None
        for i in range(2, len(title) - 1):
            sc = _break_score(title, i)
            if sc is None:
                continue
            l1, l2 = title[:i], title[i:]
            if _units(l1) * sz > width or _units(l2) * sz > width:
                continue
            key = (sc, abs(_units(l1) - _units(l2)))
            if best is None or key < best[0]:
                best = (key, [l1, l2], sz)
        if best and best[0][0] <= 1:
            return best[1], best[2]
        if best and fallback is None:
            fallback = (best[1], best[2])
    if fallback:
        return fallback
    return [title], min_size


def est_height(paras, width, spacing=1.35):
    """paras: [(text, size)] 期望高度（px）。"""
    return sum(est_lines(t, width, s) * s * 1.22 * spacing for t, s in paras)


def text(sl, x, y, w, h, runs, align="l", anchor="t", spacing=1.3, wrap=True):
    """runs: str 或 [[(文字, 字号, 颜色, 粗体), ...], ...]，每个子列表一段。"""
    if isinstance(runs, str):
        runs = [[(runs, 16, INK, False)]]
    tb = sl.shapes.add_textbox(E(x), E(y), E(max(w, 1)), E(max(h, 1)))
    tf = tb.text_frame
    tf.word_wrap = wrap
    for m in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
        setattr(tf, m, 0)
    tf.vertical_anchor = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}[anchor]
    for i, para in enumerate(runs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}[align]
        p.line_spacing = spacing
        for t, size, color, bold in para:
            r = p.add_run(); r.text = t
            f = r.font
            f.size = Pt(size * 0.75); f.bold = bold; f.name = FONT
            f.color.rgb = rgb(color)
            rpr = r._r.get_or_add_rPr()
            ea = rpr.find(qn("a:ea"))
            if ea is None:
                ea = rpr.makeelement(qn("a:ea"), {}); rpr.append(ea)
            ea.set("typeface", FONT)
    return tb


def T(s, size=16, color=None, bold=False):
    return [[(s, size, color or INK, bold)]]


# ---------------- 版面框架 ----------------
def new_deck():
    prs = Presentation()
    prs.slide_width, prs.slide_height = Emu(1280 * PX), Emu(720 * PX)
    return prs


def add_slide(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    rect(sl, 0, 0, 1280, 720, BG)
    return sl


TITLE_RESERVE = 0  # 页眉右侧给 LOGO 让出的宽度（logo.use 设定）


def title_block(sl, title, footer):
    tw = 1184 - TITLE_RESERVE
    size = 40
    while est_lines(title, tw, size) > 1 and size > 28:
        size -= 1
    text(sl, 48, 34, tw, 70, T(title, size, MAIN, True), anchor="m", spacing=1.1)
    if footer:  # 来源左对齐占一行，页码单独居中放在下面一行，两者不相交
        fs = 11.5
        while est_lines(footer, 1184, fs) > 1 and fs > 10.5:
            fs -= 0.5
        text(sl, 48, 678, 1184, 18, T(footer, fs, MUTED), align="l", anchor="m")


def heading(sl, x, y, w, t, dark=False, size=19):
    size = size
    while est_lines(t, w - 4, size) > 1 and size > 14:
        size -= 1
    # 强调短竖条挂在左侧留白里，小标题文字与下方正文、图标、表格同一条左边线
    if not dark:
        rect(sl, x - 13, y + 8, 5, 22, (MAIN, mix(MAIN, "FFFFFF", .3), 90))
        text(sl, x, y, w, 38, T(t, size, MAIN, True), anchor="m")
    else:
        rect(sl, x - 13, y + 8, 5, 22, (mix("FFFFFF", MAIN, 0.55), "FFFFFF", 90))
        text(sl, x, y, w, 38, T(t, size, "FFFFFF", True), anchor="m")
    return 19 + size * 0.6 + SP_HEAD


# ---------------- 布局：切分树 ----------------
def layout(node, x, y, w, h, gap, leaves, seams):
    if "children" not in node:
        leaves.append((node, (x, y, w, h)))
        return
    cut, rs = node["cut"], node["ratio"]
    n = len(rs)
    tot = (w if cut == "v" else h) - gap * (n - 1)
    s = sum(rs)
    pos = x if cut == "v" else y
    for i, ch in enumerate(rs):
        size = tot * ch / s
        if cut == "v":
            layout(node["children"][i], pos, y, size, h, gap, leaves, seams)
            if i < n - 1:
                seams.append(("v", pos + size + gap / 2, y, h))
            pos += size + gap
        else:
            layout(node["children"][i], x, pos, w, size, gap, leaves, seams)
            if i < n - 1:
                seams.append(("h", pos + size + gap / 2, x, w))
            pos += size + gap


def check_tree(node, x=BODY_X, y=BODY_Y, w=BODY_W, h=BODY_H, depth=1, out=None):
    """返回 (叶子数, 最深递归层数, 违规列表)。"""
    out = out if out is not None else []
    if "children" not in node:
        if w < 260 or h < 150:
            out.append(f"格子过小 {w:.0f}x{h:.0f}")
        if w * h > 0.60 * BODY_W * BODY_H + 1:
            out.append(f"单格面积超 60% {w:.0f}x{h:.0f}={w*h/(BODY_W*BODY_H)*100:.1f}%")
        return 1, depth, out
    cut, rs = node["cut"], node["ratio"]
    n = len(rs)
    tot = (w if cut == "v" else h) - 0 * (n - 1)
    s = sum(rs)
    leaves = 0
    deepest = 0
    pos = x if cut == "v" else y
    for i, ch in enumerate(rs):
        size = tot * ch / s
        if cut == "v":
            l, d, out = check_tree(node["children"][i], pos, y, size, h, depth + 1, out)
        else:
            l, d, out = check_tree(node["children"][i], x, pos, w, size, depth + 1, out)
        leaves += l; deepest = max(deepest, d)
        pos += size
    return leaves, deepest, out


# ---------------- 组件 ----------------
def kpi_row(sl, x, y, w, h, items, dark=False, color=None, label_color=None):
    color = color or MAIN
    n = len(items)
    cw = w / n
    num_color = "FFFFFF" if dark else color
    numh = min(58, h * 0.52)
    # 分隔竖线只和数字＋说明一样高，不随格子拉长（格子比内容高时线会伸到空处）
    lab_h = max(est_lines(lab, cw - 40, 16) for _, _, lab in items) * 16 * 1.3 + 6
    line_h = min(h - 8, numh + 4 + lab_h)
    for i, (num, unit, label) in enumerate(items):
        cx = x + i * cw
        if i:
            vline(sl, cx, y + 4, line_h, "FFFFFF" if dark else LINE, 1)
        pad = 0 if i == 0 else 26
        bw = cw - pad - 10
        big = 44
        # 数字与单位挤在一行：按估算宽度收缩字号
        while big > 26 and (len(num) * big * 0.56 + (len(unit) + 1) * 18 * 0.9) > bw:
            big -= 2
        unit_sz = min(18, max(13, big * 0.4))
        text(sl, cx + pad, y, bw, numh,
             [[(num, big, num_color, True), (" " + unit, unit_sz, num_color, True)]],
             anchor="b", spacing=1.0)
        text(sl, cx + pad, y + numh + 4, bw - 4, h - numh - 4,
             T(label, 16, ONDARK if dark else (label_color or SUB)), spacing=1.3)


def points(sl, x, y, w, h, items, dark=False, badge=True, size=20, gap=14):
    """items: [[lead, detail], ...]；按内容高度聚拢，整组在格内垂直居中。"""
    avail = h
    def _need(sz):
        tot = 0
        for it in items:
            tot += est_lines(it[0], w - 46, sz) * sz * 1.22 * LS_LEAD
            if len(it) > 1 and it[1]:
                tot += SP_INNER + est_lines(it[1], w - 46, sz - 2) * (sz - 2) * 1.22 * LS_BODY
        return tot + SP_ITEM_MIN * (len(items) - 1)
    while size > 16 and _need(size) > avail:  # 导语不小于 16、说明不小于 14；再放不下就压条目间距
        size -= 0.5
    bw = 36 if badge else 0
    need = []
    for it in items:
        hh = est_lines(it[0], w - bw, size) * size * 1.22 * LS_LEAD
        if len(it) > 1 and it[1]:
            hh += SP_INNER + est_lines(it[1], w - bw, size - 2) * (size - 2) * 1.22 * LS_BODY
        need.append(hh)
    n = len(items)
    inner = (h - sum(need)) / n if n else 0
    gap = max(SP_ITEM_MIN, min(inner, SP_ITEM_MAX))
    tot = sum(need) + gap * (n - 1)
    cy = y + max(0, (h - tot) / 2)
    if tot > h:
        gap = max(4, (h - sum(need)) / max(n - 1, 1))
    for i, it in enumerate(items):
        lead, detail = it[0], (it[1] if len(it) > 1 else None)
        bx = x
        ic = it[2] if len(it) > 2 else None
        if ic:
            import icons as _I
            d = 34
            circ = sl.shapes.add_shape(MSO_SHAPE.OVAL, E(x), E(cy - 2), E(d), E(d))
            _paint(circ, "FFFFFF" if dark else mix(MAIN, "FFFFFF", .86))
            if _I.icon(sl, ic, x + 7, cy + 5, 20, MAIN if dark else MAIN):
                bx = x + 46
            else:
                ic = None
        if badge and not ic:
            rect(sl, x + 2, cy + 3, 24, 24, (MAIN, mix(MAIN, "FFFFFF", .3), 90))
            text(sl, x + 2, cy + 3, 24, 24, T(str(i + 1), 14, "FFFFFF", True), align="c", anchor="m")
            bx = x + 36
        lead_h = est_lines(lead, w - (bx - x), size) * size * 1.22 * LS_LEAD
        text(sl, bx, cy, w - (bx - x), lead_h + 4,
             T(lead, size, "FFFFFF" if dark else INK, True), spacing=LS_LEAD)
        if detail:
            dy = cy + lead_h + SP_INNER
            text(sl, bx, dy, w - (bx - x), max(10, need[i] - lead_h - SP_INNER + 4),
                 T(detail, size - 2, ONDARK if dark else SUB), spacing=LS_BODY)
        cy += need[i] + gap


def bar_cards(sl, x, y, w, h, rows, dark=False, size=15):
    """rows: [[label, value_text, value(数值), note]]，横向条。"""
    rh = h / len(rows)
    vmax = max(r[2] for r in rows) if rows else 1
    for i, r in enumerate(rows):
        lab, vt, v, note = r[0], r[1], r[2], (r[3] if len(r) > 3 else "")
        yy = y + i * rh
        text(sl, x, yy, 92, rh, T(lab, size, "FFFFFF" if dark else INK, True), anchor="m")
        bw_max = w - 92 - 118
        bw = max(8, bw_max * (v / vmax))
        rect(sl, x + 96, yy + rh / 2 - 7, bw_max, 14, TRACK)
        rect(sl, x + 96, yy + rh / 2 - 7, bw, 14, (MAIN, mix(MAIN, "FFFFFF", .35), 0))
        text(sl, x + 96 + bw_max + 10, yy, 108, rh,
             [[(vt, size, MAIN, True), ("  " + note, size - 3, MUTED, False)]], anchor="m")


# ---------------- 原生表格 ----------------
def _cell_border(cell, edges=("B",), c=None, pt=0.75):
    c = c or LINE
    tcPr = cell._tc.get_or_add_tcPr()
    tag = {"L": "a:lnL", "R": "a:lnR", "T": "a:lnT", "B": "a:lnB"}
    for e in ("L", "R", "T", "B"):
        old = tcPr.find(qn(tag[e]))
        if old is not None:
            tcPr.remove(old)
    for e in ("L", "R", "T", "B"):
        ln = tcPr.makeelement(qn(tag[e]), {"w": str(int(pt * 12700)), "cap": "flat",
                                           "cmpd": "sng", "algn": "ctr"})
        if e in edges:
            fill = ln.makeelement(qn("a:solidFill"), {})
            clr = fill.makeelement(qn("a:srgbClr"), {"val": c})
            fill.append(clr); ln.append(fill)
        else:
            ln.append(ln.makeelement(qn("a:noFill"), {}))
        tcPr.append(ln)


def table(sl, x, y, w, h, header, rows, colw=None, hl=None, size=15, hsize=15,
          head_fill=None, dark=False, first_bold=True):
    nrow, ncol = len(rows) + 1, len(header)
    shp = sl.shapes.add_table(nrow, ncol, E(x), E(y), E(w), E(h))
    tbl = shp.table
    tbl.first_row = False
    tbl.horz_banding = False
    tbl.vert_banding = False
    colw = colw or [1] * ncol
    tot = sum(colw)
    for i, cw in enumerate(colw):
        tbl.columns[i].width = E(w * cw / tot)
    rh = h / nrow
    for i in range(nrow):
        tbl.rows[i].height = E(rh)
    head_fill = head_fill or HEAD_G
    # 文字列左对齐，数字列居中：一列里过半单元格以数字开头且较短，算数字列
    import re as _re
    numeric_col = [sum(1 for r in rows if _re.match(r"^[≥≤<>+\-−±约~]*[\d.]", str(r[ci]).strip()) and len(str(r[ci])) <= 12)
                   > len(rows) / 2 for ci in range(ncol)]
    for ri in range(nrow):
        vals = header if ri == 0 else rows[ri - 1]
        is_hl = (hl is not None and ri - 1 == hl)
        for ci in range(ncol):
            cell = tbl.cell(ri, ci)
            cell.margin_left = E(14); cell.margin_right = E(10)
            cell.margin_top = E(2); cell.margin_bottom = E(2)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            # 整行连续渐变：按列宽把一条渐变切段分给各格，避免"每格一段"的条纹感
            k0, k1 = sum(colw[:ci]) / tot, sum(colw[:ci + 1]) / tot
            if ri == 0:
                _paint_cell(cell, (mix(head_fill[0], head_fill[1], k0), mix(head_fill[0], head_fill[1], k1), 0))
                col, bold, fs = "FFFFFF", True, hsize
            elif is_hl:
                _hl0, _hl1 = mix(MAIN, "FFFFFF", .88), mix(MAIN, "FFFFFF", .95)
                _paint_cell(cell, (mix(_hl0, _hl1, k0), mix(_hl0, _hl1, k1), 0))
                col, bold, fs = MAIN, True, size
            else:
                _paint_cell(cell, "FFFFFF" if ri % 2 else STRIPE)
                col, bold, fs = INK, (ci == 0 and first_bold), size
            tf = cell.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT if (ci == 0 or not numeric_col[ci]) else PP_ALIGN.CENTER
            p.line_spacing = 1.1
            r = p.add_run(); r.text = vals[ci]
            f = r.font; f.size = Pt(fs * 0.75); f.bold = bold; f.name = FONT
            f.color.rgb = rgb(col)
            rpr = r._r.get_or_add_rPr()
            ea = rpr.makeelement(qn("a:ea"), {}); ea.set("typeface", FONT); rpr.append(ea)
            _cell_border(cell, ("B",) if ri else ("B",), LINE if ri else mix(MAIN, "FFFFFF", .15))
    return shp


def _paint_cell(cell, fill):
    cell.fill.gradient() if isinstance(fill, tuple) else cell.fill.solid()
    if isinstance(fill, tuple):
        cell.fill.gradient_angle = fill[2]
        st = cell.fill.gradient_stops
        st[0].color.rgb = rgb(fill[0]); st[0].position = 0.0
        st[1].color.rgb = rgb(fill[1]); st[1].position = 1.0
    else:
        cell.fill.fore_color.rgb = rgb(fill)


# ---------------- 原生图表 ----------------
def _chart_base(ch):
    ch.has_legend = False
    ch.has_title = False
    ch.font.size = Pt(11)
    ch.font.color.rgb = rgb(MUTED)
    ch.font.name = FONT


def _axis_style(ax, size=11, grid=True):
    ax.format.line.color.rgb = rgb(LINE)
    ax.format.line.width = Pt(0.75)
    if ax.has_major_gridlines:
        if grid:
            ax.major_gridlines.format.line.color.rgb = rgb("E3E9ED")
            ax.major_gridlines.format.line.width = Pt(0.75)
        else:
            ax.has_major_gridlines = False
    try:
        ax.tick_labels.font.size = Pt(size)
        ax.tick_labels.font.color.rgb = rgb(MUTED)
    except Exception:
        pass


def _manual_layout(ch, x, y, w, h):
    """固定绘图区，便于自己画标签。x/y/w/h 为相对图表框的比例。"""
    pa = ch._chartSpace.find(qn("c:chart")).find(qn("c:plotArea"))
    lay = pa.find(qn("c:layout"))
    if lay is None:
        lay = pa.makeelement(qn("c:layout"), {})
        pa.insert(0, lay)
    for tag in ("c:manualLayout",):
        old = lay.find(qn(tag))
        if old is not None:
            lay.remove(old)
    ml = lay.makeelement(qn("c:manualLayout"), {})
    def add(parent, tag, val, attr="val"):
        e = parent.makeelement(qn(tag), {attr: val}); parent.append(e)
    add(ml, "c:layoutTarget", "inner")
    add(ml, "c:xMode", "edge"); add(ml, "c:yMode", "edge")
    add(ml, "c:x", f"{x:.4f}"); add(ml, "c:y", f"{y:.4f}")
    add(ml, "c:w", f"{w:.4f}"); add(ml, "c:h", f"{h:.4f}")
    lay.append(ml)
    return (x, y, w, h)


def chart_waterfall(sl, x, y, w, h, steps, unit="万元", lab_size=11, note=True):
    """steps: [[标签, 数值, 'total'|'delta'], ...]
    原生堆积柱模拟瀑布（数据可编辑），绘图区固定，标签自绘。"""
    cats, base, val, colors = [], [], [], []
    run = 0.0
    pts = []
    for lab, v, kind in steps:
        cats.append(lab)
        if kind == "total":
            base.append(v); val.append(0.0); run = v
            colors.append(None)
        else:
            if v >= 0:
                base.append(run); val.append(v); run += v
                colors.append(mix(MAIN, "FFFFFF", .30))
            else:
                base.append(run + v); val.append(-v); run += v
                colors.append(ACC)
        pts += [base[-1], base[-1] + val[-1]]
    lo, hi = min(pts), max(pts)
    span = hi - lo
    min_scale, truncated = 0.0, False
    if (lo / hi) > 0.55 and span < hi * 0.6:
        truncated = True
        min_scale = math.floor((lo - span * 0.10) / 10) * 10
    if truncated:
        for i, (lab, v, kind) in enumerate(steps):
            if kind == "total":
                base[i] = min_scale; val[i] = v - min_scale
    max_scale = math.ceil((hi + span * 0.14) / 10) * 10

    cd = CategoryChartData()
    cd.categories = cats
    cd.add_series("基线", base)
    cd.add_series(unit, val)
    gf = sl.shapes.add_chart(XL_CHART_TYPE.COLUMN_STACKED, E(x), E(y), E(w), E(h), cd)
    ch = gf.chart
    _chart_base(ch)
    ch.plots[0].gap_width = int(min(300, max(60, (w * 0.86 / max(len(steps), 1) / 46 - 1) * 100)))  # 柱宽不超过约 46px
    ch.plots[0].overlap = 100
    s0, s1 = ch.plots[0].series
    s0.format.fill.background()
    s0.format.line.fill.background()
    s1.format.line.fill.background()
    for i, c in enumerate(colors):
        if c:
            s1.points[i].format.fill.solid()
            s1.points[i].format.fill.fore_color.rgb = rgb(c)
        else:
            s1.points[i].format.fill.solid()
            s1.points[i].format.fill.fore_color.rgb = rgb(MAIN)
    va = ch.value_axis
    va.minimum_scale = min_scale
    va.maximum_scale = max_scale
    va.major_unit = max(10, round((max_scale - min_scale) / 4, -1))
    _axis_style(va, lab_size, grid=True)
    va.has_major_gridlines = True
    ca = ch.category_axis
    _axis_style(ca, lab_size + 1, grid=False)

    fx, fy, fw, fh = _manual_layout(ch, 0.10, 0.05, 0.86, 0.70)
    px, py = x + fx * w, y + fy * h
    pw, ph = fw * w, fh * h
    n = len(steps)
    for i, (lab, v, kind) in enumerate(steps):
        top = max(base[i], base[i] + val[i])
        ry = py + ph * (1 - (top - min_scale) / (max_scale - min_scale))
        cx = px + pw * (i + 0.5) / n
        txt = f"{v:,.2f}" if kind == "total" else (f"+{abs(v):,.2f}" if v > 0 else f"−{abs(v):,.2f}")
        text(sl, cx - 60, ry - 26, 120, 20, T(txt, lab_size + 1, ACC if colors[i] == ACC else MAIN, True),
             align="c", anchor="b")
    if truncated and note:
        text(sl, x, y + h + 2, w - 4, 16,
             T(f"注：纵轴自 {min_scale:,.0f} 起，只用于看清增量幅度", 11, MUTED), align="r")
    return gf, run


def chart_bar_h(sl, x, y, w, h, cats, vals, fmt='0.00"%"', color=None, lab_size=11,
                hot=None, unit="", maxv=None):
    color = color or MAIN
    cd = CategoryChartData()
    cd.categories = cats
    cd.add_series(unit or "值", vals)
    gf = sl.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, E(x), E(y), E(w), E(h), cd)
    ch = gf.chart
    _chart_base(ch)
    ch.plots[0].gap_width = 55
    s = ch.plots[0].series[0]
    s.format.fill.solid(); s.format.fill.fore_color.rgb = rgb(color)
    s.format.line.fill.background()
    for i, v in enumerate(vals):
        hot_i = (hot is not None and i in hot)
        p = s.points[i]
        p.format.fill.solid()
        p.format.fill.fore_color.rgb = rgb(MAIN if hot_i else (NORM if hot else color))
        p.format.line.fill.background()
        dl = p.data_label
        dl.position = XL_LABEL_POSITION.OUTSIDE_END
        dl.number_format = fmt
        dl.number_format_is_linked = False
        r = dl.text_frame.paragraphs[0].runs[0] if dl.text_frame.paragraphs[0].runs else None
        dl.font.size = Pt(lab_size + 1); dl.font.bold = True
        dl.font.color.rgb = rgb(MAIN if hot_i else (MUTED if hot else MAIN))
    va = ch.value_axis
    va.minimum_scale = 0
    if maxv:
        va.maximum_scale = maxv
    _axis_style(va, lab_size, grid=True)
    va.visible = False
    ca = ch.category_axis
    _axis_style(ca, lab_size + 1, grid=False)
    ca.format.line.fill.background()
    return gf


def _lum(c):
    r, g, b = (int(c[i:i + 2], 16) / 255 for i in (0, 2, 4))
    return 0.299 * r + 0.587 * g + 0.114 * b


def chart_donut(sl, x, y, w, h, cats, vals, colors=None, center=None, lab_size=11):
    """环形图：扇区上标百分比，图表自带图例写名称。标注都是图表原生元素，改数据后自动更新。
    colors 可只给前几项，其余按默认灰阶补齐；扇区百分比字色按填充明暗取白或深色；小于 5% 的扇区不标。"""
    n = len(vals)
    base = [MAIN] + list(GRAYS) + [mix(GRAYS[-1], "FFFFFF", .5)]
    while len(base) < n:
        base.append(mix(base[-1], "FFFFFF", .4))
    colors = list(colors or [])[:n]
    colors += base[len(colors):n]
    total = float(sum(vals)) or 1.0
    pct = [v / total * 100 for v in vals]
    side = w >= h * 1.35                      # 宽格图例在右，否则在下
    if side:                                   # 图表框收窄到“环 + 图例”，图例自动竖排且紧贴环
        d0 = min(h * 0.94, w * 0.55)
        lw_ = max(len(str(c)) for c in cats) * (lab_size + 2) * 1.4 + 50
        w_new = min(w, d0 + lw_ + 30)
        x = x + (w - w_new) / 2               # “环 + 图例”整组在格内水平居中，不把空白全甩到右边
        w = w_new
    cd = CategoryChartData()
    cd.categories = cats
    cd.add_series("结构", vals)
    gf = sl.shapes.add_chart(XL_CHART_TYPE.DOUGHNUT, E(x), E(y), E(w), E(h), cd)
    ch = gf.chart
    _chart_base(ch)
    pl = ch.plots[0]
    ser = pl.series[0]
    pl.has_data_labels = True
    dls = pl.data_labels
    dls.show_percentage = True; dls.show_value = False; dls.show_category_name = False
    dls.number_format = '0%'; dls.number_format_is_linked = False
    dls.font.size = Pt(lab_size + 3); dls.font.bold = True
    for i, c in enumerate(colors):
        pt_ = ser.points[i]
        pt_.format.fill.solid(); pt_.format.fill.fore_color.rgb = rgb(c)
        pt_.format.line.color.rgb = rgb(BG); pt_.format.line.width = Pt(2.5)
        dl = pt_.data_label
        if pct[i] < 5:
            dLbl = dl._get_or_add_dLbl()
            for child in list(dLbl):
                if child.tag != qn("c:idx"):
                    dLbl.remove(child)
            dLbl.append(dLbl.makeelement(qn("c:delete"), {"val": "1"}))
        else:
            dl.font.size = Pt(lab_size + 3); dl.font.bold = True
            dl.font.color.rgb = rgb("FFFFFF" if _lum(c) < 0.6 else INK)
            dLbl = dl._get_or_add_dLbl()           # python-pptx 不提供单点 showPercent，直接改 XML
            for tag, val in (("c:showVal", "0"), ("c:showPercent", "1"), ("c:showCatName", "0")):
                e = dLbl.find(qn(tag))
                if e is not None:
                    e.set("val", val)
            if dLbl.find(qn("c:numFmt")) is None:
                dLbl.find(qn("c:idx")).addnext(dLbl.makeelement(qn("c:numFmt"), {"formatCode": "0%", "sourceLinked": "0"}))
    ser.format.line.color.rgb = rgb(BG)
    ch.has_legend = True
    lg = ch.legend
    lg.position = XL_LEGEND_POSITION.RIGHT if side else XL_LEGEND_POSITION.BOTTOM
    lg.include_in_layout = False
    lg.font.size = Pt(lab_size + 2); lg.font.color.rgb = rgb(INK)
    # 环的位置：图例在右时环靠左占高度，图例在下时环在上方居中
    if side:
        d = min(d0, w * 0.96)
        _manual_layout(ch, 0.02, (1 - d / h) / 2, d / w, d / h)
        dx, dy = x + w * 0.02, y + (h - d) / 2
    else:
        rows = math.ceil(n / 2)
        d = max(60, min(w * 0.94, h - 28 * rows - 16))
        _manual_layout(ch, (1 - d / w) / 2, 0.02, d / w, d / h)
        dx, dy = x + (w - d) / 2, y + h * 0.02
    hole = pl._element.find(qn("c:holeSize"))
    if hole is None:
        hole = pl._element.makeelement(qn("c:holeSize"), {})
        pl._element.append(hole)
    hole.set("val", "62")
    if center and d >= 140:
        cw_ = d * 0.5
        cx, cy = dx + d / 2, dy + d / 2
        text(sl, cx - cw_ / 2, cy - 34, cw_, 22, T(center[0], 11, MUTED), align="c", anchor="b")
        text(sl, cx - cw_ / 2, cy - 10, cw_, 32, T(center[1], 20, MAIN, True), align="c", anchor="t")
    return gf


def chart_bar_h(sl, x, y, w, h, cats, vals, suffix="", decimals=0, color=None, lab_size=11,
                hot=None, unit="", maxv=None, gap_width=None):
    color = color or MAIN
    if gap_width is None:  # 柱子约 22px 粗：按每类高度反推间隙
        band = h * 0.86 / max(len(cats), 1)
        gap_width = int(min(300, max(70, (band / 22 - 1) * 100)))
    cd = CategoryChartData()
    cd.categories = cats
    cd.add_series(unit or "值", vals)
    gf = sl.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, E(x), E(y), E(w), E(h), cd)
    ch = gf.chart
    _chart_base(ch)
    ch.plots[0].gap_width = gap_width
    s = ch.plots[0].series[0]
    s.format.fill.solid(); s.format.fill.fore_color.rgb = rgb(color)
    s.format.line.fill.background()
    for i, v in enumerate(vals):
        hot_i = (hot is not None and i in hot)
        p = s.points[i]
        p.format.fill.solid()
        p.format.fill.fore_color.rgb = rgb(MAIN if hot_i else (NORM if hot else color))
        p.format.line.fill.background()
        dl = p.data_label
        dl.position = XL_LABEL_POSITION.OUTSIDE_END
        dl.has_text_frame = True
        txt = f"{v:,.{decimals}f}{suffix}"
        dl.text_frame.text = txt
        r = dl.text_frame.paragraphs[0].runs[0]
        r.font.size = Pt(lab_size + 1); r.font.bold = True
        r.font.color.rgb = rgb(MAIN if hot_i else (MUTED if hot else MAIN))
    va = ch.value_axis
    va.minimum_scale = 0
    if maxv:
        va.maximum_scale = maxv
    _axis_style(va, lab_size, grid=False)
    va.visible = False
    ca = ch.category_axis
    _axis_style(ca, lab_size + 1, grid=False)
    ca.format.line.fill.background()
    return gf


def chart_column(sl, x, y, w, h, cats, vals, suffix="", decimals=0, lab_size=11, hot=None, unit="", maxv=None):
    """柱形图（竖向）：适合时间序列或少量类别对比。有重点时重点柱主色、其余灰；柱宽约 46px 以内。"""
    band = w * 0.88 / max(len(cats), 1)
    gap_width = int(min(300, max(60, (band / 46 - 1) * 100)))
    cd = CategoryChartData()
    cd.categories = cats
    cd.add_series(unit or "值", vals)
    gf = sl.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, E(x), E(y), E(w), E(h), cd)
    ch = gf.chart
    _chart_base(ch)
    ch.plots[0].gap_width = gap_width
    s = ch.plots[0].series[0]
    s.format.fill.solid(); s.format.fill.fore_color.rgb = rgb(MAIN)
    s.format.line.fill.background()
    for i, v in enumerate(vals):
        hot_i = (hot is not None and i in hot)
        p = s.points[i]
        p.format.fill.solid()
        c = MAIN if (hot_i or not hot) else NORM
        p.format.fill.gradient(); p.format.fill.gradient_angle = 90
        st = p.format.fill.gradient_stops
        st[0].color.rgb = rgb(mix(c, "FFFFFF", .25)); st[0].position = 0
        st[1].color.rgb = rgb(c); st[1].position = 1.0
        p.format.line.fill.background()
        dl = p.data_label
        dl.position = XL_LABEL_POSITION.OUTSIDE_END
        dl.has_text_frame = True
        dl.text_frame.text = f"{v:,.{decimals}f}{suffix}"
        r = dl.text_frame.paragraphs[0].runs[0]
        r.font.size = Pt(lab_size + 1); r.font.bold = True
        r.font.color.rgb = rgb(MAIN if (hot_i or not hot) else MUTED)
    va = ch.value_axis
    va.minimum_scale = 0
    if maxv:
        va.maximum_scale = maxv
    _axis_style(va, lab_size, grid=False)
    va.visible = False
    ca = ch.category_axis
    _axis_style(ca, lab_size + 1, grid=False)
    return gf


def chart_line(sl, x, y, w, h, cats, series, suffix="", decimals=0, lab_size=11, hot=None, vmin=None, vmax=None):
    """折线图：看趋势。series = [{"name": "2025 年", "vals": [...]}, ...]，第一条为重点线（主色），
    其余为对比线（灰，第二条虚线）。只有一条线且点数不多时每个点标数值；多条线只在末端标数值。
    hot：重点点的序号列表，放大标记并标数值。"""
    if series and not isinstance(series[0], dict):
        series = [{"name": "值", "vals": series}]
    cd = CategoryChartData()
    cd.categories = cats
    for s_ in series:
        cd.add_series(s_.get("name") or "值", s_["vals"])
    gf = sl.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS, E(x), E(y), E(w), E(h), cd)
    ch = gf.chart
    _chart_base(ch)
    multi = len(series) > 1
    if multi:
        ch.has_legend = True
        ch.legend.position = XL_LEGEND_POSITION.TOP
        ch.legend.include_in_layout = False
        ch.legend.font.size = Pt(lab_size)
        ch.legend.font.color.rgb = rgb(SUB)
    from pptx.enum.chart import XL_MARKER_STYLE
    from pptx.enum.dml import MSO_LINE_DASH_STYLE
    for k, s in enumerate(ch.plots[0].series):
        c = MAIN if k == 0 else (NORM if k == 1 else GRAYS[min(k - 2, len(GRAYS) - 1)])
        s.smooth = False
        s.format.line.color.rgb = rgb(c)
        s.format.line.width = Pt(2.75 if k == 0 else 1.75)
        if k == 1:
            s.format.line.dash_style = MSO_LINE_DASH_STYLE.DASH
        s.marker.style = XL_MARKER_STYLE.CIRCLE
        s.marker.size = 7 if k == 0 else 5
        s.marker.format.fill.solid(); s.marker.format.fill.fore_color.rgb = rgb("FFFFFF")
        s.marker.format.line.color.rgb = rgb(c); s.marker.format.line.width = Pt(1.75)
        vals = series[k]["vals"]
        n = len(vals)
        for i, v in enumerate(vals):
            hot_i = k == 0 and hot is not None and i in hot
            last = i == n - 1
            label = (k == 0 and not multi and n <= 8) or last or hot_i
            if hot_i:
                p = s.points[i]
                p.marker.style = XL_MARKER_STYLE.CIRCLE; p.marker.size = 11
                p.marker.format.fill.solid(); p.marker.format.fill.fore_color.rgb = rgb(MAIN)
                p.marker.format.line.color.rgb = rgb("FFFFFF")
            if label and v is not None:
                dl = s.points[i].data_label
                dl.position = XL_LABEL_POSITION.ABOVE if k == 0 else XL_LABEL_POSITION.BELOW
                dl.has_text_frame = True
                dl.text_frame.text = f"{v:,.{decimals}f}{suffix}"
                r = dl.text_frame.paragraphs[0].runs[0]
                r.font.size = Pt(lab_size + (1 if k == 0 else 0)); r.font.bold = k == 0
                r.font.color.rgb = rgb(MAIN if k == 0 else MUTED)
    va = ch.value_axis
    allv = [v for s_ in series for v in s_["vals"] if v is not None]
    if vmin is None and allv:
        lo, hi = min(allv), max(allv)
        if lo >= 0 and lo < hi * 0.4:
            vmin = 0
        else:  # 趋势图不从 0 起，纵轴下限取略低于最小值的整数档，避免线挤在顶部、下方大片空白
            step = 10 ** math.floor(math.log10(max(hi - lo, 1e-9)))
            step = step / 2 if (hi - lo) / step < 3 else step
            vmin = math.floor((lo - (hi - lo) * 0.25) / step) * step
            if lo >= 0:
                vmin = max(0, vmin)
    if vmin is not None:
        va.minimum_scale = vmin
    if vmax is not None:
        va.maximum_scale = vmax
    _axis_style(va, lab_size, grid=True)
    va.format.line.fill.background()
    ca = ch.category_axis
    _axis_style(ca, lab_size + 1, grid=False)
    return gf


# ================= 2026-11-17 增补：卡片规格 + 格内关系图示 =================
def soft_card(s, border=None):
    border = border or CARD_BORDER
    """无边框卡片的柔阴影：0.5pt 近白描边 + 大模糊低透明阴影（用户手调参数）。"""
    s.line.color.rgb = rgb(border); s.line.width = Pt(0.5)
    spPr = s._element.spPr
    old = spPr.find(qn("a:effectLst"))
    if old is not None:
        spPr.remove(old)
    eff = spPr.makeelement(qn("a:effectLst"), {})
    sh = eff.makeelement(qn("a:outerShdw"), {"blurRad": "469900", "dist": "38100", "dir": "2700000",
                                            "algn": "tl", "rotWithShape": "0"})
    clr = sh.makeelement(qn("a:prstClr"), {"val": "black"})
    a = clr.makeelement(qn("a:alpha"), {"val": "7000"})
    clr.append(a); sh.append(clr); eff.append(sh); spPr.append(eff)
    return s


def card(sl, x, y, w, h, fill=None):
    if CORNER == "round":
        s = sl.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, E(x), E(y), E(w), E(h))
        s.adjustments[0] = min(0.5, 10 / max(1, min(w, h)))
        _paint(s, fill or CELL_G)
    else:
        s = rect(sl, x, y, w, h, fill or CELL_G)
    return soft_card(s)


def _shape(sl, kind, x, y, w, h, fill):
    s = sl.shapes.add_shape(kind, E(x), E(y), E(w), E(h))
    return _paint(s, fill)


STEP_PAD = 14  # steps 色块内左右留白
STEP_GAP = 8   # 图标与文字的间距


def _text_w(t, size):
    """单行文字宽度估算（中文 1 em，西文数字 0.56 em）。"""
    return sum((0.56 if (c.isalnum() or c in ".,%-+/") else 0.36) if ord(c) < 128 else 1.0 for c in t or "") * size


def steps(sl, x, y, w, h, items, dark=False, size=19):
    """递进：[[节点名, 说明], ...] 横向节点 + 箭头。说明可为空。"""
    n = len(items)
    arw = 22
    nw = (w - arw * (n - 1)) / n
    has_sub = any(len(it) > 1 and it[1] for it in items)
    nh = min(54, h * (0.38 if has_sub else 0.9))
    nh = max(nh, min(44, h))  # 色块不能扁到图标贴上下边
    sub_h = h - nh - 10 if has_sub else 0
    sub_fs = 18
    if has_sub:
        while sub_fs > 12 and max(est_height([(it[1], sub_fs)], nw, 1.35) for it in items if len(it) > 1 and it[1]) > sub_h:
            sub_fs -= 0.5
        sub_need = max(est_height([(it[1], sub_fs)], nw, 1.35) for it in items if len(it) > 1 and it[1]) + 6
    tot = nh + (12 + min(sub_h, sub_need) if has_sub else 0)
    y0 = y + (h - tot) / 2
    for i, it in enumerate(items):
        nx = x + i * (nw + arw)
        if dark:
            fill = ("FFFFFF", mix(MAIN, "FFFFFF", .85), 0)
            fg = MAIN
        else:
            k = .0 + .18 * i
            fill = (mix(MAIN, "FFFFFF", k), mix(MAIN, "FFFFFF", k + .25), 0)
            fg = "FFFFFF"
        _shape(sl, MSO_SHAPE.RECTANGLE, nx, y0, nw, nh, fill)
        # 色块里的图标和文字四周都要留白：左右各 STEP_PAD，图标与色块上下边至少 10，不贴边、不挤
        inner = nw - 2 * STEP_PAD
        ic = it[2] if len(it) > 2 else None
        isz = min(22, nh - 20)
        fs = size
        if ic:
            while fs > 13 and _text_w(it[0], fs) + isz + STEP_GAP > inner:
                fs -= 0.5
            if _text_w(it[0], fs) + isz + STEP_GAP > inner:
                ic, fs = None, size  # 字号降到 13 仍放不下：去掉图标，只留文字
        if not ic:
            while est_lines(it[0], inner, fs) > 1 and fs > 11:
                fs -= 0.5
        if ic:
            import icons as _I
            lw_ = _text_w(it[0], fs)
            gx = nx + (nw - (isz + STEP_GAP + lw_)) / 2
            if _I.icon(sl, ic, gx, y0 + (nh - isz) / 2, isz, fg):
                text(sl, gx + isz + STEP_GAP, y0, lw_ + 4, nh, T(it[0], fs, fg, True), anchor="m")
            else:
                ic = None
        if not ic:
            text(sl, nx + STEP_PAD, y0, inner, nh, T(it[0], fs, fg, True), align="c", anchor="m")
        if i < n - 1:
            ax = nx + nw + 5
            _shape(sl, MSO_SHAPE.CHEVRON, ax, y0 + nh / 2 - 8, arw - 10, 16,
                   "FFFFFF" if dark else mix(MAIN, "FFFFFF", .55))
        if has_sub and len(it) > 1 and it[1]:
            text(sl, nx + 4, y0 + nh + 12, nw - 8, min(sub_h, sub_need),
                 T(it[1], sub_fs, ONDARK if dark else SUB), align="c", spacing=1.35)


def chain(sl, x, y, w, h, rows, dark=False, size=15):
    """因果链：[[对象, 现象, 判断], ...]，每行  对象 | 现象 → 判断。"""
    n = len(rows)
    rh = min(h / n, 92)
    y0 = y + (h - rh * n) / 2
    lw = 64
    for i, (who, what, so) in enumerate(rows):
        yy = y0 + i * rh
        if i:
            hline(sl, x, yy, w, mix(MAIN, "FFFFFF", .35) if dark else LINE, 0.75)
        cy = yy + 14
        _shape(sl, MSO_SHAPE.RECTANGLE, x, cy + 2, lw, 28,
               "FFFFFF" if dark else (MAIN, mix(MAIN, "FFFFFF", .3), 0))
        text(sl, x, cy + 2, lw, 28, T(who, 14, MAIN if dark else "FFFFFF", True), align="c", anchor="m")
        tx = x + lw + 14
        text(sl, tx, cy, w - lw - 14, 30, T(what, size - 1, ONDARK if dark else SUB), anchor="m")
        text(sl, tx, cy + 32, w - lw - 14, rh - 46,
             [[("→ ", size, MAIN if not dark else mix(MAIN, "FFFFFF", .55), True), (so, size, "FFFFFF" if dark else INK, True)]],
             anchor="t", spacing=1.25)


def compare(sl, x, y, w, h, rows, note=None, dark=False, size=15):
    """占比对比：[[名称, 文本值, 数值(0-100)], ...] 细条 + 一行注记。"""
    n = len(rows)
    lab_w, val_w = 76, 78
    bw = w - lab_w - val_w - 12
    rh = 34
    nh = 44 if note else 0
    tot = rh * n + nh
    y0 = y + (h - tot) / 2
    vmax = max(r[2] for r in rows)
    for i, (lab, vt, v) in enumerate(rows):
        yy = y0 + i * rh
        hot = (v == vmax)
        text(sl, x, yy, lab_w, rh, T(lab, size - 1, "FFFFFF" if dark else INK, False), anchor="m")
        rect(sl, x + lab_w, yy + rh / 2 - 6, bw, 12, mix(MAIN, "FFFFFF", .3) if dark else TRACK)
        c = ("FFFFFF" if dark else MAIN) if hot else (mix(MAIN, "FFFFFF", .55) if dark else NORM)
        rect(sl, x + lab_w, yy + rh / 2 - 6, bw * v / (vmax * 1.1),
             12, (c, mix(c, "FFFFFF", .3), 0))
        text(sl, x + lab_w + bw + 12, yy, val_w, rh,
             T(vt, size, ("FFFFFF" if dark else MAIN) if hot else (ONDARK if dark else MUTED), True), anchor="m")
    if note:
        text(sl, x, y0 + rh * n + 8, w, nh, T(note, size - 3, ONDARK if dark else SUB), spacing=1.3)


def feature(sl, x, y, w, h, icon_key, title, body, dark=False, size=19):
    """图标卡：浅底圆图标 + 标题 + 正文段落，适合文字多、无数据的并列事项。整组在格内垂直居中。"""
    import icons as _I
    d = 56
    ts = 21
    while est_lines(title, w, ts) > 1 and ts > 15:
        ts -= 0.5
    head = d + 16 + ts * 1.6
    fs = size
    while est_height([(body, fs)], w, 1.55) > (h - head) and fs > 12:
        fs -= 0.5
    bh = est_height([(body, fs)], w, 1.55)
    y0 = y + max(0, (h - head - bh) * 0.4)
    circ = sl.shapes.add_shape(MSO_SHAPE.OVAL, E(x), E(y0), E(d), E(d))
    _paint(circ, "FFFFFF" if dark else mix(MAIN, "FFFFFF", .86))
    _I.icon(sl, icon_key, x + 14, y0 + 14, 28, MAIN)
    text(sl, x, y0 + d + 16, w, ts * 1.5, T(title, ts, "FFFFFF" if dark else INK, True), anchor="t")
    text(sl, x, y0 + head, w, bh + 10, T(body, fs, ONDARK if dark else SUB), spacing=1.55)


def para(sl, x, y, w, h, body, dark=False, size=22, lead=None):
    """正文段落（可带一句加粗导语），自动缩字号到放得下。"""
    fs = size
    paras = ([(lead, fs)] if lead else []) + [(body, fs)]
    while est_height(paras, w, LS_BODY) + (SP_INNER if lead else 0) > h and fs > 12:
        fs -= 0.5
        paras = ([(lead, fs)] if lead else []) + [(body, fs)]
    runs = []
    if lead:
        runs.append([(lead, fs + 1, "FFFFFF" if dark else INK, True)])
    runs.append([(body, fs, ONDARK if dark else SUB, False)])
    tb = text(sl, x, y, w, h, runs, anchor="m", spacing=LS_BODY)
    if lead:  # 导语与正文之间是"段内间距"，比行距大、比条目间距小
        tb.text_frame.paragraphs[0].line_spacing = LS_LEAD
        tb.text_frame.paragraphs[1].space_before = Pt(SP_INNER * 0.75)


# ================= v2 新增：卡内小逻辑 =================
def _rshape(sl, x, y, w, h, fill, radius=None):
    """圆角或直角色块。radius=None 时跟随整套 CORNER 设置。"""
    rnd = (CORNER == "round") if radius is None else radius > 0
    if rnd:
        s = sl.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, E(x), E(y), E(w), E(h))
        r = 10 if radius is None else radius
        s.adjustments[0] = min(0.5, r / max(1, min(w, h)))
    else:
        s = sl.shapes.add_shape(MSO_SHAPE.RECTANGLE, E(x), E(y), E(w), E(h))
    return _paint(s, fill)


def tags(sl, x, y, w, h, items, lead=None, dark=False, size=15):
    """标签胶囊：一排圆角小标签，放不下自动换行；可带一行导语。"""
    cy = y
    if lead:
        ls = 17
        lh = est_height([(lead, ls)], w, 1.4)
        text(sl, x, cy, w, lh, T(lead, ls, "FFFFFF" if dark else INK, True), spacing=1.4)
        cy += lh + 12
    ph = size * 2.1
    gap = 10
    rows, cur, cw = [], [], 0
    for t in items:
        tw = len(t) * size + size * 1.8
        if cur and cw + tw > w:
            rows.append(cur); cur, cw = [], 0
        cur.append((t, tw)); cw += tw + gap
    if cur:
        rows.append(cur)
    tot = len(rows) * ph + (len(rows) - 1) * gap
    cy += max(0, (y + h - cy - tot) / 2) if not lead else 0
    fill = "FFFFFF" if dark else mix(MAIN, "FFFFFF", .86)
    for row in rows:
        cx = x
        for t, tw in row:
            _rshape(sl, cx, cy, tw, ph, fill, radius=ph / 2)
            text(sl, cx, cy, tw, ph, T(t, size, MAIN, True), align="c", anchor="m")
            cx += tw + gap
        cy += ph + gap


def tip(sl, x, y, w, h, body, kind="warn", dark=False, size=15):
    """提示条：警示或信息图标 + 一行强调色文字，放在格子底部收尾。"""
    import icons as _I
    col = ("FFFFFF" if dark else (ACC if kind == "warn" else MAIN))
    ic = "caution" if kind == "warn" else "info"
    fs = size
    while est_lines(body, w - 30, fs) > 1 and fs > 12:
        fs -= 0.5
    lh = max(24, est_height([(body, fs)], w - 30, 1.3))
    yy = y + h - lh
    _I.icon(sl, ic, x, yy + (min(lh, 24) - 20) / 2, 20, col)
    text(sl, x + 30, yy, w - 30, lh, T(body, fs, col, True), anchor="m", spacing=1.3)


def icon_steps(sl, x, y, w, h, items, dark=False, size=17):
    """图标方块 + 标题 + 短说明，横向箭头相连。"""
    import icons as _I
    n = len(items)
    arw = 26
    cw = (w - arw * (n - 1)) / n
    tile = min(56, cw * 0.4, h * 0.36)
    ds = 15
    need = max(est_height([(it[1], ds)], cw, 1.35) for it in items) if any(len(it) > 1 and it[1] for it in items) else 0
    while ds > 12 and tile + 12 + size * 1.5 + 8 + need > h:
        ds -= 0.5
        need = max(est_height([(it[1], ds)], cw, 1.35) for it in items)
    tot = tile + 12 + size * 1.5 + 8 + need
    y0 = y + max(0, (h - tot) / 2)
    for i, it in enumerate(items):
        cx = x + i * (cw + arw)
        tx = cx + (cw - tile) / 2
        _rshape(sl, tx, y0, tile, tile, ("FFFFFF", mix(MAIN, "FFFFFF", .8), 315) if dark else (MAIN, mix(MAIN, "FFFFFF", .3), 315))
        _I.icon(sl, it[2] if len(it) > 2 else None, tx + tile * 0.24, y0 + tile * 0.24, tile * 0.52, MAIN if dark else "FFFFFF")
        text(sl, cx, y0 + tile + 12, cw, size * 1.5, T(it[0], size, "FFFFFF" if dark else INK, True), align="c", anchor="m")
        if len(it) > 1 and it[1]:
            text(sl, cx + 2, y0 + tile + 12 + size * 1.5 + 8, cw - 4, need + 6,
                 T(it[1], ds, ONDARK if dark else SUB), align="c", spacing=1.35)
        if i < n - 1:
            _shape(sl, MSO_SHAPE.CHEVRON, cx + cw + 8, y0 + tile / 2 - 7, arw - 16, 14,
                   "FFFFFF" if dark else mix(MAIN, "FFFFFF", .5))


def num_pair(sl, x, y, w, h, items, note=None, dark=False):
    """两到三个大数字，中间竖线隔开，各配一句说明；下面细线后接一段注释。"""
    n = len(items)
    note_h = 0
    if note:
        note_h = est_height([(note, 14)], w, LS_BODY) + 14
    area_h = h - note_h
    cw = w / n
    ns = 48
    while ns > 30 and max(len(it[0]) for it in items) * ns * 0.58 + 40 > cw - 24:
        ns -= 2
    block = ns * 1.25 + 10 + 44
    y0 = y + max(0, (area_h - block) / 2) if not note else y + max(0, (h - block - note_h) / 2)
    area_h = (y0 - y) + block
    for i, (val, unit, lab) in enumerate(items):
        cx = x + i * cw + (24 if i else 0)
        text(sl, cx, y0, cw - 24, ns * 1.25,
             [[(val, ns, "FFFFFF" if dark else MAIN, True), (" " + unit, ns * 0.4, "FFFFFF" if dark else MAIN, True)]],
             anchor="b", spacing=1.0)
        text(sl, cx, y0 + ns * 1.25 + 10, cw - 30, 44, T(lab, 14, ONDARK if dark else SUB), spacing=1.35)
        if i:
            vline(sl, x + i * cw, y0 + 6, block - 6, mix(MAIN, "FFFFFF", .35) if dark else LINE, 1)
    if note:  # 注释直接接在数字下方，不再用细线隔出一个小格子
        text(sl, x, y + area_h + 10, w, note_h - 10, T(note, 14, ONDARK if dark else MUTED), spacing=LS_BODY)


# ================= v2 新增：封面、目录、转场、结束页、页码 =================
def page_number(sl, n):
    text(sl, 620, 698, 40, 18, T(str(n), 12, MUTED), align="c", anchor="m")


def cover(prs, title, subtitle="", org="", date="", photo=None):
    import photos as _P
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    rect(sl, 0, 0, 1280, 720, MAIN)
    if photo:
        _P.place(sl, photo, 0, 0, 1280, 720, blurred=True)  # 全图蒙版背景模糊，字更清楚
        _P.overlay(sl, 0, 0, 1280, 720, MAIN, 0.90, 0.12, 0)
    else:
        rect(sl, 0, 0, 1280, 720, DARK_G)
    lines, size = break_title(title, 980, 54, 34)
    _title_bar(sl, 370, size, len(lines))
    text(sl, 108, 200, 1000, 170, [[(l, size, "FFFFFF", True)] for l in lines], anchor="b", spacing=1.15)
    if subtitle:
        text(sl, 108, 384, 1000, 40, T(subtitle, 22, mix(MAIN, "FFFFFF", .85)), anchor="t")
    meta = "｜".join(t for t in (org, date) if t)
    if meta:
        text(sl, 108, 600, 800, 30, T(meta, 16, mix(MAIN, "FFFFFF", .85)), anchor="m")
    return sl


def agenda(prs, sections):
    """目录：sections = [(章名, 说明), …]。左侧主色块写"目录"，右侧编号列表。"""
    sl = add_slide(prs)
    rect(sl, 0, 0, 400, 720, DARK_G)
    text(sl, 64, 260, 300, 80, T("目录", 48, "FFFFFF", True), anchor="m")
    text(sl, 66, 340, 300, 30, T("CONTENTS", 16, mix(MAIN, "FFFFFF", .7), True), anchor="m")
    n = len(sections)
    rh = min(96, 560 / max(n, 1))
    y0 = 360 - rh * n / 2
    for i, sec in enumerate(sections):
        if isinstance(sec, dict):
            name, desc = sec.get("name", ""), sec.get("desc", "")
        elif isinstance(sec, (list, tuple)):
            name, desc = sec[0], (sec[1] if len(sec) > 1 else "")
        else:
            name, desc = sec, ""
        yy = y0 + i * rh
        text(sl, 470, yy, 90, rh, T("%02d" % (i + 1), 40, MAIN, True), anchor="m")
        text(sl, 580, yy + (rh / 2 - 34 if desc else 0), 620, 40 if desc else rh, T(name, 26, MAIN, True), anchor="m" if not desc else "b")
        if desc:
            text(sl, 580, yy + rh / 2 + 8, 620, 30, T(desc, 15, SUB), anchor="t")
        if i < n - 1:
            hline(sl, 470, yy + rh, 730, LINE, 0.75)
    return sl


def transition(prs, idx, name, desc="", photo=None):
    """章节转场：左侧主色块写章节号与章名；有图时右侧放半幅图。"""
    import photos as _P
    sl = add_slide(prs)
    split = 560 if photo else 1280
    if photo:
        sat, duo = _P.tone_for(photo)
        _P.place(sl, photo, 560, 0, 720, 720, tone_main=MAIN, sat=sat, duo=duo * 0.8)  # 半幅图只有接缝处有蒙版，先调色避免与主色冲突
        _P.overlay(sl, 560, 0, 720, 720, MAIN, 0.55, 0.0, 0)
    rect(sl, 0, 0, split, 720, DARK_G)
    text(sl, 80, 180, 440, 150, T("%02d" % idx, 120, "FFFFFF", True), anchor="b")
    lines, size = break_title(name, 420 if photo else 900, 40, 28)
    text(sl, 84, 350, 440 if photo else 900, 100, [[(l, size, "FFFFFF", True)] for l in lines], anchor="t", spacing=1.2)
    if desc:
        text(sl, 84, 470, 420 if photo else 900, 80, T(desc, 18, ONDARK), anchor="t", spacing=1.4)
    return sl


def _title_bar(sl, bottom, size, lines=1):
    """封面、结束页标题左侧的强调竖条，高度与标题文字等高（不高出字）。bottom 为标题文本框底边。"""
    h = size * (1.15 * (lines - 1) + BAR_K)
    rect(sl, 80, bottom - size * BAR_D - h, 6, h, (MAIN, mix(MAIN, "FFFFFF", .3), 90))


BAR_K, BAR_D = 0.95, 0.16  # 竖条高度系数（相对字号）、文本框底边到字底的距离系数


def closing(prs, words="感谢聆听", sub="", photo=None):
    """结束页：与封面同一张图、同样的主色渐变蒙版和左侧强调竖条；大字"感谢聆听"，下面一行写汇报人。"""
    import photos as _P
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    rect(sl, 0, 0, 1280, 720, MAIN)
    if photo:
        _P.place(sl, photo, 0, 0, 1280, 720, blurred=True)  # 全图蒙版背景模糊，字更清楚
        _P.overlay(sl, 0, 0, 1280, 720, MAIN, 0.90, 0.12, 0)
    else:
        rect(sl, 0, 0, 1280, 720, DARK_G)
    _title_bar(sl, 370, 64)
    text(sl, 108, 200, 1000, 170, T(words, 64, "FFFFFF", True), anchor="b", spacing=1.15)
    if sub:
        text(sl, 108, 384, 1000, 40, T(sub, 22, mix(MAIN, "FFFFFF", .85)), anchor="t")
    return sl

"""稳重风皮肤与页型：把“功能区切法”落成可调用的函数。

统一字号（1280×720 画布，px）：小标题 20，正文 16，标签 13–14，关键数字 28–40。
统一区域：页眉 0–70；主体 84–596；结论带 610–680；来源与页码 698。
每个函数在文字可能放不下时记入 WARN，组装完统一打印。
"""
import math, os, re
from pptx.util import Emu, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION, XL_TICK_LABEL_POSITION
from compose import alt, compact_text, is_solid, on_accent, metric_color, page_number, box, text, line, est_lines, new_deck, table, banner, PALETTE, W, H, M, PX, _rgb, _mix, FONT_CN
from deco import grad_tab, grad_badge, soft_card, backdrop, half_stadium, fade_bar, grad_fill, light, dark, shadow

CAP, BODY, SMALL, NUM = 20, 16, 16, 30
TOP, BOT = 86, 596
BAND_Y, BAND_H = 610, 70
WARN = []
BRAND = {"name": "", "logo": None}

def P(k): return PALETTE[k]

def need_h(s, w, size, spacing=1.4):
    size = max(size, 16)
    return est_lines(s, w * 0.94, size) * size * spacing * 1.05   # 留 6% 宽度余量，防止不同字体换行更早

def card_text_height(value, width, size, spacing):
    from compose import compact_text
    from render import preview_font
    size = max(16, size)
    font = preview_font(round(size))
    count = 0
    for paragraph in re.split(r"\r\n|[\n\v\r]", compact_text(value)):
        count += 1
        line = ""
        for token in re.findall(r"(?:[0-9]+(?:[,.][0-9]+)*[%％]?|[A-Za-z]+|.)[，。；：、）”]*", paragraph):
            if line and font.getlength(line + token) > width * .94:
                count += 1
                line = ""
            line += token
    return count * size * spacing * 1.05


def roomier_font(body, width, height, preferred, minimum, spacing):
    return next((size for size in range(preferred, minimum, -1)
                 if need_h(body, width, size, spacing) <= height), minimum)


def fit_text(sl, x, y, w, h, s, size, color, bold=False, where="", **kw):
    if need_h(s, w, size, kw.get("spacing", 1.4)) > h + 2:
        WARN.append(f"{where}: 文字可能放不下（{s[:14]}…）")
    return text(sl, x, y, w, h, s, size, color, bold, **kw)

# ---------------- 页面骨架 ----------------
def page(prs, section, title, num=None, source=None):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    page_header(sl, section, title, num, source)
    return sl


def page_header(sl, section, title, num=None, source=None):
    tw = W - 2 * M - 220
    if est_lines(section + "丨" + title, tw, 26) > 1: WARN.append(f"p{num} 标题过长：{title}")
    text(sl, M, 20, tw, 40, [[(section + "  丨  ", 26, P("navy"), True), (title, 26, P("ink"), True)]], anchor="m", heading=True, role="title")
    if BRAND["logo"]:
        pic = sl.shapes.add_picture(BRAND["logo"], Emu(int((W - M - 150) * PX)), Emu(int(BRAND.get("logo_y", 22) * PX)), height=Emu(int(BRAND.get("logo_h", 36) * PX)))
        pic.left = Emu(int((W - M) * PX) - pic.width)
    else:
        name_size = next((z for z in (14, 13, 12) if est_lines(BRAND["name"], 214, z) == 1), 0)
        if not name_size:
            WARN.append(f"p{num} 页眉右上角名称“{BRAND['name']}”过长，放不下一行；start(name=…) 请用 12 字以内的简称，全称写在封面")
            name_size = 12
        text(sl, W - M - 220, 22, 220, 40, BRAND["name"], name_size, P("navy"), True, align="r", anchor="m", spacing=1.0)
    line(sl, M, 68, W - M, 68, P("navy"), 2)
    if source: text(sl, M, 696, 1000, 16, source, 11, P("muted"), role="source")
    if num: page_number(sl, num, P("navy"))
    return sl

def band(sl, message, pills=None, y=BAND_Y, h=BAND_H):
    box(sl, M, y, W - 2 * M, h, P("dark"))
    pw = 0
    if pills:
        pw = 170 * len(pills); px = W - M - pw
        box(sl, px, y, pw, h, P("peach"))
        for i, (name, val) in enumerate(pills):
            text(sl, px + i * 170 + 16, y + 6, 150, 20, name, SMALL, on_accent(), anchor="m")
            text(sl, px + i * 170 + 16, y + 26, 150, h - 30, val, 26, on_accent(), True, anchor="m")
    box(sl, M, y, 6, h, P("peach"))
    fit_text(sl, M + 26, y, W - 2 * M - pw - 50, h, message, 19, "FFFFFF", True, where="结论带", anchor="m", role="conclusion")

# ---------------- 组件 ----------------
# 标点不落行首：后置标点随前一个字，前置括号与引号随后一个字。
_MODULE_TOKEN = re.compile(r"[（(“‘《「『【]*(?:[A-Za-z0-9]+(?:[.,/-][A-Za-z0-9]+)*[%％]?|.)[，。；：、！？）)”’》」』】,.;:!?]*")


def _module_lines(value, width, size, bold=False):
    """在编辑稿明确分行，给实际字体和粗体留出宽度余量。"""
    from render import preview_font
    font = preview_font(size)
    limit = max(1, width * (.91 if bold else .94))
    result = []
    for para in re.split(r"\r\n|[\n\v\r]", compact_text(value)):
        line = ""
        for token in _MODULE_TOKEN.findall(para):
            parts = [token] if font.getlength(token) <= limit else list(token)
            for part in parts:
                if line and font.getlength(line + part) > limit:
                    result.append(line)
                    line = ""
                line += part
        result.append(line)
    return result


def _module_copy(sl, x, y, w, lines, size, color, bold, leading):
    """用相同的显式行、固定行距计量和绘制，防止 Impress 放大百分比行距。"""
    height = len(lines) * leading
    shape = text(sl, x, y, w, height, lines, size, color, bold)
    for para in shape.text_frame.paragraphs:
        para.line_spacing = Pt(leading * .75)
        para.space_after = para.space_before = Pt(0)
    return height


def module(sl, x, y, w, h, title, body=None, points=(), metrics=(), style="dark", color=None, where="", justify=True):
    """模块：标题帽 + 一句判断 + 要点 + 底部指标对。style: dark（深色标题帽）| grad（渐变标题块）
    内容先量高度；剩余空间均分到各段之间（justify），避免内容挤在顶部、下面一大片空白。"""
    color = color or P("navy")
    top = y + (56 if style == "dark" else 60)
    pad = 18; iw = w - 2 * pad
    mh = 84 if metrics else 0
    blocks = []   # (高度, 绘制函数)
    if body:
        lines = _module_lines(body, iw, 17)
        bh = len(lines) * 26
        blocks.append((bh, lambda yy: _module_copy(sl, x + pad, yy, iw, lines, 17, P("ink"), False, 26)))
    for lead, detail in points:
        lead_lines = _module_lines(lead, iw - 18, BODY, bold=bool(detail))
        detail_lines = _module_lines(detail, iw - 18, 16) if detail else []
        lh = len(lead_lines) * 24
        dh = len(detail_lines) * 24
        def draw(yy, head=lead_lines, desc=detail_lines, lh=lh):
            box(sl, x + pad, yy + 8, 7, 7, P("peach"))
            _module_copy(sl, x + pad + 18, yy, iw - 18, head, BODY, P("navy"), bool(desc), 24)
            if desc:
                _module_copy(sl, x + pad + 18, yy + lh + 4, iw - 18, desc, 16, P("muted"), False, 24)
        blocks.append((lh + (dh + 4 if detail else 0), draw))
    used = sum(b[0] for b in blocks)
    h = min(h, top - y + used + 8 * (len(blocks) + 1) + mh + 10)
    bottom = y + h - mh - 10
    if style == "dark":
        cap = box(sl, x, y, w, 44)
        grad_fill(cap, [(0, P("dark"), None), (100, color, None)], 0, large_block=True)
        box(sl, x, y + 44, w, h - 44, P("tint"))
        text(sl, x + 16, y, w - 32, 44, title, CAP, "FFFFFF", True, anchor="m", heading=True)
    else:
        soft_card(sl, x, y + 20, w, h - 20, color)
        grad_tab(sl, x, y, w, 46, title, color, CAP)
    free = bottom - top - used
    if free < 0:
        WARN.append(f"{where}{title}: 模块内容溢出 {math.ceil(-free)}px（栏宽 {w:.0f}px）；请加宽或增高本区")
    gap = min(10, max(0, free / max(1, len(blocks))))
    if justify and blocks:
        gap = max(0, min(34, free / (len(blocks) + 1)))
        if free / max(1, bottom - top) > 0.45: WARN.append(f"{where}{title}: 内容偏少，留白 {int(free)}px")
    yy = top + (gap if justify else 0)
    for bh, draw in blocks:
        draw(yy); yy += bh + gap
    if metrics:
        my = y + h - 76
        line(sl, x + pad, my - 4, x + w - pad, my - 4, P("line"))
        mw = iw / len(metrics)
        for j, (name, val) in enumerate(metrics):
            text(sl, x + pad + j * mw, my + 2, mw - 8, 20, name, SMALL, P("muted"), anchor="m")
            text(sl, x + pad + j * mw, my + 24, mw - 8, 44, val, 28, color, True, anchor="m")
    return used / max(1, bottom - top)

def kpi_tiles(sl, x, y, w, h, items, gap=12, style="tint"):
    """指标项为 (标签, 数值, 注释[, 改善/承压/中性])；未注明含义用中性色。
    超过 6 项时自动排成两行、平分 h（h 至少给 2×100+gap），返回底边 y。"""
    if len(items) > 6:
        half = (len(items) + 1) // 2; rh = (h - gap) / 2
        if rh < 90:
            WARN.append(f"指标卡：{len(items)} 项分两行，每行只有 {rh:.0f}px，数值放不下；h 至少给 {2 * 90 + gap}px")
        kpi_tiles(sl, x, y, w, rh, items[:half], gap, style)
        return kpi_tiles(sl, x, y + rh + gap, w, rh, items[half:], gap, style)
    n = len(items); cw = (w - gap * (n - 1)) / n
    for i, item in enumerate(items):
        lab, val, note = item[:3]
        col = metric_color(item[3] if len(item) > 3 else "中性")
        cx = x + i * (cw + gap)
        box(sl, cx, y, cw, h, P("tint")); box(sl, cx, y, cw, 4, col)
        text(sl, cx + 14, y + 12, cw - 28, 20, lab, 14, P("muted"))
        text(sl, cx + 14, y + 34, cw - 28, 44, val, 32 if h >= 100 else 28, col, True, anchor="m")
        if note:
            text(sl, cx + 14, y + h - 26, cw - 28, 20, note, SMALL, dark(col, 0.1) if col == P("peach") else col, note[0] in "▲▼")
    return y + h

def panel(sl, x, y, w, h, title, insight=None):
    """带渐隐标题条的白底面板，返回内部可用区 (x, y, w, h)"""
    box(sl, x, y, w, h, "FFFFFF", P("line"))
    fade_bar(sl, x, y, w, 34, title, P("navy"), 16, a0=22)
    iy = y + 40; ih = h - 46
    if insight:
        ih -= 40
        box(sl, x + 12, y + h - 42, w - 24, 34, P("tint"))
        fit_text(sl, x + 22, y + h - 42, w - 44, 34, insight, 14, P("ink"), where=f"面板[{title}]", anchor="m")
    return x + 10, iy, w - 20, ih

def grad_shape_fill(fmt_fill, color, angle=90):
    """图表系列或数据点的渐变填充：同色由浅到深，柱子、条形、扇区统一用这一种。"""
    fmt_fill.gradient()
    fmt_fill.gradient_angle = angle
    stops = fmt_fill.gradient_stops
    stops[0].position, stops[0].color.rgb = 0.0, _rgb(light(color, 0.38))
    stops[-1].position, stops[-1].color.rgb = 1.0, _rgb(dark(color, 0.08))


def chart(sl, x, y, w, h, kind, cats, series, fmt='0.0', colors=None, legend=None, labels=True, font=16, gap=60, overlap=None, vmin=None, vmax=None):
    """kind: col | bar | line | stack | stackbar"""
    kinds = {"col": XL_CHART_TYPE.COLUMN_CLUSTERED, "bar": XL_CHART_TYPE.BAR_CLUSTERED,
             "line": XL_CHART_TYPE.LINE_MARKERS, "stack": XL_CHART_TYPE.COLUMN_STACKED, "stackbar": XL_CHART_TYPE.BAR_STACKED}
    cd = CategoryChartData(); cd.categories = [compact_text(cat) for cat in cats]
    for name, vals in series: cd.add_series(compact_text(name), vals)
    gf = sl.shapes.add_chart(kinds[kind], Emu(int(x * PX)), Emu(int(y * PX)), Emu(int(w * PX)), Emu(int(h * PX)), cd)
    ch = gf.chart; ch.has_title = False
    ch.font.size = Pt(max(12, font * .75)); ch.font.name = FONT_CN
    show_leg = legend if legend is not None else len(series) > 1
    ch.has_legend = show_leg
    if show_leg:
        ch.legend.position = XL_LEGEND_POSITION.TOP; ch.legend.include_in_layout = False
    pal = colors or PALETTE.get("chart_list") or [P("navy"), P("peach"), "8E8385", light(P("navy"), 0.55)]
    pl = ch.plots[0]
    if kind != "line":
        pl.gap_width = gap
        if overlap is not None: pl.overlap = overlap
    if labels:
        pl.has_data_labels = True
        pl.data_labels.number_format = fmt; pl.data_labels.number_format_is_linked = False
        pl.data_labels.font.size = Pt(max(12, font * .75))
        if kind in ("col", "bar"): pl.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
        elif kind == "line": pl.data_labels.position = XL_LABEL_POSITION.ABOVE
        else: pl.data_labels.position = XL_LABEL_POSITION.CENTER
    for i, s in enumerate(pl.series):
        c = pal[i % len(pal)]
        if kind == "line":
            s.format.line.color.rgb = _rgb(c); s.format.line.width = Pt(2.25); s.smooth = False
            s.marker.format.fill.solid(); s.marker.format.fill.fore_color.rgb = _rgb(c); s.marker.size = 7
            s.marker.format.line.color.rgb = _rgb(c)
        else:
            grad_shape_fill(s.format.fill, c, 90 if kind in ("col", "stack") else 0)
            s.invert_if_negative = False
        if kind in ("stack", "stackbar") and labels:
            s.data_labels.font.color.rgb = _rgb("FFFFFF" if i == 0 else P("ink"))
    va = ch.value_axis
    va.has_major_gridlines = kind == "line"
    if va.has_major_gridlines: va.major_gridlines.format.line.color.rgb = _rgb(P("line"))
    va.visible = kind == "line"
    if kind != "line":
        values = [v for _, vals in series for v in vals if v is not None]
        if kind in ("stack", "stackbar"):
            columns = list(zip(*(vals for _, vals in series)))
            values = [sum(v for v in column if v is not None and v * sign > 0)
                      for column in columns for sign in (1, -1)]
        lo, hi = min([0, *values]), max([0, *values])
        # Reserve room for outside labels at both ends of a horizontal plot.
        pad = (hi - lo or 1) * (.22 if kind == "bar" else .13)
        if kind == "bar" and labels and lo < 0:
            # LOW category ticks sit at the left edge. Outside-end data labels
            # need a fixed text-width gutter even when the plot is very narrow.
            from render import preview_font
            face = preview_font(max(16, font))
            label_w = max(face.getlength(f"{v:,.2f}") for v in values) + 14
            category_w = min(w*.33, max(face.getlength(str(cat)) for cat in cats))
            remaining = max(w*.18, w-category_w-36-2*label_w)
            pad = max(pad, (hi-lo)*label_w/remaining)
        if vmin is None: vmin = lo - pad if lo < 0 else 0
        if vmax is None: vmax = hi + pad if hi > 0 or lo == 0 else 0
        if labels:
            body = pl.data_labels._element.xpath('./c:txPr/a:bodyPr')
            if body: body[0].set('wrap', 'none')
    if vmin is not None: va.minimum_scale = vmin
    if vmax is not None: va.maximum_scale = vmax
    ca = ch.category_axis; ca.tick_labels.font.size = Pt(max(12, font * .75))
    if kind in ("bar", "stackbar"):
        ca.reverse_order = True   # 横向条形图按传入顺序自上而下排列，与表格、文字的阅读顺序一致
    ca.format.line.color.rgb = _rgb(P("line"))
    if kind != "line": ca.tick_label_position = XL_TICK_LABEL_POSITION.LOW
    return ch

def row_matrix(sl, x, y, w, h, heads, rows, label_w=150, emph_rows=(), size=16):
    """行式矩阵：左侧深色行标签 + 表头 + 单元格（单元格可为 str 或 (大字, 小字)）。
    行高按文字收紧，不再均分 h；返回实际底边 y，空出的高度留给结论带或说明。"""
    nh = 40; cw = (w - label_w) / len(heads)
    def _cell_need(c):
        if isinstance(c, tuple):
            return 22 * 1.3 + need_h(c[1], cw - 20, SMALL, 1.3) + 14
        return need_h(c, cw - 24, size, 1.35) + 14
    natural = max(max([need_h(lab, label_w - 28, 16, 1.3) + 14] + [_cell_need(c) for c in cells]) for lab, cells in rows)
    natural = max(natural, 52) + 4
    rh = min((h - nh) / len(rows), natural)
    used = nh + 4 + rh * len(rows)
    if h - used >= 40:
        WARN.append(f"行式矩阵：按文字收高后下方空出约 {h - used:.0f}px，请在矩阵下方（y={y + used + 8:.0f} 起）放结论带、说明或小卡片，不要留白")
    def header_cell(left, width):
        cap = box(sl, left, y, width - 4, nh)
        start = _mix(P("dark"), P("navy"), (left - x) / w)
        end = _mix(P("dark"), P("navy"), (left + width - x) / w)
        grad_fill(cap, [(0, start, None), (100, end, None)], 0, large_block=True)
    header_cell(x, label_w)
    for j, hd in enumerate(heads):
        header_cell(x + label_w + j * cw, cw)
        text(sl, x + label_w + j * cw + 10, y, cw - 24, nh, hd, 15, "FFFFFF", True, anchor="m", align="c")
    for i, (lab, cells) in enumerate(rows):
        ry = y + nh + 4 + i * rh
        hl = i in emph_rows
        box(sl, x, ry, label_w - 4, rh - 4, P("dark") if not hl else P("peach"))
        text(sl, x + 12, ry, label_w - 28, rh - 4, lab, 16, "FFFFFF" if not hl else on_accent(), True, anchor="m")
        for j, c in enumerate(cells):
            cx = x + label_w + j * cw
            box(sl, cx, ry, cw - 4, rh - 4, P("peach_l") if hl else (P("tint") if i % 2 == 0 else "FFFFFF"))
            if isinstance(c, tuple):
                big, small = c
                # 大字、小字作为一组在单元格内垂直居中，上下各占自己的高度，不再按比例切分而互相压住
                big_h = 22 * 1.3; small_h = need_h(small, cw - 20, SMALL, 1.3)
                top = ry + max(2, (rh - 4 - big_h - 2 - small_h) / 2)
                text(sl, cx + 8, top, cw - 20, big_h, big, 22, P("navy") if not hl else dark(P("peach"), 0.25), True, align="c", anchor="m", spacing=1.0)
                text(sl, cx + 8, top + big_h + 2, cw - 20, small_h, small, SMALL, P("muted"), align="c", spacing=1.3)
            else:
                fit_text(sl, cx + 10, ry, cw - 24, rh - 4, c, size, P("ink"), where="矩阵单元格", align="c", anchor="m")
    return y + used

def card_grid(sl, x, y, w, h, items, cols, gap=16, color=None, start=1, style="grad"):
    """卡片网格：items [(标题, 说明, (数值, 标签) 或 None[, 状态])]。
    正文顶对齐，字号 17–20px；同一网格的卡片等高，高度按文字最多的一张收紧。
    不为填满区域而放大字号或拉高卡片：返回卡片实际占用的高度，下方空出的部分放同页其他内容。
    第 4 项写状态（高 / 中 / 低，或 按期 / 滞后 / 风险）时序号徽标按状态着色，并在标题右侧标出状态。"""
    if not items or cols < 1:
        raise ValueError("卡片网格需要内容和正列数")
    from render import preview_font
    cols = min(cols, len(items), max(1, int((w + gap) / (190 + gap))))
    rows = (len(items) + cols - 1) // cols
    cw = (w - gap * (cols - 1)) / cols; chh = (h - gap * (rows - 1)) / rows
    items = [tuple(it) + (None,) * (4 - len(it)) for it in items]
    has_status = any(it[3] for it in items)
    mh = 62 if any(it[2] for it in items) else 0
    head, pad = 64, 16
    room = chh - head - mh - pad
    def body_h(value, size):   # 两种估算取大值，与 fit_text 的溢出判断保持一致
        return max(card_text_height(value, cw - 40, size, 1.45), need_h(value, cw - 40, size, 1.45))
    body_size = next((z for z in (20, 19, 18, 17) if all(body_h(it[1], z) <= room for it in items)), 17)
    need = max(body_h(it[1], body_size) for it in items)
    card_h = min(chh, head + need + pad + mh)
    used = rows * card_h + gap * (rows - 1)
    spare = h - used
    if spare > max(110, h * 0.28):
        WARN.append(f"卡片网格: 区域未填满，卡片按文字收高后只占区域高度的 {used / h:.0%}，下方空出 {int(spare)}px。"
                    "把这一区的高度调小（S.rows / content_page 的 weight），空出的位置放同页其他内容（措施、关键数、依据），或与相邻页合并")
    elif spare > 48:
        WARN.append(f"卡片网格: 内容偏少，卡片下方空出 {int(spare)}px")
    for i, (t, d, m, status) in enumerate(items):
        cx = x + (i % cols) * (cw + gap); cy = y + (i // cols) * (card_h + gap)
        if status:
            from zones import status_color
            col = status_color(status, i)
        else:
            col = color or (P("navy") if has_status else alt(i))
        soft_card(sl, cx, cy, cw, card_h, P("navy"))
        grad_badge(sl, cx + 36, cy + 36, 44, str(start + i), col, 20)
        tag_w = 0
        if status and not re.fullmatch(r"#?[0-9a-fA-F]{6}", status.strip()):
            tag_w = preview_font(15).getlength(status) + 28
            tb = box(sl, cx + cw - 20 - tag_w, cy + 22, tag_w, 28, col, shape=MSO_SHAPE.ROUNDED_RECTANGLE, adj=[0.5])
            text(sl, cx + cw - 20 - tag_w, cy + 22, tag_w, 28, status, 15, "FFFFFF", True, align="c", anchor="m", spacing=1.0)
            tag_w += 8
        fit_text(sl, cx + 68, cy + 18, cw - 84 - tag_w, 36, t, CAP, P("navy"), True, where=f"卡片{t}", anchor="m", heading=True)
        dy = cy + head
        dh_avail = card_h - head - mh - 4
        fit_text(sl, cx + 20, dy, cw - 40, dh_avail, d, body_size, P("ink"), where=f"卡片{t}", spacing=1.45, anchor="t")
        if m:
            v, lab = m
            line(sl, cx + 20, cy + card_h - mh, cx + cw - 20, cy + card_h - mh, P("line"))
            # 数值与标签分两个文本框：间距固定 10px，底边按基线对齐（小字号上提 3px）
            vw = min(cw - 40, preview_font(28).getlength(compact_text(v)) * 1.08 + 6)
            text(sl, cx + 20, cy + card_h - mh + 8, vw, 44, v, 28, P("navy"), True, anchor="b", spacing=1.0)
            if cw - 40 - vw - 10 > 40:
                text(sl, cx + 20 + vw + 10, cy + card_h - mh + 8, cw - 40 - vw - 10, 41, lab, 16, P("muted"), anchor="b", spacing=1.0)
    return used

def compare(sl, left, right, y=TOP + 6, h=None, vs="VS", x0=M, width=None):
    """左右对照：left/right = {title, color, rows:[(标签, 数值, 说明)], note}"""
    h = h or (BOT - y)
    width = width or W - 2 * M
    gx = width / 15
    cw = (width - gx) / 2
    metric_w = cw * 0.35
    middle = x0 + width / 2
    for k, (blk, x) in enumerate(((left, x0), (right, x0 + width - cw))):
        col = blk.get("color", P("navy") if k == 0 else P("peach"))
        soft_card(sl, x, y + 22, cw, h - 22, col)
        grad_tab(sl, x, y, cw, 50, blk["title"], col, 22)
        n = len(blk["rows"]); top = y + 66; rh = (h - 66 - (40 if blk.get("note") else 6)) / n
        for i, (lab, val, desc) in enumerate(blk["rows"]):
            ry = top + i * rh
            if i: line(sl, x + 22, ry - 4, x + cw - 22, ry - 4, P("line"))
            text(sl, x + 22, ry + 4, metric_w, rh - 12, [[(val, 30, col if col == P("navy") else dark(col, 0.2), True)], [(lab, 14, P("muted"), False)]], anchor="m", spacing=1.15)
            fit_text(sl, x + metric_w + 30, ry + 4, cw - metric_w - 52, rh - 12, desc, 17, P("ink"), where=f"对照[{blk['title']}]", anchor="m", spacing=1.5)
            if need_h(desc, cw - metric_w - 52, 17, 1.5) < (rh - 12) * 0.35: WARN.append(f"对照[{blk['title']}]: 说明偏短（{desc[:10]}…）")
        if blk.get("note"):
            fit_text(sl, x + 22, y + h - 40, cw - 44, 32, blk["note"], 14, P("muted"), where="对照注", anchor="m")
    if vs:
        d = gx + 36
        if isinstance(vs, tuple):   # (标签, 数值)：单独出现的数字必须带标签
            lab, val = vs
            grad_badge(sl, middle, y + h / 2, d, "", P("dark"))
            text(sl, middle - d / 2, y + h / 2 - 30, d, 22, lab, 14, _mix(P("dark"), "FFFFFF", 0.8), True, align="c", anchor="m")
            text(sl, middle - d / 2, y + h / 2 - 8, d, 34, val, 22, "FFFFFF", True, align="c", anchor="m")
        else:
            grad_badge(sl, middle, y + h / 2, d, vs, P("dark"), 18)

def hero_left(sl, x, y, w, h, label, big, unit, note, rows=()):
    """左栏总述：主色竖栏 + 一个大数字 + 若干小指标"""
    box(sl, x, y, w, h, P("navy"))
    grad_fill(sl.shapes[-1], [(0, P("navy2"), None), (100, P("navy"), None)], 90)
    text(sl, x + 24, y + 22, w - 48, 24, label, 16, "FFFFFF", True)
    text(sl, x + 24, y + 54, w - 48, 70, [[(big, 56, "FFFFFF", True), (" " + unit, 20, "FFFFFF", True)]], anchor="m")
    fit_text(sl, x + 24, y + 128, w - 48, 48, note, 14, _mix(P("navy"), "FFFFFF", 0.78), where="左栏注", spacing=1.35)
    ry = y + 186
    if rows:
        rh = (h - 186 - 12) / len(rows)
        for i, (lab, val) in enumerate(rows):
            yy = ry + i * rh
            line(sl, x + 24, yy, x + w - 24, yy, _mix(P("navy"), "FFFFFF", 0.35))
            text(sl, x + 24, yy + 6, w - 48, 24, lab, SMALL, _mix(P("navy"), "FFFFFF", 0.75), spacing=1.0, anchor="m")
            text(sl, x + 24, yy + 32, w - 48, rh - 36, val, 26, "FFFFFF", True, anchor="m", spacing=1.0)
        if rh < 72:
            WARN.append(f"左栏小指标：{len(rows)} 行在 {h:.0f}px 高的左栏里每行只有 {rh:.0f}px，数值放不下；减行或加高")

def roof(sl, roof_title, roof_text, pillars, base, y=TOP + 4):
    """屋顶-支柱-基座"""
    # 屋顶：三角 + 横条
    tri = box(sl, M, y, W - 2 * M, 50, shape=MSO_SHAPE.ISOSCELES_TRIANGLE)
    grad_fill(tri, [(0, P("navy2"), None), (100, P("navy"), None)], 90)
    box(sl, M, y + 50, W - 2 * M, 46, P("navy"))
    text(sl, M, y + 16, W - 2 * M, 34, roof_title, 18, "FFFFFF", True, align="c", anchor="m")
    fit_text(sl, M + 30, y + 50, W - 2 * M - 60, 46, roof_text, 18, "FFFFFF", True, where="屋顶", align="c", anchor="m")
    n = len(pillars); gap = 14; pw = (W - 2 * M - gap * (n - 1)) / n
    py = y + 108; ph = BOT - 56 - py
    for i, pil in enumerate(pillars):
        t, pts = pil[0], pil[1]; ms = pil[2] if len(pil) > 2 else ()
        px = M + i * (pw + gap)
        module(sl, px, py, pw, ph, t, points=[p if isinstance(p, tuple) else (p, None) for p in pts], metrics=ms, style="grad", color=alt(i), where="支柱")
    box(sl, M, BOT - 44, W - 2 * M, 44, P("tint")); box(sl, M, BOT - 44, W - 2 * M, 3, P("navy"))
    fit_text(sl, M + 20, BOT - 44, W - 2 * M - 40, 44, base, BODY, P("navy"), True, where="基座", align="c", anchor="m")

# ---------------- 封面、目录、过渡 ----------------
def cover(prs, title, subtitle, org, date):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    box(sl, 0, 0, W, H, P("navy"))
    grad_fill(sl.shapes[-1], [(0, P("navy2"), None), (100, P("dark"), None)], 45, large_block=True)
    if BRAND.get("logo_card"):   # 彩色 LOGO 放在白色圆角卡片上（多色 LOGO 不适合反白）
        from PIL import Image
        iw, ih = Image.open(BRAND["logo_card"]).size
        lh = 72 if iw / ih > 2.2 else 92; lw = min(420, lh * iw / ih); lh = lw * ih / iw   # 横版 LOGO 压低高度、限制宽度
        cw_, ch_ = lw + 40, lh + 28
        box(sl, M * 2, 60, cw_, ch_, "FFFFFF", shape=MSO_SHAPE.ROUNDED_RECTANGLE, adj=[0.12])
        sl.shapes.add_picture(BRAND["logo_card"], Emu(int((M * 2 + 20) * PX)), Emu(int(74 * PX)), width=Emu(int(lw * PX)), height=Emu(int(lh * PX)))
    elif BRAND.get("logo_white"):
        from PIL import Image
        iw, ih = Image.open(BRAND["logo_white"]).size
        lh = 60 if iw / ih > 2.2 else 96; lw = min(380, lh * iw / ih); lh = lw * ih / iw
        sl.shapes.add_picture(BRAND["logo_white"], Emu(int(M * 2 * PX)), Emu(int(70 * PX)), width=Emu(int(lw * PX)), height=Emu(int(lh * PX)))
    box(sl, M * 2, 250, 80, 6, P("peach"))
    text(sl, M * 2, 272, 900, 80, title, 46, "FFFFFF", True, anchor="m", heading=True)
    text(sl, M * 2, 356, 900, 40, subtitle, 22, _mix(P("navy"), "FFFFFF", 0.8), anchor="m")
    text(sl, M * 2, 600, 700, 26, org, 16, "FFFFFF", True)
    text(sl, M * 2, 630, 700, 24, date, 14, _mix(P("navy"), "FFFFFF", 0.7))
    return sl

def agenda(prs, items, num=2, message=None, title=None):
    """目录仅展示章节与说明；兼容旧条目的第三项和 message 参数。"""
    n = len(items)
    if not n:
        raise ValueError("目录至少需要一个章节")
    sl = page(prs, "目录", title or "汇报内容", num=num)
    sl._element.cSld.name = "niuppt-agenda"
    gap = 16
    width = (W - 2 * M - gap * (n - 1)) / n
    # 第三项可给本章要点列表（如各页标题），写在说明下方，避免卡片下半空白；卡片按内容收高并在主体区居中
    points = [list(item[2]) if len(item) > 2 and isinstance(item[2], list) else [] for item in items]   # 旧版第三项（指标元组）忽略
    desc_h = max(card_text_height(item[1], width - 44, BODY, 1.5) for item in items) + 8
    pts_h = max([sum(card_text_height(p, width - 64, BODY, 1.3) + 8 for p in ps) for ps in points] or [0])
    need = 36 + 80 + 22 + 50 + 24 + desc_h + (24 + pts_h if pts_h else 0) + 28
    avail = 676 - TOP - 36      # 目录没有结论带，卡片一直延伸到页脚上方
    # 卡片占满主体区，内容在卡片内垂直居中：页面不再上下留白
    card_h = avail
    card_y = TOP + 36
    top = card_y + max(0, (avail - need) / 2)
    for i, item in enumerate(items):
        heading, description = item[:2]
        x = M + i * (width + gap)
        color = alt(i)
        soft_card(sl, x, card_y, width, card_h, color)
        text(sl, x + 22, top + 32, width - 44, 80, f"{i + 1:02d}", 60, light(color, 0.45), True)
        grad_tab(sl, x + 22, top + 134, width - 44, 50, heading, color, 20)
        fit_text(sl, x + 22, top + 208, width - 44, desc_h, description, BODY, P("ink"), where="目录", spacing=1.5)
        yy = top + 208 + desc_h + 16
        if points[i]:
            line(sl, x + 22, yy - 8, x + width - 22, yy - 8, P("line"))
        for p in points[i]:
            ph = card_text_height(p, width - 64, BODY, 1.3)
            dot = box(sl, x + 26, yy + 8, 8, 8, color, shape=MSO_SHAPE.OVAL)
            text(sl, x + 42, yy, width - 64, ph + 4, p, BODY, P("muted"), spacing=1.3)
            yy += ph + 8
    return sl


def divider(prs, no, title, sub, style=None, note=""):
    """代码版章节转场（无分类库时使用）：居中大号章号、章名、章节摘要，底部短线与目录说明胶囊。"""
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    sl._element.cSld.name = "niuppt-divider"
    box(sl, 0, 0, W, H, "FFFFFF")
    cx, tw = W / 2, 980
    text(sl, cx - 300, 92, 600, 170, f"{no:02d}", 150, P("navy"), True, align="c", anchor="m", spacing=1.0)
    text(sl, cx - tw / 2, 262, tw, 70, title, 52, P("navy"), True, align="c", anchor="m", spacing=1.0, heading=True)
    if note:
        text(sl, cx - tw / 2, 352, tw, 80, note, 18, P("muted"), align="c", spacing=1.4)
    line(sl, cx - 18, 462, cx + 18, 462, P("navy"), 2.5)
    if sub:
        from skin_solid import card_text_height
        from render import preview_font
        pw = min(tw, preview_font(18).getlength(sub) + 48)
        box(sl, cx - pw / 2, 486, pw, 40, P("navy"), shape=MSO_SHAPE.ROUNDED_RECTANGLE, adj=[0.5])
        text(sl, cx - pw / 2, 486, pw, 40, sub, 18, "FFFFFF", True, align="c", anchor="m", spacing=1.0)
    return sl

"""区域图示：每个函数把一种内容关系画进任意矩形区 (x, y, w, h)。

与整页页型不同，这些函数不关心页面骨架：可以放在标准页的主体区、左侧色栏页的右侧区、
上下 / 左右分栏后的任一格里。同一种骨架配不同的区域图示，页面就不会千篇一律。

两套皮肤共用；SOFT=True 时用圆角、渐变、投影，False 时用直角、平涂、细线。
每个函数在文字放不下或内容过少时记入 WARN。
成稿调用时，要点写成 (中观点, 支撑点)：页标题下要有这两级，不能只留中观点标题。确认大纲停在中观点，不在这里检查。
"""
import math
from lxml import etree
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from compose import box, text, line, est_lines, PX, _mix, _rgb, alt
from deco import grad_fill, grad_line, grad_badge, soft_card, shadow, light, dark, _grad_el
from skin_solid import P, need_h, fit_text, WARN

SOFT = False      # start() 按皮肤设置

def N(): return P("navy")
def A(): return P("peach")
def _c(i): return alt(i)
def _ink(col): return dark(col, 0.12)          # 用作文字的主题色（略压暗，保证对比度）

def _check(where, need, avail, low=0.35):
    if need > avail + 4: WARN.append(f"{where}: 文字可能放不下（超出约 {int(need - avail)}px）")
    elif avail > 60 and need < avail * low: WARN.append(f"{where}: 内容偏少，留白较多")

def _panel(sl, x, y, w, h, color, fill=None):
    if SOFT:
        return soft_card(sl, x, y, w, h, color, fill=fill or "FFFFFF", radius=0.06)
    s = box(sl, x, y, w, h, fill or P("tint")); box(sl, x, y, w, 4, color); return s

def _fill_block(sl, x, y, w, h, color, radius=0.06):
    if SOFT:
        s = box(sl, x, y, w, h, shape=MSO_SHAPE.ROUNDED_RECTANGLE, adj=[radius])
        grad_fill(s, [(0, light(color, 0.2), None), (100, dark(color, 0.12), None)], 45); shadow(s, 12, 4, 14, color)
    else:
        s = box(sl, x, y, w, h)
        grad_fill(s, [(0, color, None), (100, dark(color, 0.18), None)], 45)
    return s

def _badge(sl, cx, cy, d, label, color, size=None):
    if SOFT: return grad_badge(sl, cx, cy, d, label, color, size)
    s = box(sl, cx - d / 2, cy - d / 2, d, d, color, shape=MSO_SHAPE.RECTANGLE)
    text(sl, cx - d / 2, cy - d / 2, d, d, label, size or d * 0.42, "FFFFFF", True, align="c", anchor="m")
    return s

def _arrow(sl, x1, y1, x2, y2, color, w=2, dash=False, head=True):
    c = line(sl, x1, y1, x2, y2, color, w)
    ln = c._element.find(qn("p:spPr")).find(qn("a:ln"))
    if dash: etree.SubElement(ln, qn("a:prstDash")).set("val", "dash")
    if head:
        t = etree.SubElement(ln, qn("a:tailEnd")); t.set("type", "triangle"); t.set("w", "med"); t.set("len", "med")
    return c

def _poly(sl, pts, stops, ang=90):
    fb = sl.shapes.build_freeform(int(pts[0][0] * PX), int(pts[0][1] * PX), scale=1.0)
    fb.add_line_segments([(int(a * PX), int(b * PX)) for a, b in pts[1:]], close=True)
    s = fb.convert_to_shape(); s.line.fill.background()
    st = s._element.find(qn("p:style"))
    if st is not None: s._element.remove(st)
    spPr = s._element.find(qn("p:spPr"))
    for e in spPr.findall(qn("a:solidFill")): spPr.remove(e)
    if SOFT:
        spPr.find(qn("a:custGeom")).addnext(_grad_el(stops, ang))
    else:
        s.fill.solid()
        s.fill.fore_color.rgb = _rgb(stops[-1][1])
    return s

def _pts(points):
    return [p if isinstance(p, tuple) else (p, None) for p in points]

def _points_h(points, w, size=16):
    return sum(need_h(a, w - 18, size, 1.35) + (need_h(b, w - 18, max(16, size-1), 1.35) + 2 if b else 0) + 8 for a, b in points)

def _draw_points(sl, x, y, w, points, color, size=16):
    yy = y
    for a, b in points:
        box(sl, x, yy + 7, 7, 7, color, shape=MSO_SHAPE.OVAL if SOFT else MSO_SHAPE.RECTANGLE)
        ah = need_h(a, w - 18, size, 1.35)
        text(sl, x + 16, yy, w - 16, ah + 2, a, size, P("ink"), bool(b), spacing=1.35)
        yy += ah
        if b:
            bh = need_h(b, w - 18, max(16, size-1), 1.35)
            text(sl, x + 16, yy + 2, w - 16, bh + 2, b, max(16, size-1), P("muted"), spacing=1.35); yy += bh + 2
        yy += 8
    return yy

# ============================ 流程 / 顺序 ============================
def chevron_flow(sl, x, y, w, h, steps, ah=60):
    """箭头流程：steps [(标题 ≤8 字, 说明 或 [要点], (数值, 标签) 可选)]，3–6 步，横向排开"""
    n = len(steps); sw = w / n
    for i, st in enumerate(steps):
        t, d = st[0], st[1]; m = st[2] if len(st) > 2 else None
        sx = x + i * sw; col = _c(i)
        s = box(sl, sx, y, sw + (14 if i < n - 1 else 0), ah, shape=MSO_SHAPE.PENTAGON if i == 0 else MSO_SHAPE.CHEVRON, adj=[0.32])
        if SOFT: grad_fill(s, [(0, light(col, 0.15), None), (100, col, None)], 0)
        else: s.fill.solid(); s.fill.fore_color.rgb = _rgb(col)
        text(sl, sx + (14 if i else 8), y, sw - 30, ah, [[(f"{i + 1:02d} ", 15, _mix(col, "FFFFFF", 0.7), True), (t, 18, "FFFFFF", True)]], align="c", anchor="m")
        py = y + ah + 12; pw_ = sw - (14 if i < n - 1 else 0)
        ty = py + 16; tw = pw_ - 32; tx = sx + 16
        mh = 70 if m else 0
        avail = y + h - ty - mh
        if isinstance(d, (list, tuple)):
            pts = _pts(d); z = 18; need = _points_h(pts, tw, z)
            _draw_points(sl, tx, ty, tw, pts, col, z)
        else:
            z = 18; need = need_h(d, tw, z, 1.5); text(sl, tx, ty, tw, need + 4, d, z, P("ink"), spacing=1.5)
        _check(f"箭头流程[{t}]", need, avail)
        if m:
            line(sl, tx, y + h - 66, tx + tw, y + h - 66, light(col, 0.6))
            text(sl, tx, y + h - 60, tw, 52, [[(m[0], 26, _ink(col), True)], [(m[1], 13, P("muted"), False)]], anchor="m", spacing=1.05)

RISK_RED = "B83B32"   # 风险态固定用深红，与主色、强调色都能区分
STATUS_USED = []      # 本份报告里用到状态色的位置；start(status_colors=True) 时装饰交替色避开强调色


def status_color(status, i=0):
    """节点状态 → 颜色：按期/完成/风险可控 用主色，滞后/延期 用强调色，风险/重点保障 用深红；
    风险等级 高/中/低 依次为深红、强调色、主色；也可直接给 6 位色值。"""
    st = (status or "").strip().lstrip("#")
    if len(st) == 6 and all(c in "0123456789abcdefABCDEF" for c in st):
        return st.upper()
    level = {"高": RISK_RED, "高风险": RISK_RED, "中": A(), "中等": A(), "中风险": A(), "低": N(), "低风险": N()}.get(st)
    if level:                                   # 风险等级：高 / 中 / 低
        STATUS_USED.append(st); return level
    if any(k in st for k in ("可控", "受控")):   # “风险可控”是受控状态，不与“需重点保障”同用深红
        STATUS_USED.append(st); return N()
    if any(k in st for k in ("风险", "保障", "预警", "受阻", "risk")):
        STATUS_USED.append(st); return RISK_RED
    if any(k in st for k in ("滞后", "延期", "延误", "推迟", "delay")):
        STATUS_USED.append(st); return A()
    if any(k in st for k in ("按期", "完成", "正常", "达成", "提前", "ok", "done")):
        STATUS_USED.append(st); return N()
    return _c(i)


def timeline(sl, x, y, w, h, events, legend=None):
    """时间轴：events [(时间, 标题, 说明[, 状态])]，3–10 个，卡片上下交错。
    第 4 项写状态（按期 / 滞后 / 风险，或 6 位色值）时，节点、连线和卡片按状态着色，否则主色与强调色交替。
    legend=True 时在右上角画状态图例。"""
    n = len(events); slot = w / n; ay = y + h / 2
    bar = box(sl, x, ay - 3, w, 6, shape=MSO_SHAPE.ROUNDED_RECTANGLE, adj=[0.5])
    grad_fill(bar, [(0, light(N(), 0.5), None), (100, light(A(), 0.3), None)], 0)
    cw = slot * 1.55; ch = h / 2 - 46
    statuses = [ev[3] if len(ev) > 3 else None for ev in events]
    if legend is None:
        legend = any(statuses)
    if legend and any(statuses):
        seen = []
        for st in statuses:
            if st and status_color(st) not in [c for _, c in seen]:
                seen.append((st, status_color(st)))
        lx = x + w
        for st, col in reversed(seen):
            tw = 18 * len(st) + 34; lx -= tw
            _badge(sl, lx + 8, y - 14, 12, "", col)
            text(sl, lx + 20, y - 26, tw - 20, 24, st, 16, P("ink"), anchor="m", spacing=1.0)
    for i, ev in enumerate(events):
        tm, t, d = ev[:3]
        cx = x + slot * (i + 0.5); col = status_color(statuses[i], i); up = i % 2 == 0
        _badge(sl, cx, ay, 22, "", col)
        text(sl, cx - slot / 2, ay + (10 if up else -42), slot, 32, tm, 18, _ink(col), True, align="c", anchor="m")
        cy = y if up else ay + 46
        line(sl, cx, ay + (-12 if up else 12), cx, cy + (ch if up else 0), light(col, 0.4), 1.5)
        bx = max(x, min(x + w - cw, cx - cw / 2))
        _panel(sl, bx, cy, cw, ch, col)
        text(sl, bx + 16, cy + 10, cw - 32, 28, t, 18, _ink(col), True, anchor="m")
        if not d:
            continue
        z = 18; need = need_h(d, cw - 32, z, 1.45)
        text(sl, bx + 16, cy + 42, cw - 32, ch - 50, d, z, P("ink"), spacing=1.45)
        _check(f"时间轴[{tm}]", need, ch - 50)

def stairs(sl, x, y, w, h, steps, gap=14):
    """阶梯柱：steps [(标题, 说明 或 [要点], (数值, 标签) 可选)]，3–5 级，柱子逐级升高（递进、提升、目标分解）"""
    n = len(steps); cw = (w - gap * (n - 1)) / n
    for i, st in enumerate(steps):
        t, d = st[0], st[1]; m = st[2] if len(st) > 2 else None
        bh = h * (0.62 + 0.38 * i / max(1, n - 1)); bx = x + i * (cw + gap); by = y + h - bh; col = _c(i)
        if SOFT:
            s = box(sl, bx, by, cw, bh, shape=MSO_SHAPE.ROUNDED_RECTANGLE, adj=[0.06])
            grad_fill(s, [(0, light(col, 0.82), None), (100, light(col, 0.95), None)], 90)
        else:
            box(sl, bx, by, cw, bh, P("tint"))
        box(sl, bx, by, cw, 6, col)
        text(sl, bx + 16, by + 14, cw - 32, 50, [[(f"{i + 1:02d}", 34, light(col, 0.45) if col == N() else col, True)]], anchor="m")
        text(sl, bx + 16, by + 64, cw - 32, 30, t, 19, _ink(col), True, anchor="m")
        ty = by + 102; mh = 70 if m else 0; avail = y + h - ty - mh - 10
        if isinstance(d, (list, tuple)):
            pts = _pts(d); z = 18; need = _points_h(pts, cw - 32, z); _draw_points(sl, bx + 16, ty, cw - 32, pts, col, z)
        else:
            z = 18; need = need_h(d, cw - 32, z, 1.45); text(sl, bx + 16, ty, cw - 32, avail, d, z, P("ink"), spacing=1.45)
        _check(f"阶梯[{t}]", need, avail)
        if m:
            line(sl, bx + 16, y + h - 72, bx + cw - 16, y + h - 72, light(col, 0.6))
            text(sl, bx + 16, y + h - 66, cw - 32, 58, [[(m[0], 28, _ink(col), True)], [(m[1], 13, P("muted"), False)]], anchor="m", spacing=1.05)

def gantt(sl, x, y, w, h, periods, tasks, label_w=190):
    """进度计划（甘特）：periods 时间段名；tasks [(事项, 起始序号, 结束序号, 说明)]，序号从 0 起，可为小数；4–8 项"""
    n = len(periods); gw = (w - label_w) / n; hh = 34
    for j, p in enumerate(periods):
        px = x + label_w + j * gw
        b = box(sl, px + 3, y, gw - 6, hh, P("tint") if not SOFT else light(N(), 0.88), shape=MSO_SHAPE.ROUNDED_RECTANGLE if SOFT else MSO_SHAPE.RECTANGLE, adj=[0.5] if SOFT else None)
        text(sl, px, y, gw, hh, p, 14, _ink(N()), True, align="c", anchor="m")
        line(sl, px, y + hh + 6, px, y + h, P("line"), 0.75)
    line(sl, x + w, y + hh + 6, x + w, y + h, P("line"), 0.75)
    rh = (h - hh - 12) / len(tasks)
    for i, (t, a, b, note) in enumerate(tasks):
        ry = y + hh + 12 + i * rh; col = _c(i)
        if i: line(sl, x, ry, x + w, ry, P("line"), 0.5)
        fit_text(sl, x, ry, label_w - 14, rh, t, 16, P("ink"), True, where="甘特事项", anchor="m")
        bx = x + label_w + a * gw + 4; bw = max(24, (b - a + 1) * gw - 8); bhh = min(30, rh * 0.55)
        s = box(sl, bx, ry + (rh - bhh) / 2, bw, bhh, shape=MSO_SHAPE.ROUNDED_RECTANGLE, adj=[0.5])
        grad_fill(s, [(0, light(col, 0.25), None), (100, col, None)], 0)
        if note:
            fits = est_lines(note, bw - 20, 13) == 1
            if fits: text(sl, bx + 10, ry, bw - 20, rh, note, 13, "FFFFFF", True, anchor="m")
            elif bx + bw + 8 < x + w - 60: text(sl, bx + bw + 8, ry, x + w - bx - bw - 8, rh, note, 13, P("muted"), anchor="m")

# ============================ 并列 / 列表 ============================
def numbered_rows(sl, x, y, w, h, rows, metric_w=190, metric_size=26):
    """编号行：rows [(标题, 说明, (数值, 标签) 或 None)]，3–6 行，适合要点、措施、问题清单"""
    n = len(rows); rh = h / n
    has_m = any(r[2] for r in rows)
    tw = w - 100 - (metric_w if has_m else 10)
    heights = [(need_h(t, tw, 18, 1.3), need_h(d, tw, 15, 1.4) if d else 0) for t, d, _ in rows]
    minimum = [th + dh + 16 for th, dh in heights]
    variable = max(minimum) > rh and sum(minimum) <= h
    spare = (h - sum(minimum)) / n
    ry = y
    for i, (t, d, m) in enumerate(rows):
        rh = minimum[i] + spare if variable else h / n
        col = _c(i)
        if SOFT:
            s = box(sl, x, ry + 4, w, rh - 8, light(col, 0.93), shape=MSO_SHAPE.ROUNDED_RECTANGLE, adj=[0.18])
        elif i:
            line(sl, x, ry, x + w, ry, P("line"))
        text(sl, x + 14, ry, 70, rh, f"{i + 1:02d}", 32, light(col, 0.4) if col == N() else col, True, anchor="m")
        th, dh = heights[i]
        top = ry + max(6, (rh - th - dh - 4) / 2)
        text(sl, x + 92, top, tw, th + 2, t, 18, _ink(col), True, spacing=1.3)
        if d: text(sl, x + 92, top + th + 4, tw, dh + 2, d, 15, P("ink"), spacing=1.4)
        _check(f"编号行[{t[:8]}]", th + dh + 4, rh - 12, low=0.2)
        if m:
            mx = x + w - metric_w
            line(sl, mx, ry + 14, mx, ry + rh - 14, light(col, 0.6))
            text(sl, mx + 18, ry, metric_w - 24, rh, [[(m[0], metric_size, _ink(col), True)], [(m[1], 13, P("muted"), False)]], anchor="m", spacing=1.1)
        ry += rh

def big_numbers(sl, x, y, w, h, items, highlight=None, gap=14):
    """大数字列：items [(数值, 标签, 说明)]，3–5 列；highlight 为强调列序号（实色底）"""
    n = len(items); cw = (w - gap * (n - 1)) / n
    from render import preview_font
    sizes = [44 if len(v) <= 6 else (34 if len(v) <= 9 else 26) for v, _, _ in items]
    widths = [preview_font(size).getlength(v) * 1.08 for (v, _, _), size in zip(items, sizes)]
    widest = max(preview_font(max(sizes)).getlength(v) * 1.08 for v, _, _ in items)
    if max(widths) > cw - 40 and w - widest - 128 >= 100:
        rows = [(lab, d, (v, "")) for v, lab, d in items]
        return numbered_rows(sl, x, y, w, h, rows, metric_w=math.ceil(widest) + 26,
                             metric_size=max(sizes))
    for i, (v, lab, d) in enumerate(items):
        cx = x + i * (cw + gap); hl = i == highlight
        if hl: _fill_block(sl, cx, y, cw, h, N())
        else: _panel(sl, cx, y, cw, h, _c(i))
        fg = "FFFFFF" if hl else _ink(_c(i)); sub = _mix(N(), "FFFFFF", 0.82) if hl else P("ink")
        fs = 44 if len(v) <= 6 else (34 if len(v) <= 9 else 26)
        if n == 1 and cw > 320 and h > 220:
            fs = min(84, max(fs, int((cw-50)/max(1,preview_font(1).getlength(v))*.85)))
            text(sl,cx+24,y+h*.19,cw-48,h*.32,v,fs,fg,True,anchor="m")
            text(sl,cx+24,y+h*.52,cw-48,h*.13,lab,22,fg,True,anchor="m")
            text(sl,cx+24,y+h*.68,cw-48,h*.20,d,18,sub,anchor="m")
            continue
        if h < 160:
            label_h = need_h(lab,cw-24,16,1.05)
            note_h = need_h(d,cw-24,16,1.05) if d else 0
            vh = h-label_h-note_h-12
            text(sl,cx+12,y+3,cw-24,vh,v,min(fs,max(16,int(vh*.8))),fg,True,anchor="m")
            text(sl,cx+12,y+vh+4,cw-24,label_h,lab,16,fg,True,anchor="m",spacing=1.05)
            if d: text(sl,cx+12,y+vh+label_h+6,cw-24,note_h,d,16,sub,anchor="m",spacing=1.05)
            continue
        z = 18; need = need_h(d, cw - 40, z, 1.5)
        oy = max(0, (h - 146 - need) / 2 - 10)          # 内容整体垂直居中，避免下半截空白
        yy = y + oy
        text(sl, cx + 20, yy + 18, cw - 40, 60, v, fs, fg, True, anchor="m")
        text(sl, cx + 20, yy + 82, cw - 40, 26, lab, 17, fg, True, anchor="m")
        line(sl, cx + 20, yy + 118, cx + 60, yy + 118, "FFFFFF" if hl else _c(i), 2.5)
        text(sl, cx + 20, yy + 132, cw - 40, need + 6, d, z, sub, spacing=1.5)
        _check(f"大数字[{lab}]", need, h - 146, low=0.5)

def stat_strip(sl, x, y, w, h, stats):
    """指标通栏：一条主色实底，stats [(数值, 标签, 注释)] 3–5 个，白字。适合放在页面上部做总览"""
    _fill_block(sl, x, y, w, h, N(), radius=0.12)
    n = len(stats); cw = w / n
    for i, (v, lab, note) in enumerate(stats):
        cx = x + i * cw
        if i: line(sl, cx, y + 20, cx, y + h - 20, _mix(N(), "FFFFFF", 0.4), 1)
        # 上下各留 10px，标签、注释各 22px；数值框占剩余高度，字号随框高收缩，避免压住注释
        note_h = 22 if note else 0
        vh = h - 20 - 22 - note_h - 4
        fs = min(34 if len(v) <= 8 else 26, int(vh / 1.25))
        if fs < 28:
            WARN.append(f"指标通栏：高 {h:.0f}px 放不下 28px 以上的数值，至少给 {(28 * 1.25 + 46 + note_h):.0f}px")
            fs = max(fs, 20)
        text(sl, cx + 24, y + 10, cw - 40, 22, lab, 15, _mix(N(), "FFFFFF", 0.8), anchor="m", spacing=1.0)
        text(sl, cx + 24, y + 34, cw - 40, vh, v, fs, "FFFFFF", True, anchor="m", spacing=1.0)
        if note: text(sl, cx + 24, y + h - 10 - note_h, cw - 40, note_h, note, 13, _mix(A(), "FFFFFF", 0.35) if not SOFT else _mix(N(), "FFFFFF", 0.85), True, anchor="m", spacing=1.0)

def callout(sl, x, y, w, h, title, body, metric=None):
    """重点框：主色实底的观点块（核心判断、一句话结论、重要提示）；metric=(数值, 标签)"""
    _fill_block(sl, x, y, w, h, N())
    text(sl, x + 22, y + 6, 60, 60, "“", 64, _mix(N(), "FFFFFF", 0.45), True)
    text(sl, x + 26, y + 58, w - 52, 60, title, 22, "FFFFFF", True, spacing=1.3)
    th = need_h(title, w - 52, 22, 1.3)
    mh = 90 if metric else 0
    avail = h - 58 - th - 20 - mh - 16; z = 18; need = need_h(body, w - 52, z, 1.6)
    text(sl, x + 26, y + 58 + th + 16, w - 52, avail, body, z, _mix(N(), "FFFFFF", 0.88), spacing=1.6)
    _check("重点框", need, avail)
    if metric:
        line(sl, x + 26, y + h - 92, x + w - 26, y + h - 92, _mix(N(), "FFFFFF", 0.4))
        text(sl, x + 26, y + h - 84, w - 52, 74, [[(metric[0], 36, "FFFFFF", True)], [(metric[1], 14, _mix(N(), "FFFFFF", 0.8), False)]], anchor="m", spacing=1.05)

def progress_bars(sl, x, y, w, h, rows, label_w=170, note_w=None):
    """进度条：rows [(名称, 百分比 0–100, 数值文本, 说明[, 状态[, 计划百分比]])]，3–7 行（完成率、占比、达成度）。
    第 5 项写状态（按期 / 滞后 / 风险，或 6 位色值）时该行按状态着色并在数值下标出状态；
    第 6 项给计划百分比时在条上画计划刻度，实际与计划的差距一眼可见。
    行高按文字收紧（每行不超过 120px），多出的高度留给同页其他内容，不把三五行拉满整个区域。"""
    n = len(rows)
    note_w = note_w if note_w is not None else (w * 0.3 if any(r[3] for r in rows) else 0)
    bw = w - label_w - 120 - note_w
    has_status = any(len(r) > 4 and r[4] for r in rows)
    rh = min(h / n, 120)
    spare = h - rh * n
    if spare > max(110, h * 0.28):
        WARN.append(f"进度条: 区域未填满，{n} 行按文字收紧后只占区域高度的 {rh * n / h:.0%}，下方空出 {int(spare)}px。"
                    "把这一区的高度调小，空出的位置放同页其他内容，或与相邻页合并")
    for i, row in enumerate(rows):
        lab, pct, vt, note = row[:4]
        status = row[4] if len(row) > 4 else None
        plan = row[5] if len(row) > 5 else None
        ry = y + i * rh; col = status_color(status, i) if status else (N() if has_status else _c(i)); cy = ry + rh / 2
        fit_text(sl, x, ry, label_w - 14, rh, lab, 17, P("ink"), True, where="进度条名称", anchor="m")
        bx = x + label_w
        box(sl, bx, cy - 9, bw, 18, light(col, 0.88), shape=MSO_SHAPE.ROUNDED_RECTANGLE, adj=[0.5])
        f = box(sl, bx, cy - 9, max(18, bw * min(100, max(0, pct)) / 100), 18, shape=MSO_SHAPE.ROUNDED_RECTANGLE, adj=[0.5])
        grad_fill(f, [(0, light(col, 0.35), None), (100, col, None)], 0)
        if plan is not None:
            px_ = bx + bw * min(100, max(0, plan)) / 100
            line(sl, px_, cy - 16, px_, cy + 16, P("ink"), 2)
            text(sl, px_ - 60, cy + 16, 120, 20, f"计划 {plan:g}%", 13, P("muted"), align="c", anchor="m", spacing=1.0)
        if status and not all(c in "0123456789abcdefABCDEF" for c in status.strip().lstrip("#")):
            text(sl, bx + bw + 14, ry, 106, rh, [[(vt, 22, _ink(col), True)], [(status, 14, _ink(col), True)]], anchor="m", spacing=1.05)
        else:
            text(sl, bx + bw + 14, ry, 106, rh, vt, 22, _ink(col), True, anchor="m")
        if note_w and note:
            fit_text(sl, x + w - note_w, ry, note_w, rh, note, 16, P("muted"), where="进度条说明", anchor="m", spacing=1.35)

def pictogram(sl, x, y, w, h, items, gap=24):
    """人形占比：items [(标签, 百分比, 说明)]，1–3 组；每组 10 个人形，按比例着色"""
    n = len(items); cw = (w - gap * (n - 1)) / n
    for i, (lab, pct, d) in enumerate(items):
        cx = x + i * (cw + gap); col = _c(i)
        if n > 1 or SOFT: _panel(sl, cx, y, cw, h, col)
        iw = min(cw / 10 * 0.62, 34); step = (cw - 40) / 10; k = round(pct / 10)
        top = y + 26
        for j in range(10):
            c = col if j < k else "D8DCE3"
            px = cx + 20 + j * step + (step - iw) / 2
            box(sl, px + iw * 0.25, top, iw * 0.5, iw * 0.5, c, shape=MSO_SHAPE.OVAL)
            box(sl, px, top + iw * 0.56, iw, iw * 1.05, c, shape=MSO_SHAPE.ROUND_2_SAME_RECTANGLE, adj=[0.45, 0])
        ty = top + iw * 1.7 + 14
        text(sl, cx + 20, ty, cw - 40, 54, [[(f"{pct:g}%", 38, _ink(col), True), ("  " + lab, 17, P("ink"), True)]], anchor="m")
        avail = y + h - ty - 70; z = 18; need = need_h(d, cw - 40, z, 1.5)
        text(sl, cx + 20, ty + 62, cw - 40, avail, d, z, P("ink"), spacing=1.5)
        _check(f"人形占比[{lab}]", need, avail)

# ============================ 对比 / 矩阵 / 关系 ============================
def matrix2x2(sl, x, y, w, h, quads, axes=("横轴", "纵轴"), highlight=None):
    """四象限：quads 4 个 (标题, 说明)，顺序为 左上、右上、左下、右下；axes=(横轴名, 纵轴名)，箭头方向为“高”"""
    ax = 34; gx = x + ax; gw = w - ax; gh = h - 30; g = 10
    cw = (gw - g) / 2; ch = (gh - g) / 2
    for i, (t, d) in enumerate(quads):
        cx = gx + (i % 2) * (cw + g); cy = y + (i // 2) * (ch + g); hl = i == highlight
        if hl: _fill_block(sl, cx, cy, cw, ch, N(), radius=0.05)
        else:
            if SOFT: box(sl, cx, cy, cw, ch, light(_c(i), 0.92), shape=MSO_SHAPE.ROUNDED_RECTANGLE, adj=[0.05])
            else: box(sl, cx, cy, cw, ch, P("tint"))
        fg = "FFFFFF" if hl else _ink(_c(i)); sub = _mix(N(), "FFFFFF", 0.85) if hl else P("ink")
        text(sl, cx + 20, cy + 14, cw - 40, 32, t, 19, fg, True, anchor="m")
        z = 18; need = need_h(d, cw - 40, z, 1.5)
        text(sl, cx + 20, cy + 52, cw - 40, ch - 62, d, z, sub, spacing=1.5)
        _check(f"四象限[{t}]", need, ch - 62)
    _arrow(sl, gx - 12, y + gh + 8, x + w, y + gh + 8, P("muted"), 1.5)
    _arrow(sl, gx - 12, y + gh + 8, gx - 12, y - 4, P("muted"), 1.5)
    text(sl, x + w - 240, y + gh + 10, 240, 20, axes[0] + " →", 13, P("muted"), True, align="r", anchor="m")
    tb = text(sl, x - 70, y + 80, 150, 20, axes[1] + " →", 13, P("muted"), True, align="c", anchor="m"); tb.rotation = -90

def venn(sl, x, y, w, h, sets, center=None, list_w=None):
    """韦恩图：sets [(名称 ≤6 字, 说明)] 2–3 个；center 为交集名称；右侧列出各集合说明"""
    n = len(sets); list_w = list_w or w * 0.48
    vw = w - list_w - 30; d = min(h * (0.78 if n == 2 else 0.62), vw * (0.62 if n == 2 else 0.6))
    cx0, cy0 = x + vw / 2, y + h / 2
    if n == 2: pos = [(cx0 - d * 0.3, cy0), (cx0 + d * 0.3, cy0)]
    else: pos = [(cx0 - d * 0.3, cy0 - d * 0.2), (cx0 + d * 0.3, cy0 - d * 0.2), (cx0, cy0 + d * 0.3)]
    for i, ((t, _), (px, py)) in enumerate(zip(sets, pos)):
        col = [N(), A(), light(N(), 0.35)][i]
        s = box(sl, px - d / 2, py - d / 2, d, d, shape=MSO_SHAPE.OVAL)
        grad_fill(s, [(0, col, 62), (100, col, 80)], 45)
        ox = (px - cx0) * 1.1; oy = (py - (cy0 if n == 2 else cy0 + d * 0.0)) * 1.0
        text(sl, px + ox - d * 0.32, py + oy * 0.9 - 18, d * 0.64, 36, t, 19, "FFFFFF", True, align="c", anchor="m")
    if center:
        text(sl, cx0 - d * 0.3, cy0 - 14 + (0 if n == 2 else -d * 0.02), d * 0.6, 28, center, 15, "FFFFFF", True, align="c", anchor="m")
    lx = x + w - list_w; rh = h / n
    for i, (t, ds) in enumerate(sets):
        ry = y + i * rh; col = [N(), A(), light(N(), 0.35)][i]
        box(sl, lx, ry + 10, 6, rh - 20, col)
        text(sl, lx + 20, ry + 8, list_w - 20, 30, t, 19, _ink(col), True, anchor="m")
        z = 18; need = need_h(ds, list_w - 20, z, 1.5)
        text(sl, lx + 20, ry + 42, list_w - 20, rh - 52, ds, z, P("ink"), spacing=1.5)
        _check(f"韦恩[{t}]", need, rh - 52)

def hub(sl, x, y, w, h, center, nodes):
    """中心发散：center=(标题, 副标题)；nodes [(标题, 说明)] 4–8 个，左右两列，与中心圆虚线相连"""
    n = len(nodes); d = min(h * 0.56, w * 0.24, 230); cx, cy = x + w / 2, y + h / 2
    left = nodes[: (n + 1) // 2]; right = nodes[(n + 1) // 2:]
    colw = (w - d) / 2 - 50
    if SOFT:
        r = box(sl, cx - d / 2 - 14, cy - d / 2 - 14, d + 28, d + 28, shape=MSO_SHAPE.OVAL); grad_fill(r, [(0, N(), 8), (100, A(), 14)], 45)
    if SOFT: grad_badge(sl, cx, cy, d, "", N())
    else:
        disc = box(sl, cx - d / 2, cy - d / 2, d, d, shape=MSO_SHAPE.OVAL)
        grad_fill(disc, [(0, N(), None), (100, dark(N(), 0.2), None)], 45)
    text(sl, cx - d / 2 + 14, cy - 34, d - 28, 40, center[0], 24, "FFFFFF", True, align="c", anchor="m")
    if len(center) > 1 and center[1]:
        text(sl, cx - d / 2 + 20, cy + 8, d - 40, 44, center[1], 14, _mix(N(), "FFFFFF", 0.85), align="c", spacing=1.3)
    k = 0
    tw_all = colw - 62
    for side, items in (("l", left), ("r", right)):
        if not items:
            continue
        # 每侧按文字定行高，整组垂直居中：5 个节点分成 3/2 时右侧不会出现半张空卡
        needs = [need_h(ds, tw_all, 18, 1.4) + 54 + 12 for _, ds in items]
        side_h = min(h, sum(needs) * 1.2)
        top = y + (h - side_h) / 2
        heights = [side_h * v / sum(needs) for v in needs]
        ry = top
        for j, (t, ds) in enumerate(items):
            rh = heights[j]
            k += 1; col = _c(k - 1)
            bx = x if side == "l" else x + w - colw
            if SOFT: soft_card(sl, bx, ry + 6, colw, rh - 12, col, radius=0.1)
            else: box(sl, bx, ry + 6, colw, rh - 12, P("tint"))
            _badge(sl, bx + (colw - 22 if side == "l" else 22), ry + rh / 2, 30, str(k), col, 14)
            tx = bx + (16 if side == "l" else 46); tw = tw_all
            text(sl, tx, ry + 14, tw, 26, t, 17, _ink(col), True, anchor="m")
            z = 18; need = need_h(ds, tw, z, 1.4)
            text(sl, tx, ry + 42, tw, rh - 54, ds, z, P("ink"), spacing=1.4)
            _check(f"中心发散[{t}]", need, rh - 54, low=0.2)
            ex = bx + colw if side == "l" else bx
            ang = math.atan2(ry + rh / 2 - cy, ex - cx)
            sx, sy = cx + (d / 2 + 4) * math.cos(ang), cy + (d / 2 + 4) * math.sin(ang)
            _arrow(sl, sx, sy, ex + (6 if side == "l" else -6), ry + rh / 2, light(col, 0.35), 1.5, dash=True, head=False)
            ry += rh

def cycle(sl, x, y, w, h, steps, center=None):
    """循环：steps [(标题, 说明)] 3–6 个，围成一圈（闭环管理、PDCA、周期性工作）；center 为圆心文字"""
    n = len(steps); d = min(h - 10, w * 0.42); cx, cy = x + w / 2, y + h / 2
    gap = 4
    for i in range(n):
        st = -90 + i * 360 / n + gap / 2; en = st + 360 / n - gap
        s = box(sl, cx - d / 2, cy - d / 2, d, d, shape=MSO_SHAPE.BLOCK_ARC)
        s.adjustments[0] = (st % 360) * 0.6; s.adjustments[1] = (en % 360) * 0.6; s.adjustments[2] = 0.2
        col = _c(i)
        grad_fill(s, [(0, light(col, 0.2), None), (100, col, None)], 45)
        mid = math.radians((st + en) / 2); rr = d / 2 - d * 0.1
        text(sl, cx + rr * math.cos(mid) - 20, cy + rr * math.sin(mid) - 16, 40, 32, f"{i + 1:02d}", 15, "FFFFFF", True, align="c", anchor="m")
    if center:
        text(sl, cx - d * 0.3, cy - 30, d * 0.6, 60, center, 22 if len(center) <= 6 else 17, _ink(N()), True, align="c", anchor="m", spacing=1.2)
    colw = (w - d) / 2 - 36
    right = list(range((n + 1) // 2)); left = list(range((n + 1) // 2, n))[::-1]
    for side, idx in (("r", right), ("l", left)):
        m = len(idx)
        if not m: continue
        rh = h / m
        for j, i in enumerate(idx):
            t, ds = steps[i]; col = _c(i); ry = y + j * rh
            bx = x + w - colw if side == "r" else x
            al = "l" if side == "r" else "r"
            text(sl, bx, ry + 6, colw, 30, [[(f"{i + 1:02d}  ", 17, col if col != N() else light(N(), 0.2), True), (t, 18, _ink(col), True)]], align=al, anchor="m")
            z = 18; need = need_h(ds, colw, z, 1.45)
            text(sl, bx, ry + 42, colw, rh - 50, ds, z, P("ink"), align=al, spacing=1.45)
            _check(f"循环[{t}]", need, rh - 50, low=0.25)

def tree(sl, x, y, w, h, root, branches, root_w=None):
    """树状分解：root 总目标；branches [(分支名, [要点])] 2–5 个；向下展开（目标分解、组织架构、任务拆分）"""
    n = len(branches); rw = root_w or min(w * 0.42, 420); rx = x + (w - rw) / 2
    rh = max(50, est_lines(root, rw - 32, 19) * 19 * 1.3 + 22)
    _fill_block(sl, rx, y, rw, rh, N(), radius=0.3)
    fit_text(sl, rx + 16, y, rw - 32, rh, root, 19, "FFFFFF", True, where="树根", align="c", anchor="m", spacing=1.3)
    g = 14; bw = (w - g * (n - 1)) / n; jy = y + rh + 22; by = jy + 24
    line(sl, x + w / 2, y + rh, x + w / 2, jy, P("muted"), 1.5)
    line(sl, x + bw / 2, jy, x + w - bw / 2, jy, P("muted"), 1.5)
    for i, (t, pts) in enumerate(branches):
        bx = x + i * (bw + g); col = _c(i)
        line(sl, bx + bw / 2, jy, bx + bw / 2, by, P("muted"), 1.5)
        s = box(sl, bx, by, bw, 44, col, shape=MSO_SHAPE.ROUNDED_RECTANGLE if SOFT else MSO_SHAPE.RECTANGLE, adj=[0.3] if SOFT else None)
        if SOFT: grad_fill(s, [(0, light(col, 0.2), None), (100, col, None)], 0)
        fit_text(sl, bx + 10, by, bw - 20, 44, t, 17, "FFFFFF", True, where="树枝", align="c", anchor="m")
        py = by + 54; ph = y + h - py
        if SOFT: box(sl, bx, py, bw, ph, light(col, 0.93), shape=MSO_SHAPE.ROUNDED_RECTANGLE, adj=[0.05])
        else: box(sl, bx, py, bw, ph, P("tint"))
        pts = _pts(pts); z = 18; need = _points_h(pts, bw - 32, z)
        _draw_points(sl, bx + 16, py + 14, bw - 32, pts, col, z)
        _check(f"树枝[{t}]", need, ph - 20)

def funnel(sl, x, y, w, h, layers, fw=None):
    """漏斗：layers [(层名 ≤6 字, 数值文本, 说明)] 3–6 层，自上而下收窄（转化、筛选、逐级落实）"""
    n = len(layers); fw = fw or w * 0.42; g = 6; lh = (h - g * (n - 1)) / n
    k = 0.62 / n
    for i, (t, v, d) in enumerate(layers):
        ly = y + i * (lh + g); w0 = fw * (1 - i * k); w1 = fw * (1 - (i + 1) * k); c0 = x + fw / 2
        col = _mix(N(), A(), i / max(1, n - 1))
        _poly(sl, [(c0 - w0 / 2, ly), (c0 + w0 / 2, ly), (c0 + w1 / 2, ly + lh), (c0 - w1 / 2, ly + lh)],
              [(0, light(col, 0.15), None), (100, col, None)], 90)
        text(sl, c0 - w1 / 2 + 6, ly, w1 - 12, lh, t, 17 if lh > 50 else 15, "FFFFFF", True, align="c", anchor="m")
        lx = x + fw + 24; ex = c0 + (w0 + w1) / 4
        _arrow(sl, ex + 6, ly + lh / 2, lx - 6, ly + lh / 2, light(col, 0.4), 1, dash=True, head=False)
        text(sl, lx, ly, 170, lh, v, 24, _ink(col), True, anchor="m")
        tw = x + w - lx - 180
        z = 18
        fit_text(sl, lx + 180, ly, tw, lh, d, z, P("ink"), where=f"漏斗[{t}]", anchor="m", spacing=1.4)

# ============================ 数值 ============================
def gauges(sl, x, y, w, h, items, gap=16):
    """环形进度：items [(标签, 百分比 0–100, 中心数值文本, 说明)] 2–5 个（达成率、完成度、占比）"""
    n = len(items); cw = (w - gap * (n - 1)) / n
    for i, (lab, pct, vt, d) in enumerate(items):
        cx0 = x + i * (cw + gap); col = _c(i)
        if SOFT: soft_card(sl, cx0, y, cw, h, col)
        dd = min(cw - 70, h * 0.5, 170); gx = cx0 + cw / 2; gy = y + 24 + dd / 2
        bg = box(sl, gx - dd / 2, gy - dd / 2, dd, dd, light(col, 0.85), shape=MSO_SHAPE.DONUT, adj=[0.13])
        if pct > 0:
            a = box(sl, gx - dd / 2, gy - dd / 2, dd, dd, shape=MSO_SHAPE.BLOCK_ARC)
            st = 270; en = (270 + min(359.5, pct * 3.6)) % 360
            a.adjustments[0] = st * 0.6; a.adjustments[1] = en * 0.6; a.adjustments[2] = 0.13
            grad_fill(a, [(0, light(col, 0.3), None), (100, col, None)], 45)
        fs = 30 if len(vt) <= 5 else (24 if len(vt) <= 7 else 19)
        text(sl, gx - dd / 2 + 10, gy - 24, dd - 20, 48, vt, fs, _ink(col), True, align="c", anchor="m")
        text(sl, cx0 + 14, gy + dd / 2 + 12, cw - 28, 26, lab, 17, P("ink"), True, align="c", anchor="m")
        ty = gy + dd / 2 + 44; avail = y + h - ty - 10
        z = 16; need = need_h(d, cw - 36, z, 1.45) if d else 0
        if d: text(sl, cx0 + 18, ty, cw - 36, avail, d, z, P("muted"), align="c", spacing=1.45)
        _check(f"环形进度[{lab}]", need, avail, low=0)

ALL = ["chevron_flow", "timeline", "stairs", "gantt", "numbered_rows", "big_numbers", "stat_strip", "callout",
       "progress_bars", "pictogram", "matrix2x2", "venn", "hub", "cycle", "tree", "funnel", "gauges"]

"""页面骨架（除标准页外的四种）与分区工具。

标准页     page()            页眉 + 主体区 + 底部结论带（皮肤自带）
开放页     page_open()       页眉 + 导语一句 + 通栏主体区，无结论带（结论写在导语里）
侧栏页     page_side()       左侧主色栏放章节、标题、导语、关键数字；右侧为主体区
分屏页     page_split()      上部主色通栏放标题与 3–5 个指标；下部为主体区
要点页     page_statement()  一个超大数字 + 一句判断 + 3–4 个佐证数字，整页无其他内容

每个骨架函数返回 (slide, area)，area = (x, y, w, h) 为主体区；要点页的 area 为 None。
分区：cols(area, [2, 1]) / rows(area, [1, 1]) 按比例切成若干格，每格再放一个区域图示。
"""
from pptx.util import Emu
from pptx.enum.shapes import MSO_SHAPE
from compose import alt, page_number, box, text, line, est_lines, W, H, M, PX, _mix
from deco import grad_fill, backdrop, light, dark, shadow, half_stadium
from skin_solid import P, BRAND, WARN, fit_text

def cols(area, ratios, gap=16):
    x, y, w, h = area; tot = sum(ratios); avail = w - gap * (len(ratios) - 1); out = []; cx = x
    for r in ratios:
        cw = avail * r / tot; out.append((cx, y, cw, h)); cx += cw + gap
    return out

def rows(area, ratios, gap=16):
    x, y, w, h = area; tot = sum(ratios); avail = h - gap * (len(ratios) - 1); out = []; cy = y
    for r in ratios:
        ch = avail * r / tot; out.append((x, cy, w, ch)); cy += ch + gap
    return out

def _logo(sl, x, y, h, white=False, right=True, maxw=220):
    p = (BRAND.get("logo_card") or BRAND.get("logo_white") or BRAND.get("logo")) if white else BRAND.get("logo")
    if not p:
        if BRAND.get("name"): text(sl, x - (maxw if right else 0), y, maxw, h, BRAND["name"], 14, "FFFFFF" if white else P("navy"), True, align="r" if right else "l", anchor="m")
        return
    from PIL import Image
    iw, ih = Image.open(p).size; lw = min(maxw, h * iw / ih); lh = lw * ih / iw
    if white and BRAND.get("logo_card"):
        box(sl, (x-lw if right else x)-3, y-3, lw+6, lh+6, "FFFFFF",
            shape=MSO_SHAPE.ROUNDED_RECTANGLE, adj=[0.12])
    sl.shapes.add_picture(p, Emu(int((x - lw if right else x) * PX)), Emu(int(y * PX)), width=Emu(int(lw * PX)), height=Emu(int(lh * PX)))

def _foot(sl, x0, source, num, white=False):
    col = _mix(P("navy"), "FFFFFF", 0.75) if white else P("muted")
    if source: text(sl, x0, 694, W - M - x0 - 60, 16, source, 11, col, anchor="m", role="source")
    if num: page_number(sl, num, P("navy"))

def page_open(prs, skin, section, title, lead, num=None, source=None):
    """开放页：用皮肤的页眉，标题下加一句导语（即本页结论），主体区通栏到底"""
    sl = skin.page(prs, section, title, num=num, source=source)
    soft = skin.__name__ == "skin_soft"
    ly = 112 if soft else 78
    if lead:
        if est_lines(lead, W - 2 * M - 40, 17) > 1: WARN.append(f"p{num} 导语超过一行：{lead[:16]}…")
        if soft:
            b = box(sl, M, ly, W - 2 * M, 34, light(P("navy"), 0.93), shape=MSO_SHAPE.ROUNDED_RECTANGLE, adj=[0.5])
            text(sl, M + 18, ly, W - 2 * M - 36, 34, lead, 17, dark(P("navy"), 0.2), True, anchor="m")
        else:
            box(sl, M, ly + 4, 4, 24, P("peach"))
            text(sl, M + 16, ly, W - 2 * M - 16, 32, lead, 17, P("ink"), True, anchor="m")
    top = ly + 50
    return sl, (M, top, W - 2 * M, 676 - top)

def page_side(prs, skin, section, title, lead=None, metrics=(), num=None, source=None, side_w=390):
    """侧栏页：左栏主色，放章节名、标题（可两三行）、导语和 1–3 个关键数字 (数值, 标签)；右侧为主体区"""
    sl = prs.slides.add_slide(prs.slide_layouts[6]); soft = skin.__name__ == "skin_soft"
    pnl = box(sl, 0, 0, side_w, H)
    grad_fill(pnl, [(0, light(P("navy"), 0.12) if soft else P("navy"), None), (100, dark(P("navy"), 0.35), None)], 60 if soft else 90, large_block=True)
    if soft:
        backdrop(sl, "circle", side_w - 220, H - 260, 380, 380, "FFFFFF", 10, 0, ang=135)
        backdrop(sl, "circle", W - 300, -260, 560, 560, P("navy"), 6, 0, ang=135)
    _logo(sl, 44, 40, 40, white=True, right=False, maxw=side_w - 88)
    text(sl, 44, 112, side_w - 88, 24, section, 15, _mix(P("navy"), "FFFFFF", 0.75), True, anchor="m", heading=True, role="section")
    box(sl, 44, 144, 48, 4, P("peach") if not soft else "FFFFFF")
    tl = est_lines(title, side_w - 88, 30)
    if tl > 4: WARN.append(f"p{num} 侧栏标题过长：{title}")
    th = tl * 30 * 1.3 + 6
    text(sl, 44, 164, side_w - 88, th, title, 30, "FFFFFF", True, spacing=1.3, heading=True, role="title")
    yy = 164 + th + 18
    if lead:
        lh = est_lines(lead, side_w - 88, 16) * 16 * 1.6 + 6
        text(sl, 44, yy, side_w - 88, lh, lead, 16, _mix(P("navy"), "FFFFFF", 0.85), spacing=1.6); yy += lh + 10
    if metrics:
        mh = 84; my = H - 60 - mh * len(metrics)
        if my < yy: WARN.append(f"p{num} 侧栏内容过多，指标与导语重叠")
        for i, (v, lab) in enumerate(metrics):
            y0 = my + i * mh
            line(sl, 44, y0, side_w - 44, y0, _mix(P("navy"), "FFFFFF", 0.4), 1)
            text(sl, 44, y0 + 8, side_w - 88, mh - 12, [[(v, 34, "FFFFFF", True)], [(lab, 14, _mix(P("navy"), "FFFFFF", 0.78), False)]], anchor="m", spacing=1.05)
    ax = side_w + 40
    if source: text(sl, ax, 694, W - M - ax - 60, 16, source, 11, P("muted"), anchor="m", role="source")
    if num: text(sl, W - M - 40, 694, 40, 16, str(num), 11, P("muted"), align="r", anchor="m", role="page_number")
    return sl, (ax, 44, W - M - ax, 632)

def page_split(prs, skin, section, title, metrics=(), num=None, source=None, band_h=230):
    """分屏页：上部主色通栏放章节、标题和 3–5 个指标 (数值, 标签, 注释)；下部为主体区"""
    sl = prs.slides.add_slide(prs.slide_layouts[6]); soft = skin.__name__ == "skin_soft"
    b = box(sl, 0, 0, W, band_h)
    grad_fill(b, [(0, light(P("navy"), 0.1) if soft else P("navy"), None), (100, dark(P("navy"), 0.3), None)], 20, large_block=True)
    if soft:
        backdrop(sl, "circle", W - 380, -240, 520, 520, "FFFFFF", 12, 0, ang=135)
        backdrop(sl, "circle", W - 200, 40, 300, 300, P("peach"), 22, 0, ang=135)
    text(sl, M, 22, 600, 22, section, 14, _mix(P("navy"), "FFFFFF", 0.75), True, anchor="m", heading=True, role="section")
    if est_lines(title, W - 2 * M - 240, 28) > 1: WARN.append(f"p{num} 标题过长：{title}")
    text(sl, M, 46, W - 2 * M - 240, 44, title, 28, "FFFFFF", True, anchor="m", heading=True, role="title")
    _logo(sl, W - M, 26, 44, white=True)
    if metrics:
        n = len(metrics); mw = (W - 2 * M) / n; my = 108
        for i, mt in enumerate(metrics):
            v, lab = mt[0], mt[1]; note = mt[2] if len(mt) > 2 else None
            mx = M + i * mw
            if i: line(sl, mx, my + 8, mx, band_h - 22, _mix(P("navy"), "FFFFFF", 0.35), 1)
            text(sl, mx + (20 if i else 0), my, mw - 30, 22, lab, 14, _mix(P("navy"), "FFFFFF", 0.78), anchor="m")
            text(sl, mx + (20 if i else 0), my + 24, mw - 30, 52, v, 34 if len(v) <= 8 else 26, "FFFFFF", True, anchor="m")
            if note: text(sl, mx + (20 if i else 0), my + 78, mw - 30, 22, note, 13, _mix(P("peach"), "FFFFFF", 0.4), True, anchor="m")
    _foot(sl, M, source, num)
    top = band_h + 26
    return sl, (M, top, W - 2 * M, 676 - top)

def page_statement(prs, skin, section, big, unit, statement, facts=(), num=None, source=None):
    """要点页：一个超大数字（big + unit）+ 一句判断 + 3–4 个佐证数字 (数值, 标签)，用于章节开篇或关键结论"""
    sl = prs.slides.add_slide(prs.slide_layouts[6]); soft = skin.__name__ == "skin_soft"
    if soft:
        backdrop(sl, "circle", -200, -260, 760, 760, P("navy"), 10, 0, ang=135)
        backdrop(sl, "circle", W - 380, H - 300, 520, 520, P("peach"), 12, 0, ang=315)
    else:
        box(sl, 0, 0, W, H, P("tint")); box(sl, 0, 0, 14, H, P("navy"))
    _logo(sl, W - M, 30, 40)
    text(sl, 110, 96, 800, 28, section, 17, P("navy"), True, anchor="m", heading=True, role="section")
    box(sl, 110, 132, 60, 5, P("peach"))
    text(sl, 100, 160, 1080, 170, [[(big, 140, P("navy"), True), (" " + unit, 40, P("navy"), True)]], anchor="m")
    if est_lines(statement, 1060, 32) > 2: WARN.append(f"p{num} 要点页判断句过长")
    text(sl, 110, 340, 1060, 100, statement, 32, P("ink"), True, spacing=1.3, heading=True, role="title")
    if facts:
        n = len(facts); fw = (W - 220) / n; fy = 500
        line(sl, 110, fy - 14, W - 110, fy - 14, P("line"), 1)
        for i, (v, lab) in enumerate(facts):
            fx = 110 + i * fw
            text(sl, fx, fy, fw - 20, 110, [[(v, 34, dark(alt(i), 0.15) if i % 2 else P("navy"), True)], [(lab, 15, P("muted"), False)]], anchor="t", spacing=1.25)
    _foot(sl, 110, source, num)
    return sl, None

def report_header(sl, section, title, source, number, style, skin, decoration=None):
    """三段式统一页眉；余下的区域完全交给材料分区。"""
    if skin == "soft":
        backdrop(sl, "circle", W - 260, -350, 520, 520, P("navy"), 6, 0)
    if style == "ribbon":
        ribbon = box(sl, 0, 0, W, 96)
        grad_fill(ribbon, [(0, P("dark"), None), (100, P("navy"), None)], 0, large_block=True)
    color = "FFFFFF" if style == "ribbon" else P("navy")
    text(sl, M, 12, W-2*M-180, 24, section, 16, color, True, anchor="m", heading=True, role="section")
    width = W-2*M-145
    font = next((s for s in (32,30,28,26,24) if est_lines(title,width,s) == 1), 22)
    text(sl, M, 39, width, 49, title, font, color, True, anchor="m", heading=True, role="title")
    _logo(sl, W-M, 19, 56, white=style=="ribbon", maxw=125)
    if decoration:
        decoration(sl,(M,90,W-2*M,54))
    else:
        box(sl, M, 95, W-2*M, 2, P("line"))
        box(sl, M, 94, 100, 4, P("peach"))
    _foot(sl, M, source, number)

def region_heading(sl, rect, title, index, style):
    x,y,w,_ = rect
    width = 43 if style == "number" else 9
    bar = box(sl, x, y + 3, w, 36)
    grad_fill(bar, [(0, dark(P("navy"), .32), None), (100, P("navy"), None)], 0, large_block=True)
    bar.name = "分区渐变标题条"
    if style == "square":
        box(sl, x + 6, y + 12, width, 18, P("peach"))
    if style == "number":
        text(sl,x,y+6,width,30,f"{index:02d}",17,"FFFFFF",True,align="c",anchor="m")
    room = w-width-14
    size = next((n for n in (22,20,18,16) if est_lines(title,room,n)==1),16)
    text(sl,x+width+12,y,room,42,title,size,"FFFFFF",True,anchor="m", heading=True)

def evidence_units(sl, area, items, *, kind, skin, aspect=None):
    """items 各含 label、text 和可选 value，每条论据对应一个原生可编辑单元。"""
    from layout_combo import unit_boxes
    from skin_solid import card_text_height as need_h
    from render import preview_font
    if kind not in ("card", "module", "numbered_rows", "big_numbers"):
        raise ValueError("请选择卡片、模块、分项行或指标块")
    if not items or any(not item.get("text") or not item.get("label") for item in items):
        raise ValueError("每个单元需要标签和论据")
    cells = unit_boxes(area, len(items), aspect or {"numbered_rows":5, "big_numbers":1.4}.get(kind,2.3))
    blocks = []
    for item, (x,y,w,h) in zip(items,cells):
        lane = kind == "numbered_rows" and w > 770
        side = min(223, w*.30) if lane else 0
        headw = side-28 if lane else w-36
        head = item["label"]
        value = item.get("value","")
        body = item["text"]
        label_size = next((z for z in range(21,15,-1) if preview_font(z).getlength(head)<headw),16)
        label_h = max(26, need_h(head,headw,label_size,1.12)+3)
        value_size = next((z for z in range(33,17,-1) if preview_font(z).getlength(value)<headw),16)
        value_h = max(33,need_h(value,headw,value_size,1.1)+3) if value else 0
        bh = h-24 if lane else h-label_h-value_h-29
        bw = w-side-36
        fontsize = next((z for z in range(24,15,-1) if need_h(body,bw,z,1.23)+3<=bh),None)
        if fontsize is None or (lane and label_h+value_h>h-14):
            raise ValueError(f"功能区放不下「{head}」的完整论据，请调整分区")
        if kind in ("card", "module"):
            bh = need_h(body, bw, fontsize, 1.23) + 3
            h = label_h + value_h + bh + 29
        blocks.append((item,x,y,w,h,side,headw,label_size,label_h,value_size,value_h,bw,bh,fontsize))
    previous_row = None
    next_y = area[1]
    for idx,(item,x,y,w,h,side,headw,ls,lh,vs,vh,bw,bh,fs) in enumerate(blocks):
        if kind in ("card", "module"):
            if previous_row != y:
                previous_row, row_y = y, next_y
            y = row_y
            next_y = max(next_y, y + h + 12)
        color = P("navy")
        if kind in ("card","module"):
            box(sl,x,y,w,h,light(color,.965) if skin=="soft" else "FFFFFF",P("line"),
                shape=MSO_SHAPE.ROUNDED_RECTANGLE if skin=="soft" else MSO_SHAPE.RECTANGLE,
                adj=[.035] if skin=="soft" else None)
        elif kind == "big_numbers":
            box(sl,x,y,w,h,light(color,.94))
            box(sl,x,y,5,h,color)
        else:
            line(sl,x,y+2,x+w,y+2,P("line"),1.4)
        hy = y + max(10,(h-lh-vh)/2) if side else y+12
        text(sl,x+16,hy,headw,lh,item["label"],ls,color,True,anchor="m",spacing=1.12, heading=True)
        if item.get("value"):
            text(sl,x+16,hy+lh,headw,vh,item["value"],vs,color,True,anchor="m",spacing=1.1)
        by = y+12 if side else hy+lh+vh+4
        text(sl,x+side+16,by,bw,bh,item["text"],fs,P("ink"),anchor="m",spacing=1.23)

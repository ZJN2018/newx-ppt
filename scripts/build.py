# -*- coding: utf-8 -*-
"""NEWX PPT：从 deck.json 生成整套可编辑 PPTX，并渲染逐页 PNG 与总览图。

用法：
    python3 build.py deck.json --out out/ [--evidence evidence.md] [--no-render] [--force]

deck.json 的格式见 references/split.md 第六节与 references/deck-json.md。
生成前先运行规格检查（check_spec.py）；有"错误"时停止，修改 deck.json 后重跑。--force 只用于调试。
"""
import argparse
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

import deckkit as K
import palette as P
import photos as PH
import structure as S
import check_spec
import logo as LG
from pptx.util import Pt
from pptx.oxml.ns import qn
from pptx.enum.shapes import MSO_SHAPE

NO_BG = {"table", "bar_h", "column", "line", "donut", "waterfall", "photo"}
GAP = {"纯卡片": 20, "纯线条": 0, "线卡混用": 0}
LINE_MODES = ("纯线条", "线卡混用")
INSET = 28  # 分隔线两端各缩进，不顶头


# ---------------- 配色 ----------------
def choose_palette(st):
    """settings.preset 可以是色系名，也可以是 LOGO 原色值（"#1E50A2"）；
    没写时依次看 logo_color（直接用原色）、industry（按标题在行业候选色里稳定地挑一个），都没有用藏蓝。"""
    name = st.get("preset")
    if not name and st.get("logo_color"):
        c = st["logo_color"].lstrip("#").upper()
        name = "#" + c if P.usable(c) else P.nearest_preset(c)
    if not name and st.get("industry"):
        cands = P.by_industry(st["industry"], st.get("title", ""))
        name = cands[0] if cands else None
    try:
        pal = P.preset(name or "藏蓝", corner=st.get("corner", "square"))
    except KeyError:
        pal = P.preset("藏蓝", corner=st.get("corner", "square"))
    P.apply(K, pal)
    return pal


# ---------------- 格子与组件 ----------------
class Ctx:
    def __init__(self, out_dir, online, sources=None):
        self.out_dir = out_dir
        self.online = online
        self.sources = sources
        self.used = set()
        self.n_lib = 0  # 已用自带图库图片数
        self.n_web = 0  # 已用联网图片数（用户自己给的图不计）
        self.layout = []  # 每个正文页的格子坐标，供渲染后留白检查
        self.asset_jobs = []  # 原版创意素材页
        self.photo_miss = []  # 本页没取到图的图片格


def get_photo(ctx, topic):
    """取图并计数：图库图片占比不超过 PH.LIB_MAX_SHARE，超出时先联网取，联网取不到才用图库。"""
    lib_ok = (ctx.n_lib + 1) / (ctx.n_lib + ctx.n_web + 1) <= PH.LIB_MAX_SHARE
    p = PH.get(topic, ctx.out_dir, ctx.used, ctx.online, ctx.sources, lib_ok=lib_ok)
    if p and PH.is_library(p):
        ctx.n_lib += 1
    elif p and os.path.abspath(p).startswith(os.path.abspath(ctx.out_dir)):
        ctx.n_web += 1
    return p


def photo(ctx, topic, where):
    """按主题取图；取不到就打印提醒并返回 None（该处不放图）。"""
    p = get_photo(ctx, topic)
    if not p:
        print("提醒：%s 没有取到与「%s」相关的图，已不放图" % (where, topic))
        return None
    im = next((i for i in PH.lib() if os.path.join(PH.DIR, i["file"]) == p), None)
    if im and im.get("avoid"):
        print("提醒：%s 用了 %s（%s），只适合作背景，不要让读者以为是本单位实景" % (where, im["file"], im["avoid"]))
    return p


def draw_photo_cell(sl, box, el, ctx, where):
    """正文图片格：默认留在正文区内，不越过内容区边界、不贴页边。
    只有显式写 "bleed": true 时，才把贴着正文区边缘的一侧铺到页面左右边缘。"""
    x, y, w, h = box
    path = photo(ctx, el.get("topic", ""), where)
    if not path:  # 没取到图：这一格是空的，留白检查要照常量它，该页也不算“带图”
        ctx.photo_miss.append(box)
        return
    if el.get("bleed", False):
        if abs(x - K.BODY_X) < 1.5:
            w, x = w + x, 0
        if abs(x + w - K.BODY_X - K.BODY_W) < 1.5:
            w = 1280 - x
    sat, duo = PH.tone_for(path)
    PH.place(sl, path, x, y, w, h, tone_main=K.MAIN, anchor=el.get("anchor", "center"), sat=sat, duo=duo)
    # 正文配图不压文字（不写"××场景示意"之类的说明），图片只做配图


def draw_leaf(sl, box, leaf, sep, ctx):
    x, y, w, h = box
    el = leaf["element"]
    et = el["type"]
    if et == "asset":  # 原版创意素材放进这一格（出 PPTX 后由 assets.apply 放置）
        ctx.asset_jobs.append(dict({k: v for k, v in el.items() if k != "type"}, slide=ctx.page, area=list(box)))
        return
    if et == "photo":  # 图片格不套卡片、不留内边距，直接贴格子边
        draw_photo_cell(sl, box, el, ctx, "第 %s 页图片格" % ctx.page)
        return
    style = leaf.get("card") or ("dark" if leaf.get("fill") == "重点" else None)
    dark = style == "dark"
    carded = dark or style in ("header", "number", "label") or (sep == "纯卡片" and et not in NO_BG)
    if style is None and carded:
        style = "plain"
    pad = 26 if dark else 24
    if not carded:
        pad = 22 if sep in LINE_MODES else 14  # 至少 14：小标题左侧短竖条挂在留白里
    cx, cy, cw, ch = x + pad, y + pad - 4, w - 2 * pad, h - 2 * pad + 4

    if dark:
        K.card(sl, x, y, w, h, K.DARK_G)
    elif style == "header":
        K.card(sl, x, y, w, h, "FFFFFF" if K.BG != "FFFFFF" else K.CELL_G)
        bar = 44
        if K.CORNER == "round":  # 上圆下方，贴合圆角卡片
            K._rshape(sl, x, y, w, bar, K.HEAD_G)
            K._rshape(sl, x, y + bar / 2, w, bar / 2, K.HEAD_G, radius=0)
        else:
            K._rshape(sl, x, y, w, bar, K.HEAD_G, radius=0)
        K.text(sl, x + pad, y, w - 2 * pad, bar, K.T(leaf.get("title", ""), 18, "FFFFFF", True), anchor="m")
        cy, ch = y + bar + 16, h - bar - 16 - pad + 4
    elif style == "label":
        K.card(sl, x, y, w, h, K.CELL_G)
        lw = min(190, w * 0.24)
        K._rshape(sl, x, y, lw, h, K.DARK_G)
        if K.CORNER == "round":
            K._rshape(sl, x + lw / 2, y, lw / 2, h, K.DARK_G, radius=0)
        size = 24
        while K.est_lines(leaf.get("title", ""), lw - 40, size) > 3 and size > 16:
            size -= 1
        K.text(sl, x + 20, y, lw - 40, h, K.T(leaf.get("title", ""), size, "FFFFFF", True), anchor="m", spacing=1.3)
        cx, cw = x + lw + 24, w - lw - 48
    elif style == "number":
        K.card(sl, x, y, w, h, "FFFFFF" if K.BG != "FFFFFF" else K.CELL_G)
        d = 30
        s = sl.shapes.add_shape(MSO_SHAPE.OVAL, K.E(x + 14), K.E(y + 14), K.E(d), K.E(d))
        K._paint(s, (K.MAIN, K.mix(K.MAIN, "FFFFFF", .25), 315))
        K.text(sl, x + 14, y + 14, d, d, K.T("%02d" % leaf.get("_num", 1), 13, "FFFFFF", True), align="c", anchor="m")
        cy, ch = y + 54, h - 54 - pad + 4
    elif carded:
        K.card(sl, x, y, w, h, K.CELL_G)

    if leaf.get("title") and style not in ("header", "label"):
        hh = K.heading(sl, cx, cy, cw, leaf["title"], dark=dark, size=19 if cw > 420 else 17)
        cy += hh
        ch -= hh
    draw_element(sl, (cx, cy, cw, ch), el, dark, ctx)


def draw_element(sl, box, el, dark, ctx):
    x, y, w, h = box
    t = el["type"]
    if t == "statement":
        size = el.get("size", 22)
        sub = el.get("sub")
        while K.est_height([(el["text"], size)], w, 1.4) > h * (0.62 if sub else 1.0) and size > 14:
            size -= 1
        fg = "FFFFFF" if dark else K.MAIN
        subfg = K.ONDARK if dark else K.SUB
        anc = el.get("anchor", "m")
        tb = K.text(sl, x, y, w, h, [[(el["text"], size, fg, True)]], anchor=anc, spacing=1.4)
        if sub:
            p2 = tb.text_frame.add_paragraph()
            p2.line_spacing = 1.45
            p2.space_before = Pt(14)
            r = p2.add_run(); r.text = sub
            f = r.font; f.size = Pt(max(13, size - 6) * 0.75); f.name = K.FONT
            f.color.rgb = K.rgb(subfg)
            rpr = r._r.get_or_add_rPr(); ea = rpr.makeelement(qn("a:ea"), {}); ea.set("typeface", K.FONT); rpr.append(ea)
    elif t == "kpi_row":
        K.kpi_row(sl, x, y, w, h, el["items"], dark=dark)
    elif t == "points":
        K.points(sl, x, y, w, h, el["items"], dark=dark, badge=el.get("badge", True))
    elif t == "bar_cards":
        K.bar_cards(sl, x, y + 2, w, h - 4, el["rows"], dark=dark)
    elif t == "table":
        size = el.get("size", 15)
        nrow = len(el["rows"]) + 1
        rh = h / nrow
        while rh < size * 2.05 and size > 11:
            size -= 0.5
        # 行高只要放得下文字：按最长单元格估算，最多到字号的 3.2 倍；表格贴格子顶部，不拉伸填满
        colw = el.get("colw") or [1] * len(el["header"])
        lines = max(K.est_lines(str(v), w * cw / sum(colw) - 24, size)
                    for r in [el["header"]] + el["rows"] for v, cw in zip(r, colw))
        rh = min(rh, max(size * 2.4, lines * size * 1.25 + 16), size * 3.2 * max(1, lines))
        K.table(sl, x, y, w, rh * nrow, el["header"], el["rows"], el.get("colw"), hl=el.get("hl"), size=size, hsize=size)
    elif t == "waterfall":
        K.chart_waterfall(sl, x, y + 4, w, h - 8, el["steps"], unit=el.get("unit", "万元"))
    elif t == "bar_h":
        # 条形图按条数取高度（每条约 44px），不拉满格子；格子多出来的地方由规格检查要求收格、放图
        h = min(h, len(el["cats"]) * K.BAR_ROW + 24)
        K.chart_bar_h(sl, x, y + 4, w, h - 8, el["cats"], el["vals"], suffix=el.get("suffix", ""),
                      decimals=el.get("decimals", 0), unit=el.get("unit", ""), hot=el.get("hot"),
                      maxv=el.get("maxv"), gap_width=el.get("gap_width"))
    elif t == "column":
        K.chart_column(sl, x, y + 4, w, h - 8, el["cats"], el["vals"], suffix=el.get("suffix", ""),
                       decimals=el.get("decimals", 0), unit=el.get("unit", ""), hot=el.get("hot"), maxv=el.get("maxv"))
    elif t == "line":
        K.chart_line(sl, x, y + 4, w, h - 8, el["cats"], el.get("series") or el["vals"], suffix=el.get("suffix", ""),
                     decimals=el.get("decimals", 0), hot=el.get("hot"), vmin=el.get("min"), vmax=el.get("max"))
    elif t == "donut":
        K.chart_donut(sl, x, y + 2, w, h - 4, el["cats"], el["vals"], colors=el.get("colors"), center=el.get("center"))
    elif t == "steps":
        K.steps(sl, x, y, w, h, el["items"], dark=dark)
    elif t == "chain":
        K.chain(sl, x, y, w, h, el["rows"], dark=dark)
    elif t == "compare":
        K.compare(sl, x, y, w, h, el["rows"], note=el.get("note"), dark=dark)
    elif t == "feature":
        K.feature(sl, x, y + 4, w, h - 4, el["icon"], el["title"], el["body"], dark=dark)
    elif t == "para":
        K.para(sl, x, y, w, h, el["body"], dark=dark, lead=el.get("lead"))
    elif t == "tags":
        K.tags(sl, x, y, w, h, el["items"], lead=el.get("lead"), dark=dark)
    elif t == "tip":
        K.tip(sl, x, y, w, h, el["text"], kind=el.get("kind", "warn"), dark=dark)
    elif t == "icon_steps":
        K.icon_steps(sl, x, y, w, h, el["items"], dark=dark)
    elif t == "num_pair":
        K.num_pair(sl, x, y, w, h, el["items"], note=el.get("note"), dark=dark)
    elif t == "split_note":
        nh = K.est_height([(el["note"], 14)], w, 1.45) + 26
        draw_element(sl, (x, y, w, h - nh), el["main"], dark, ctx)
        K.hline(sl, x, y + h - nh + 8, w, K.mix(K.MAIN, "FFFFFF", .35) if dark else K.LINE, 0.75)
        K.text(sl, x, y + h - nh + 20, w, nh - 20, K.T(el["note"], 14, K.ONDARK if dark else K.MUTED), spacing=1.45)
    elif t == "photo":
        draw_photo_cell(sl, box, el, ctx, "第 %s 页图片" % ctx.page)
    elif t == "stack":
        tot = sum(wt for _, wt in el["items"])
        cy = y
        for idx, (sub, wt) in enumerate(el["items"]):
            hh = h * wt / tot
            if idx == 0 and sub["type"] == "statement":
                sub = dict(sub, anchor="b")  # 结论句贴住下方图示，避免中间空出一大段
            draw_element(sl, (x, cy, w, hh), sub, dark, ctx)
            cy += hh
    else:
        raise ValueError("未知组件 " + t)


def _number_cards(leaves):
    """给 card=number 的格子按阅读顺序编号，并返回需要画箭头的相邻对。"""
    nums = [(lf, b) for lf, b in leaves if lf.get("card") == "number"]
    nums.sort(key=lambda it: (round(it[1][1] / 40), it[1][0]))
    pairs = []
    for i, (lf, b) in enumerate(nums):
        lf["_num"] = i + 1
        if i and abs(nums[i - 1][1][1] - b[1]) < 2 and abs(nums[i - 1][1][3] - b[3]) < 2:
            pairs.append((nums[i - 1][1], b))
    return pairs


def draw_page(prs, spec, ctx, page_no):
    ctx.page = page_no
    sl = K.add_slide(prs)
    K.title_block(sl, spec["title"], spec.get("footer", ""))
    LG.body(sl)
    K.page_number(sl, page_no)
    sep = spec.get("separation", "纯卡片")
    if "creative" in spec:
        cr = spec["creative"]
        if cr.get("kind") == "asset":  # 原版创意素材：出 PPTX 后由 assets.apply 放进本页
            ctx.asset_jobs.append(dict({k: v for k, v in cr.items() if k != "kind"}, slide=page_no))
            ctx.layout.append({"slide": page_no, "title": spec["title"], "has_photo": True, "cells": []})
            return sl
        draw_creative(sl, cr)
        return sl
    leaves, seams = [], []
    K.layout(spec["partition"], K.BODY_X, K.BODY_Y, K.BODY_W, K.BODY_H, GAP.get(sep, 0), leaves, seams)
    pairs = _number_cards(leaves)
    extents = []
    ctx.photo_miss = []
    for leaf, box in leaves:
        n0 = len(sl.shapes)
        draw_leaf(sl, box, leaf, sep, ctx)
        extents.append(_content_extent(list(sl.shapes)[n0:], box))
    miss = [tuple(b) for b in ctx.photo_miss]

    def cell_type(lf, box):
        t = lf["element"]["type"]
        return "空图片格" if t == "photo" and tuple(box) in miss else t
    ctx.layout.append({"slide": page_no, "title": spec["title"],
                       "has_photo": any(cell_type(lf, b) == "photo" for lf, b in leaves),
                       "cells": [{"box": [round(v, 1) for v in box], "name": lf.get("title") or lf["element"]["type"],
                                  "type": cell_type(lf, box)} for lf, box in leaves]})
    if sep in LINE_MODES:
        draw_dividers(sl, leaves, extents, seams, ctx)
    for a, b in pairs:  # 编号卡之间的箭头
        gx = a[0] + a[2]
        gw = b[0] - gx
        if gw >= 12:
            K._shape(sl, MSO_SHAPE.CHEVRON, gx + gw / 2 - 5, a[1] + a[3] / 2 - 8, 10, 16, K.mix(K.MAIN, "FFFFFF", .45))
    return sl



# ---------------- 分隔线 ----------------
LINE_GAP = 20  # 分隔线之间、分隔线与组件内细线之间至少留的空隙（px）


def _text_extent(sh):
    """文本框里文字实际占的范围（文本框常比文字大，按估算字数行数收窄）。"""
    x, y, w, h = sh.left / K.PX, sh.top / K.PX, sh.width / K.PX, sh.height / K.PX
    tf = sh.text_frame
    tot_h, max_w = 0.0, 0.0
    for p in tf.paragraphs:
        runs = [r for r in p.runs if r.text]
        if not runs:
            continue
        size = max((r.font.size.pt / 0.75) if r.font.size else 16 for r in runs)
        txt = "".join(r.text for r in runs)
        lines = K.est_lines(txt, w, size)
        tot_h += lines * size * 1.22 * (p.line_spacing if isinstance(p.line_spacing, float) else 1.2)
        max_w = max(max_w, w if lines > 1 else min(w, K._units(txt) * size if hasattr(K, "_units") else w))
    if not tot_h:
        return None
    tot_h = min(tot_h, h)
    anc = tf.vertical_anchor
    from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
    ty = y + (h - tot_h) / 2 if anc == MSO_ANCHOR.MIDDLE else (y + h - tot_h if anc == MSO_ANCHOR.BOTTOM else y)
    al = tf.paragraphs[0].alignment
    tx = x + (w - max_w) / 2 if al == PP_ALIGN.CENTER else (x + w - max_w if al == PP_ALIGN.RIGHT else x)
    return tx, ty, tx + max_w, ty + tot_h


def _content_extent(shapes, box):
    """一格里组件实际占的范围 (x0, y0, x1, y1)；卡片底、整格色块不算内容。"""
    bx, by, bw, bh = box
    xs = []
    for sh in shapes:
        if sh.width is None:
            continue
        x, y, w, h = sh.left / K.PX, sh.top / K.PX, sh.width / K.PX, sh.height / K.PX
        if w >= bw - 2 and h >= bh - 2:  # 整格卡片底
            continue
        if sh.has_text_frame and sh.text_frame.text.strip() and sh.shape_type == 17:  # 文本框
            e = _text_extent(sh)
            if e:
                xs.append(e)
            continue
        xs.append((x, y, x + w, y + h))
    if not xs:
        return None
    return (min(e[0] for e in xs), min(e[1] for e in xs), max(e[2] for e in xs), max(e[3] for e in xs))


def draw_dividers(sl, leaves, extents, seams, ctx):
    """纯线条、线卡混用页的分隔线：
    - 不沿着卡片边描线；两端各缩进 INSET，长度不超过两侧内容中较长的一侧；
    - 分隔线之间、分隔线与组件内的细线之间不相碰，至少留 LINE_GAP；
    - 线被内容截得很短时提醒：旁边的格子留空大，收格或加挨角的图片格。"""
    cards = [b for lf, b in leaves if lf.get("fill") == "重点" or lf.get("card") in ("dark", "header", "number", "label")]

    def on_card(sm):
        for bx, by, bw, bh in cards:
            if sm[0] == "v" and min(abs(sm[1] - bx), abs(sm[1] - bx - bw)) < 1.5 and sm[2] < by + bh and sm[2] + sm[3] > by:
                return True
            if sm[0] == "h" and min(abs(sm[1] - by), abs(sm[1] - by - bh)) < 1.5 and sm[2] < bx + bw and sm[2] + sm[3] > bx:
                return True
        return False

    photos = [b for lf, b in leaves if (lf.get("element") or {}).get("type") == "photo"]

    def on_photo(sm):  # 整条都贴着图片格的那一刀不画线，图片本身就是分隔；只贴一段的，画线时把那段去掉
        for bx, by, bw, bh in photos:
            if sm[0] == "v" and min(abs(sm[1] - bx), abs(sm[1] - bx - bw)) < 1.5 and sm[2] >= by - 2 and sm[2] + sm[3] <= by + bh + 2:
                return True
            if sm[0] == "h" and min(abs(sm[1] - by), abs(sm[1] - by - bh)) < 1.5 and sm[2] >= bx - 2 and sm[2] + sm[3] <= bx + bw + 2:
                return True
        return False

    lines = []
    for kind, pos, start, length in [sm for sm in seams if not on_card(sm) and not on_photo(sm)]:
        lo, hi = start + INSET, start + length - INSET
        spans = []
        for (lf, (x, y, w, h)), ext in zip(leaves, extents):
            if not ext or (lf.get("element") or {}).get("type") == "photo":
                continue
            if kind == "v" and (abs(x + w - pos) < 2 or abs(x - pos) < 2) and y < hi and y + h > lo:
                spans.append((ext[1], ext[3]))
            if kind == "h" and (abs(y + h - pos) < 2 or abs(y - pos) < 2) and x < hi and x + w > lo:
                spans.append((ext[0], ext[2]))
        if spans:  # 不比两侧内容中较长的一侧更长
            lo, hi = max(lo, min(s[0] for s in spans)), min(hi, max(s[1] for s in spans))
        for bx, by, bw, bh in photos:  # 不沿着图片边画线：线在图片旁边的那一段去掉
            if kind == "v" and min(abs(pos - bx), abs(pos - bx - bw)) < 2:
                a0, a1 = by, by + bh
            elif kind == "h" and min(abs(pos - by), abs(pos - by - bh)) < 2:
                a0, a1 = bx, bx + bw
            else:
                continue
            if a0 <= lo + 1 and a1 > lo:
                lo = a1 + LINE_GAP
            if a1 >= hi - 1 and a0 < hi:
                hi = a0 - LINE_GAP
        lines.append([kind, pos, lo, hi, length - 2 * INSET])
    # 组件里已有的细线（如表格、因果链行线）也算，避免相碰
    inner = []
    for sh in sl.shapes:
        if sh.shape_type == 9 or (sh.width is not None and (sh.width < 3 * K.PX or sh.height < 3 * K.PX)):
            x, y, w, h = sh.left / K.PX, sh.top / K.PX, sh.width / K.PX, sh.height / K.PX
            inner.append(["v", x, y, y + h] if w < h else ["h", y, x, x + w])
    for ln in lines:  # 和其他分隔线、组件细线垂直相遇时，缩回 LINE_GAP
        for other in [o for o in lines if o is not ln] + inner:
            if other[0] == ln[0]:
                continue
            if other[2] - LINE_GAP <= ln[1] <= other[3] + LINE_GAP:
                if abs(ln[2] - other[1]) < LINE_GAP or (ln[2] < other[1] < ln[2] + LINE_GAP):
                    ln[2] = other[1] + LINE_GAP
                if abs(ln[3] - other[1]) < LINE_GAP or (ln[3] - LINE_GAP < other[1] < ln[3]):
                    ln[3] = other[1] - LINE_GAP
    for kind, pos, lo, hi, full in lines:
        if hi - lo < 40:
            continue
        if kind == "v":
            K.vline(sl, pos, lo, hi - lo, K.LINE, 0.75)
        else:
            K.hline(sl, lo, pos, hi - lo, K.LINE, 0.75)
        if full > 0 and (hi - lo) < 0.6 * full:
            print("提醒：第 %s 页分隔线按两侧内容截到原长的 %d%%，旁边格子留空较大：收格，或加一个挨角的图片格"
                  % (ctx.page, round((hi - lo) / full * 100)))


# ---------------- 创意型 ----------------
def draw_creative(sl, cr):
    kind = cr.get("kind", "timeline")
    if kind == "timeline":
        _timeline(sl, cr)
    elif kind == "radial":
        _radial(sl, cr)
    else:
        raise ValueError("未知创意版式 " + kind)


def _timeline(sl, cr):
    """时间脊柱：一句结论 + 可选大数字；横轴上 3–6 个节点卡片上下交错。"""
    st = cr.get("statement", "")
    size = 26
    while K.est_lines(st, 720, size) > 1 and size > 18:
        size -= 1
    K.text(sl, 48, 126, 720, 44, K.T(st, size, K.MAIN, True), anchor="m", spacing=1.25)
    if cr.get("big"):
        n, unit, lab = cr["big"]
        K.text(sl, 800, 104, 432, 70, [[(n, 60, K.MAIN, True), (" " + unit, 19, K.MAIN, True)]], anchor="b", spacing=1.0)
        K.text(sl, 804, 178, 424, 24, K.T(lab, 14, K.SUB), anchor="t", spacing=1.2)
    nodes = cr["nodes"]
    n = len(nodes)
    axis_y = 396
    x0, x1 = 120, 1160
    K.hline(sl, 60, axis_y, 1160, K.TRACK, 3)
    cw = min(280, (x1 - x0) / max(n - 1, 1) * 1.7)
    chh = 150
    any_hot = any(nd.get("hot") for nd in nodes)  # 有重点节点时，普通节点退成灰
    for i, nd in enumerate(nodes):
        x = x0 + (x1 - x0) * (i / max(n - 1, 1))
        up = (i % 2 == 0)
        col = K.MAIN if (nd.get("hot") or not any_hot) else K.NORM
        d = 34
        s = sl.shapes.add_shape(MSO_SHAPE.OVAL, K.E(x - d / 2), K.E(axis_y - d / 2), K.E(d), K.E(d))
        K._paint(s, (col, K.mix(col, "FFFFFF", .3), 315))
        K.text(sl, x - d / 2, axis_y - d / 2, d, d, K.T(nd.get("tag", str(i + 1)), 12, "FFFFFF", True), align="c", anchor="m")
        cy = axis_y - 30 - chh if up else axis_y + 30
        K.vline(sl, x, (cy + chh) if up else (axis_y + d / 2), 30 - d / 2, col, 1.25)
        cxp = min(max(x - cw / 2, 48), 1232 - cw)
        K.card(sl, cxp, cy, cw, chh, K.CELL_G)
        K.rect(sl, cxp, cy, cw, 5, (col, K.mix(col, "FFFFFF", .3), 0))
        K.text(sl, cxp + 18, cy + 14, cw - 36, 24, [[(nd["label"], 15, col, True), ("  " + nd.get("sub", ""), 12, K.MUTED, False)]], anchor="m")
        if nd.get("value"):
            K.text(sl, cxp + 18, cy + 40, cw - 36, 40, [[(nd["value"], 28, K.MAIN, True), (" " + nd.get("unit", ""), 12, K.SUB, False)]], anchor="m")
            ny = cy + 84
        else:
            ny = cy + 44
        K.text(sl, cxp + 18, ny, cw - 36, cy + chh - ny - 10, K.T(nd.get("note", ""), 13, K.SUB), anchor="t", spacing=1.3)
    if cr.get("foot"):
        K.text(sl, 48, 640, 1184, 24, K.T(cr["foot"], 13, K.SUB), anchor="m", spacing=1.2)


def _radial(sl, cr):
    """中心辐射：中心圆写主题，周围 4–6 个卫星卡片（图标 + 标题 + 说明）。"""
    import math
    import icons as I
    cxc, cyc = 640, 398
    R = 118
    s = sl.shapes.add_shape(MSO_SHAPE.OVAL, K.E(cxc - R), K.E(cyc - R), K.E(2 * R), K.E(2 * R))
    K._paint(s, K.DARK_G)
    K.text(sl, cxc - R + 20, cyc - R, 2 * R - 40, 2 * R, K.T(cr["center"], 22, "FFFFFF", True), align="c", anchor="m", spacing=1.3)
    sats = cr["items"]
    n = len(sats)
    cw, chh = 330, 112
    for i, it in enumerate(sats):
        ang = -math.pi / 2 + 2 * math.pi * i / n
        px = cxc + math.cos(ang) * 400
        py = cyc + math.sin(ang) * 205
        bx = min(max(px - cw / 2, 48), 1232 - cw)
        by = min(max(py - chh / 2, 120), 674 - chh)
        lx, ly = cxc + math.cos(ang) * (R + 8), cyc + math.sin(ang) * (R + 8)
        ln = sl.shapes.add_connector(1, K.E(lx), K.E(ly), K.E(bx + cw / 2), K.E(by + chh / 2))
        ln.line.color.rgb = K.rgb(K.mix(K.MAIN, "FFFFFF", .55)); ln.line.width = Pt(1); K.flat(ln)
        K.card(sl, bx, by, cw, chh, "FFFFFF" if K.BG != "FFFFFF" else K.CELL_G)
        I.icon(sl, it.get("icon"), bx + 18, by + 20, 28, K.MAIN)
        K.text(sl, bx + 58, by + 14, cw - 74, 30, K.T(it["title"], 17, K.INK, True), anchor="m")
        K.text(sl, bx + 58, by + 46, cw - 74, chh - 54, K.T(it.get("note", ""), 13.5, K.SUB), anchor="t", spacing=1.35)
    if cr.get("statement"):
        K.text(sl, 48, 120, 1184, 30, K.T(cr["statement"], 18, K.SUB), anchor="m")


# ---------------- 整稿 ----------------
def pick_styles(st):
    """首页、目录、转场版式：由脚本按封面标题随机抽一种，一套之内统一，同一份稿子重跑不变。
    只有用户明确指定了版式（settings.styles_from 为 "user"）时才用 settings.styles；
    模型自己写的 styles 一律忽略，避免每套都选同一组、不同公司的稿子长得一样。"""
    import hashlib
    import random
    given = st.get("styles") or {}
    if given and st.get("styles_from") != "user":
        print("结构页版式：settings.styles 未标 styles_from: user，按规则忽略，改为随机抽取"
              "（只有用户明确要求某种版式时才写 styles，并加 \"styles_from\": \"user\"）")
        given = {}
    rng = random.Random(int(hashlib.md5((st.get("title") or "").encode("utf-8")).hexdigest(), 16))
    out, drawn = {}, []
    for key, allowed in (("cover", S.COVER_STYLES), ("agenda", S.AGENDA_STYLES), ("transition", S.TRANSITION_STYLES)):
        pick = rng.choice(allowed)  # 先抽再看是否指定，保证每一项的抽签结果不受其他项影响
        out[key] = given.get(key) or pick
        if not given.get(key):
            drawn.append("%s %s" % ({"cover": "首页", "agenda": "目录", "transition": "转场"}[key], pick))
    if drawn:
        print("结构页版式：%s（未指定，按标题随机抽取，套内统一；要固定或更换就写进 settings.styles）" % "、".join(drawn))
    return out


BLANK_FAIL = []  # 渲染后留白检查必须处理的页；build() 仍产出文件，main() 据此返回 3


def build(deck, out_dir, evidence_text=None, render=True, force=False):
    BLANK_FAIL[:] = []
    os.makedirs(out_dir, exist_ok=True)
    errors, warnings = check_spec.check_deck(deck, evidence_text)
    for w in warnings:
        print("提醒：" + w)
    for e in errors:
        print("错误：" + e)
    if errors and not force:
        print("规格检查未通过，共 %d 个错误。修改 deck.json 后重跑。" % len(errors))
        return None
    st = deck.get("settings", {})
    pal = choose_palette(st)
    print("配色：%s（主色 %s，%s 组合，%s）" % (P.label(pal["preset"]), pal["main"], pal["scheme"], "圆角" if pal["corner"] == "round" else "直角"))
    ph = st.get("photos", {})
    ctx = Ctx(out_dir, PH.online_default(ph), ph.get("sources"))
    rec = os.path.join(out_dir, "图片来源.json")
    if os.path.exists(rec):  # 每次生成重写来源记录，只记本稿实际用到的图；下载过的图留在缓存里复用
        os.remove(rec)
    if ctx.online:
        PH.budget_start()
        n, sec = PH.prefetch(PH.deck_topics(deck), out_dir, ctx.sources)
        if n:
            print("预取图片：新下载 %d 张，用时 %.0f 秒（其余命中缓存）" % (n, sec))
    prs = K.new_deck()
    styles = pick_styles(st)
    if st.get("logo") and not LG.use(st["logo"], out_dir):
        print("提醒：settings.logo「%s」不是可用的图片文件，本稿不放 LOGO" % st["logo"])
    elif not st.get("logo"):
        LG.use(None, out_dir)
    cover_photo = photo(ctx, ph["cover"], "封面") if ph.get("cover") else None
    cs = styles.get("cover", "band")
    meta = S._meta(st.get("org", ""), st.get("date", ""), st.get("reporter"))
    if not S.cover(prs, cs, st.get("title", ""), st.get("subtitle", ""), meta, cover_photo):
        if cs != "band":
            print("提醒：封面版式 %s 需要照片，没有取到，已用原版式" % cs)
        K.cover(prs, st.get("title", ""), st.get("subtitle", ""), st.get("org", ""), st.get("date", ""), cover_photo)
    LG.corner(prs.slides[-1])
    sections = deck.get("sections", [])
    if sections:
        ag = styles.get("agenda", "panel")
        ag_photo = (photo(ctx, ph["agenda"], "目录") if ph.get("agenda") else cover_photo) if ag != "panel" else None
        if not S.agenda(prs, ag, sections, ag_photo):
            K.agenda(prs, sections)
        LG.corner(prs.slides[-1])
    sec_index = {s["name"]: i + 1 for i, s in enumerate(sections)}
    current = None
    page_no = len(prs.slides._sldIdLst)
    for spec in deck["pages"]:
        sec = spec.get("section")
        if sec and sec != current and sec in sec_index:
            info = sections[sec_index[sec] - 1]
            tp = info.get("photo")
            tphoto = photo(ctx, tp, "转场「%s」" % sec) if tp else None
            if not tphoto:
                # 转场页一律带图：本章没取到图时，先按章名再按通用办公/城市图补，最后借用封面图；不退回纯色底
                tphoto = (get_photo(ctx, sec) or get_photo(ctx, "办公、会议、城市") or cover_photo)
                if tphoto:
                    print("提醒：转场「%s」没有取到指定主题的图，已用 %s 补上" % (sec, os.path.basename(tphoto)))
            if not S.transition(prs, styles.get("transition", "half"), sec_index[sec], sec, info.get("desc", ""), tphoto):
                K.transition(prs, sec_index[sec], sec, info.get("desc", ""), tphoto)
            LG.corner(prs.slides[-1])
            current = sec
        page_no = len(prs.slides._sldIdLst) + 1
        spec.setdefault("footer", st.get("footer", ""))
        draw_page(prs, spec, ctx, page_no)
    cl = deck.get("closing", {})
    reporter = st.get("reporter")
    sub = cl.get("sub") or (("汇报人：" + reporter) if reporter else st.get("org", ""))
    words = cl.get("words", "感谢聆听")
    if not S.cover(prs, cs, words, sub, "", cover_photo):  # 结束页与首页同一版式、同一张图
        K.closing(prs, words, sub, cover_photo)
    LG.corner(prs.slides[-1])
    if LG.LAST:
        names = {"original": "原色", "white": "反白", "dark": "深色剪影"}
        cnt = {}
        for where, _, how in LG.LAST:
            cnt.setdefault(where, {}).setdefault(names[how], 0)
            cnt[where][names[how]] += 1
        print("LOGO 放置：" + "；".join("%s %s" % (w, "、".join("%s %d 页" % kv for kv in d.items())) for w, d in cnt.items()))
    name = st.get("filename") or st.get("title") or "deck"
    out = os.path.join(out_dir, name + ".pptx")
    prs.save(out)
    if ctx.asset_jobs:
        import assets
        for r in assets.apply(out, assets.palette_for(K), ctx.asset_jobs):
            for n in r["notes"]:
                print("提醒：第 %d 页 %s：%s" % (r["slide"], r["code"], n.replace("[创意] ", "")))
            for e in r["errors"]:
                print("错误：第 %d 页 %s 未放上：%s" % (r["slide"], r["code"], e))
    try:
        import ledger
        ledger.append(deck, pal)
    except OSError:
        pass
    if st.get("public_output"):
        import public_output
        public_output.sanitize_pptx(out)
    print("已生成", out, "共", len(prs.slides._sldIdLst), "页")
    tot = ctx.n_lib + ctx.n_web
    if tot:
        share = ctx.n_lib / tot
        print("图片来源：联网 %d 张，自带图库 %d 张（图库占 %.0f%%）" % (ctx.n_web, ctx.n_lib, share * 100))
        if share > PH.LIB_MAX_SHARE:
            print("提醒：自带图库图片占 %.0f%%，超过 %.0f%% 上限%s" % (share * 100, PH.LIB_MAX_SHARE * 100,
                  "（联网不可用或没取到合适的图）" if not ctx.online or PH._NET["down"] else "，把 topic 改写成中英文关键词让它联网取图"))
    if os.path.exists(os.path.join(out_dir, "图片来源.json")):
        print("版权提醒：" + PH.COPYRIGHT_NOTE)
    if render:
        try:
            import render as R
            pngs = R.render(out, out_dir)
            rep = blank_check(ctx.layout, pngs, out_dir, name)
            BLANK_FAIL[:] = [it["页"] for it in rep if it.get("必须")]
        except Exception as ex:  # 渲染依赖 LibreOffice，缺失时只提示
            print("提醒：预览渲染失败（%s），PPTX 已生成。" % ex)
    return out


def blank_check(layout, pngs, out_dir, name):
    """渲染后量每格的实际空块，打印提醒并写入 <名>-留白检查.json。"""
    import check_blank
    pages = []
    for pg in layout:
        if 0 < pg["slide"] <= len(pngs):
            pages.append(dict(pg, png=os.path.basename(pngs[pg["slide"] - 1])))
    json.dump({"pages": pages}, open(os.path.join(out_dir, name + "-layout.json"), "w", encoding="utf-8"),
              ensure_ascii=False, indent=1)
    rep = check_blank.check({"pages": pages}, out_dir)
    check_blank.print_report(rep)
    json.dump(rep, open(os.path.join(out_dir, name + "-留白检查.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    return rep


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("deck")
    ap.add_argument("--out", default="out")
    ap.add_argument("--evidence")
    ap.add_argument("--no-render", action="store_true")
    ap.add_argument("--force", action="store_true")
    a = ap.parse_args(argv)
    deck = json.load(open(a.deck, encoding="utf-8"))
    ev = open(a.evidence, encoding="utf-8").read() if a.evidence else None
    out = build(deck, a.out, ev, render=not a.no_render, force=a.force)
    if not out:
        return 1
    return 3 if BLANK_FAIL else 0


if __name__ == "__main__":
    sys.exit(main())

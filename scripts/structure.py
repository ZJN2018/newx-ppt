# -*- coding: utf-8 -*-
"""整稿结构页的多种版式：首页（及同款结束页）、目录页、章节转场页。

版式只借参考稿的构图（照片位置、渐隐方向、文字区），颜色跟随整套主色，照片先调色再铺。
在 deck.json 的 settings.styles 里选：
    {"cover": "fade" | "center" | "bottom" | "band",
     "agenda": "photo_left" | "photo_right" | "columns" | "panel",
     "transition": "full" | "light" | "split" | "band" | "half"}
不写时沿用原版式（band / panel / half）。需要照片的版式没取到照片时，自动退回原版式。
"""
import deckkit as K
import photos as PH

W, H = 1280, 720
COVER_STYLES = ("band", "fade", "center", "bottom")
AGENDA_STYLES = ("panel", "photo_left", "photo_right", "columns")
TRANSITION_STYLES = ("half", "full", "light", "split", "band")
# 中文说明，供文档与提示使用
NAMES = {
    "cover": {"band": "原版：整幅照片＋主色蒙版", "fade": "左字右图，照片向文字区渐隐", "center": "整幅照片，中部浅色带，文字块垂直居中",
              "bottom": "上图下字"},
    "agenda": {"panel": "原版：左侧主色块＋列表", "photo_left": "左图右列表", "photo_right": "左列表右图", "columns": "整幅照片，编号横排"},
    "transition": {"half": "原版：左色块＋右半幅照片", "full": "整幅照片，主色压暗，编号章名靠左", "light": "整幅照片，浅色渐隐，主色编号",
                   "split": "左字右图，大号编号", "band": "上图下字条"},
}


# ---------------- 小工具 ----------------
# 封面、结束页的照片基本保留原色，只极轻地向主色靠：饱和度压得太低会发灰、看着别扭
COVER_TONE = {"sat": 0.88, "duo": 0.10}

def _photo(sl, path, x, y, w, h, sat=0.55, duo=0.25, blurred=False):
    """整幅照片：保留较多原色，只轻度向主色调色，避免与主色冲突。
    贴页边的一侧多铺 1px 出血，避免渲染时页边露出一条细缝。"""
    if x <= 0:
        x, w = x - 1, w + 1
    if x + w >= W:
        w += 1
    if y <= 0:
        y, h = y - 1, h + 1
    if y + h >= H:
        h += 1
    PH.place(sl, path, x, y, w, h, tone_main=K.MAIN, sat=sat, duo=duo, blurred=blurred)


def _fade(sl, x, y, w, h, color, a0, a1, ang):
    """渐隐层：从 a0 到 a1 的不透明度。ang：0 = 左→右，90 = 上→下。
    （photos.overlay 的 90° 是下→上，这里换过来，调用处统一按“上→下”写。）"""
    if ang == 90:
        a0, a1 = a1, a0
    PH.overlay(sl, x, y, w, h, color, a0, a1, ang)


def _vfade(sl, x, y, w, h, color, stops):
    """上→下多档渐隐：stops = [(位置0~1, 不透明度), ...]，一层形状，没有接缝。"""
    from lxml import etree
    from pptx.oxml.ns import qn
    s = PH.overlay(sl, x, y, w, h, color, 0.0, 0.0, 90)
    gs = s.fill._xPr.find(qn("a:gradFill")).find(qn("a:gsLst"))
    tpl = etree.fromstring(etree.tostring(gs[0]))
    for g in list(gs):
        gs.remove(g)
    # overlay 的 90° 是下→上，位置倒过来写
    for pos, a in sorted(((1 - p, a) for p, a in stops)):
        g = etree.fromstring(etree.tostring(tpl))
        g.set("pos", str(int(round(pos * 100000))))
        g.find(qn("a:srgbClr")).find(qn("a:alpha")).set("val", str(int(a * 100000)))
        gs.append(g)
    return s


def _center_block(sl, title, subtitle, meta, color, sub_color, size=50, cy=H * 0.47, x=140, w=1000, pill=True,
                  meta_color=None, rule=False):
    """标题、副题、信息条作为一个整体，按视觉中心（略高于几何中心）垂直居中。返回 (上沿, 下沿)。"""
    lines, size = K.break_title(title, w, size, 30)
    th = len(lines) * size * 1.22
    parts = [th]
    if subtitle:
        parts += [16, 30]
    if rule:
        parts += [14, 2]
    if meta:
        parts += [30, 34]
    top = cy - sum(parts) / 2
    y = top
    K.text(sl, x, y, w, th, [[(l, size, color, True)] for l in lines], anchor="m", align="c", spacing=1.15)
    y += th
    if subtitle:
        y += 16
        K.text(sl, x, y, w, 30, K.T(subtitle, 20, sub_color), align="c", anchor="m")
        y += 30
    if rule:
        y += 14
        K.hline(sl, W / 2 - 60, y, 120, K.MAIN, 1.5)
        y += 2
    if meta:
        y += 30
        if pill:
            pw = min(900, K._units(meta) * 15 + 60)
            K._rshape(sl, (W - pw) / 2, y, pw, 34, (K.MAIN, K.mix(K.MAIN, "FFFFFF", .25), 0), radius=17)
            K.text(sl, (W - pw) / 2, y, pw, 34, K.T(meta, 14, "FFFFFF", True), align="c", anchor="m")
        else:
            K.text(sl, x, y, w, 34, K.T(meta, 14, meta_color or sub_color), align="c", anchor="m")
        y += 34
    return top, y


def _accent(sl, x, y, w=44):
    K.rect(sl, x, y, w, 5, (K.MAIN, K.mix(K.MAIN, "FFFFFF", .3), 0))


def _title(sl, x, y, w, h, title, size, color, anchor="b", align="l", lo=30):
    lines, size = K.break_title(title, w, size, lo)
    K.text(sl, x, y, w, h, [[(l, size, color, True)] for l in lines], anchor=anchor, align=align, spacing=1.15)
    return len(lines), size


def _meta(org, date, reporter=None):
    parts = []
    if reporter:
        parts.append("汇报人：" + reporter)
    parts += [t for t in (org, date) if t]
    return "｜".join(parts)


def _new(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    K.rect(sl, 0, 0, W, H, K.BG)
    return sl


# ---------------- 首页（结束页同款） ----------------
def cover(prs, style, title, subtitle="", meta="", photo=None):
    """style 见 COVER_STYLES；结束页调用时 title 为“感谢聆听”，subtitle 为汇报人。"""
    if style not in COVER_STYLES or style == "band" or not photo:
        return None
    sl = _new(prs)
    if style == "fade":
        # 照片铺右侧 64%，向左渐隐到页底色；标题在左
        # 照片从 30% 处开始，渐隐带放宽到 44% 并分多档，过渡平缓、不出现断崖，照片露出得更多
        # 照片铺右侧 70%，左侧 60% 宽度内渐隐到页底色（渐隐做进图片，蒙版下看不到照片边）
        PH.place_faded(sl, photo, W * 0.30, -1, W * 0.70 + 1, H + 2, K.BG,
                       [(0.0, 1.0), (0.12, 0.95), (0.35, 0.6), (0.6, 0.0)], side="left", tone_main=K.MAIN, **PH.cover_tone(photo, K.MAIN))
        _accent(sl, 80, 228)
        _title(sl, 80, 248, 640, 150, title, 50, K.MAIN, anchor="t")
        if subtitle:
            K.text(sl, 80, 404, 560, 60, K.T(subtitle, 20, K.SUB), anchor="t", spacing=1.4)
        if meta:
            K.text(sl, 80, 620, 640, 30, K.T(meta, 14, K.MUTED), anchor="m")
    elif style == "bottom":
        # 照片占上方 62%，下方留白写标题
        ph = H * 0.62
        K.rect(sl, 0, 0, W, H, K.BG)
        PH.place_faded(sl, photo, -1, -1, W + 2, ph + 1, K.BG, [(0.0, 1.0), (0.3, 0.0)], side="bottom",
                       tone_main=K.MAIN, **PH.cover_tone(photo, K.MAIN))
        _accent(sl, 80, ph + 10)
        n, _ = _title(sl, 80, ph + 26, 1100, 110, title, 44, K.MAIN, anchor="t")
        if subtitle:
            K.text(sl, 80, ph + 30 + n * 52, 1000, 34, K.T(subtitle, 19, K.SUB), anchor="t")
        if meta:
            K.text(sl, 80, H - 46, 1000, 26, K.T(meta, 13, K.MUTED), anchor="m")
    elif style == "center":
        # 整幅照片，中部一条浅色带托住文字，上下渐隐露出照片；文字块整体按视觉中心垂直居中
        _photo(sl, photo, 0, 0, W, H, blurred=True, **PH.cover_tone(photo, K.MAIN))
        _vfade(sl, 0, 0, W, H, K.BG, [(0.0, 0.55), (0.24, 0.93), (0.66, 0.93), (1.0, 0.25)])
        _center_block(sl, title, subtitle, meta, K.MAIN, K.SUB)
    return sl


# ---------------- 目录 ----------------
def _items(sections):
    out = []
    for sec in sections:
        if isinstance(sec, dict):
            out.append((sec.get("name", ""), sec.get("desc", "")))
        elif isinstance(sec, (list, tuple)):
            out.append((sec[0], sec[1] if len(sec) > 1 else ""))
        else:
            out.append((sec, ""))
    return out


def agenda(prs, style, sections, photo=None):
    if style not in AGENDA_STYLES or style == "panel" or not photo:
        return None
    items = _items(sections)
    n = len(items)
    sl = _new(prs)
    if style in ("photo_left", "photo_right"):
        # 一侧一栏照片，另一侧编号列表，条目之间细分隔线
        pw = 420
        left = style == "photo_left"
        px = 0 if left else W - pw
        PH.place_faded(sl, photo, px - (1 if left else 0), -1, pw + 1, H + 2, K.BG, [(0.0, 1.0), (0.3, 0.0)],
                       side="right" if left else "left", tone_main=K.MAIN, **dict(zip(("sat", "duo"), PH.tone_for(photo))))
        lx = 510 if left else 96
        lw = 680 if left else 680
        K.text(sl, lx, 92, 200, 70, K.T("目录", 44, K.MAIN, True), anchor="m")
        K.text(sl, lx + 120, 108, 300, 40, K.T("CONTENTS", 15, K.mix(K.MAIN, "FFFFFF", .45), True), anchor="m")
        rh = min(84, 460 / max(n, 1))
        y0 = 190 + (460 - rh * n) / 2
        for i, (name, desc) in enumerate(items):
            y = y0 + i * rh
            K.text(sl, lx, y, 80, rh, K.T("%02d" % (i + 1), 30, K.MAIN, True), anchor="m")
            if desc:
                K.text(sl, lx + 90, y + 8, lw - 100, rh / 2 - 4, K.T(name, 22, K.INK, True), anchor="b")
                K.text(sl, lx + 90, y + rh / 2 + 4, lw - 100, rh / 2 - 10, K.T(desc, 14, K.MUTED), anchor="t")
            else:
                K.text(sl, lx + 90, y, lw - 100, rh, K.T(name, 22, K.INK, True), anchor="m")
            if i < n - 1:
                K.hline(sl, lx, y + rh, lw, K.LINE, 0.75)
    elif style == "columns":
        # 整幅照片，上部浅色渐隐，居中“目录”，编号横排
        _photo(sl, photo, 0, 0, W, H, blurred=True)
        # 整页一层多档浅色蒙版（比原来重一层），上部托住文字、下部照片退后；铺满整页，不留接缝
        _vfade(sl, 0, 0, W, H, K.BG, [(0.0, 0.97), (0.45, 0.88), (0.75, 0.70), (1.0, 0.55)])
        K.text(sl, 0, 130, W, 80, K.T("目录", 52, K.MAIN, True), align="c", anchor="m")
        cw = min(250, 1140 / max(n, 1))
        x0 = (W - cw * n) / 2
        for i, (name, desc) in enumerate(items):
            x = x0 + i * cw
            K.text(sl, x, 262, cw, 60, K.T("%02d" % (i + 1), 44, K.MAIN, True), align="c", anchor="m")
            K.text(sl, x + 12, 330, cw - 24, 64, K.T(name, 21, K.INK, True), align="c", anchor="t", spacing=1.25)
            if desc:
                K.text(sl, x + 16, 392, cw - 32, 50, K.T(desc, 13.5, K.MUTED), align="c", anchor="t", spacing=1.3)
            if i < n - 1:
                K.vline(sl, x + cw, 276, 150, K.mix(K.MAIN, "FFFFFF", .6), 0.75)
    return sl


# ---------------- 章节转场 ----------------
def transition(prs, style, idx, name, desc="", photo=None):
    if style not in TRANSITION_STYLES or style == "half" or not photo:
        return None
    sl = _new(prs)
    num = "%02d" % idx
    if style == "full":
        # 整幅照片，从左往右压主色，白色大号章节号与章名
        _photo(sl, photo, 0, 0, W, H, blurred=True)
        # 蒙版铺满整页、右端仍保留约 40%，白字和照片都压得住（原来只铺 75% 宽且渐到 0，偏浅）
        _fade(sl, 0, 0, W, H, K.mix(K.MAIN, "000000", .15), 0.94, 0.40, 0)
        K.text(sl, 90, 170, 600, 160, K.T(num, 120, "FFFFFF", True), anchor="b")
        _title(sl, 94, 340, 620, 120, name, 44, "FFFFFF", anchor="t", lo=28)
        if desc:
            K.text(sl, 96, 470, 560, 70, K.T(desc, 17, K.ONDARK), anchor="t", spacing=1.4)
    elif style == "light":
        # 整幅照片，左侧浅色渐隐，主色大号章节号；照片右侧露出，不模糊
        _photo(sl, photo, 0, 0, W, H)
        _fade(sl, 0, 0, W * 0.72, H, K.BG, 0.98, 0.0, 0)
        K.text(sl, 90, 150, 600, 160, K.T(num, 120, K.MAIN, True), anchor="b")
        _accent(sl, 96, 330, 56)
        _title(sl, 94, 352, 620, 120, name, 42, K.INK, anchor="t", lo=28)
        if desc:
            K.text(sl, 96, 470, 540, 70, K.T(desc, 16, K.SUB), anchor="t", spacing=1.4)
    elif style == "split":
        # 左侧留白写大号编号与章名，右侧照片向左渐隐
        PH.place_faded(sl, photo, W * 0.48, -1, W * 0.52 + 1, H + 2, K.BG,
                       [(0.0, 1.0), (0.08, 0.95), (0.42, 0.0)], side="left", tone_main=K.MAIN, **dict(zip(("sat", "duo"), PH.tone_for(photo))))
        K.text(sl, 90, 150, 520, 170, K.T(num, 130, K.MAIN, True), anchor="b")
        K.rect(sl, 96, 336, 56, 5, (K.MAIN, K.mix(K.MAIN, "FFFFFF", .3), 0))
        _title(sl, 94, 358, 520, 120, name, 40, K.INK, anchor="t", lo=26)
        if desc:
            K.text(sl, 96, 480, 480, 80, K.T(desc, 16, K.SUB), anchor="t", spacing=1.45)
    elif style == "band":
        # 照片占上方，底部白条写章节号与章名
        ph = 480
        _photo(sl, photo, 0, 0, W, ph)
        K.rect(sl, 0, ph, W, H - ph, "FFFFFF" if K.BG != "FFFFFF" else K.mix(K.MAIN, "FFFFFF", .96))
        K.text(sl, 80, ph, 150, H - ph, K.T(num, 76, K.MAIN, True), anchor="m")
        K.vline(sl, 236, ph + 70, H - ph - 140, K.LINE, 1.0)
        _title(sl, 262, ph + 40, 900, 80, name, 36, K.INK, anchor="b", lo=26)
        if desc:
            K.text(sl, 264, ph + 128, 900, 50, K.T(desc, 15, K.MUTED), anchor="t")
    return sl

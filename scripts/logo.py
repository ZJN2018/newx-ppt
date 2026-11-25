# -*- coding: utf-8 -*-
"""单位 LOGO：封面、目录、转场、结束页放右上角，正文页放页眉右上角。

settings.logo 写本机图片路径（PNG/JPG/SVG 转出的 PNG，透明底最好）。LOGO 不改形、不改色相。
用户没给 LOGO 时，先从源材料里找：python3 logo.py find 年报.pdf [更多 PDF 或图片] --out logo-candidates
（年报、半年报、公司介绍的封面和前几页通常有一张干净的 LOGO 图），看候选图后选一张写进 settings.logo。

读入后按底色分三类处理（v2.6.2）：
- 透明底：直接用。
- 白底或黑底（四角是同一种近白或近黑，黑底多半是 PDF 透明度丢失）：与底色相同的像素变透明，
  边缘做去底色抗锯齿。字的空心等内部同色部分也透明，原色版只用在浅色背景上，看起来不变。
- 色块底（四角是同一种非白颜色，如黄底黑字的方形标）：底色本身就是 LOGO 的一部分，原样用，不抠、不反白；
  只有色块底色和页面背景撞色时才垫白色底卡。
放的时候先估出 LOGO 所在位置的背景色（把这一页已经画上去的照片、蒙版、色块叠合起来取样），再选版本：
- 浅色背景：原色 LOGO，不加底卡。
- 深色、主色或照片背景：纯白反白版（形状不变、颜色全换白，内部原本的白色镂空）。
- 反白后丢掉太多形状（多色渐变标等），或原色、白色在这块背景上都看不清：退回白色圆角底卡＋原色 LOGO。
"""
import os
import shutil
import subprocess
import sys
import tempfile

from PIL import Image, ImageFilter
from pptx.oxml.ns import qn
from pptx.util import Emu

import deckkit as K

PX = 9525
W = 1280
# 结构页：底卡高度与内边距；正文页：LOGO 高度。最大宽度防止横长 LOGO 过宽
PLATE_H, PLATE_PAD, PLATE_MAXW, PLATE_RIGHT, PLATE_TOP = 52, 10, 240, 40, 30
BODY_H, BODY_MAXW, BODY_RIGHT, BODY_TOP = 32, 170, 48, 53

MAX_SIDE = 1200        # 抠图前把超大的图缩到这个长边，LOGO 在页面上最宽 240px，足够
BG_TOL = 60            # 判为底色的容差（三通道差之和），吃掉 JPEG 噪点
WHITE_KNOCK = (0.80, 0.92)  # 反白时，最暗通道亮度从 0.80 到 0.92 的像素逐渐镂空（近白、低饱和）
WHITE_KEEP = 0.55      # 反白后保留不到一半：LOGO 主要是白色笔画，原图在深色背景上本来就清楚，不用反白版
BLOB_FILL = 0.80       # 不透明面积占外框 80% 以上、反白后又没有镂空：会变成一整块白，不用反白版
MIN_SHOW = 2.0         # 选版本时要求的最低对比度
LIGHT_BG = 0.35        # 背景相对亮度不低于这个值算浅色背景：放原色；低于算深色背景：放反白版
MIN_CONTRAST = 2.0     # 原色、白色两版都低于这个对比度时退回底卡
BADGE_MIN = 1.4        # 色块标的底色与背景对比度低于这个值（撞色）时垫底卡

_CUR = {"path": None, "white": None, "kind": None, "ink": "000000", "edge": "FFFFFF", "aspect": 1.0}
LAST = []  # 每次放 LOGO 的决定：(位置, 背景色, 用了哪个版本)，供测试和日志用


# ---------------- 读入与抠图 ----------------
def _corners(im, k=3):
    """四角 k×k 小块的平均 RGBA。"""
    w, h = im.size
    k = max(1, min(k, w, h))  # 过窄、过矮的图只取实际有的像素
    out = []
    for x0, y0 in ((0, 0), (w - k, 0), (0, h - k), (w - k, h - k)):
        px = [im.getpixel((x0 + i, y0 + j)) for i in range(k) for j in range(k)]
        out.append(tuple(sum(p[c] for p in px) / len(px) for c in range(4)))
    return out


def classify(im):
    """transparent（透明底）/ white（白底）/ black（黑底，多半是 PDF 透明度丢失）/ badge（色块底，原样用）。"""
    cs = _corners(im)
    if im.getchannel("A").getextrema()[0] < 40:  # 图片已有透明区域：作者已经抠好，不再按四角颜色抠
        return "transparent"
    if all(c[3] > 200 for c in cs):
        if all(min(c[:3]) >= 225 for c in cs):
            return "white"
        if all(max(c[:3]) <= 30 for c in cs):
            return "black"
    return "badge"


def matte(im, tol=BG_TOL):
    """去掉与底色相同的像素（底色取四角），返回 RGBA。边缘过渡像素做去底色抗锯齿。
    LOGO 内部和底色一样的部分（字的空心、白色衬底）也一并变透明：原色版只放在浅色背景上，
    透明和白色看起来一样；深色背景用反白版，不会露出深色。"""
    im = im.convert("RGBA")
    bg = tuple(int(sum(c[i] for c in _corners(im)) / 4) for i in range(3))
    w, h = im.size
    src = im.load()
    mask = Image.new("L", (w, h), 0)
    mp = mask.load()
    for y in range(h):
        for x in range(w):
            r, g, b, _ = src[x, y]
            if abs(r - bg[0]) + abs(g - bg[1]) + abs(b - bg[2]) <= tol:
                mp[x, y] = 255
    band = mask.filter(ImageFilter.MaxFilter(5))  # 底色外扩 2px：抗锯齿过渡带
    bp = band.load()
    out = im.copy()
    op = out.load()
    for y in range(h):
        for x in range(w):
            if mp[x, y]:
                op[x, y] = (255, 255, 255, 0)
            elif bp[x, y]:
                r, g, b, a = op[x, y]
                # 去底色（color-to-alpha）：把与底色混合的边缘像素还原成前景色＋部分透明
                al = 0.0
                for i, c in enumerate((r, g, b)):
                    d = (c - bg[i]) / (255 - bg[i]) if c > bg[i] else ((bg[i] - c) / bg[i] if c < bg[i] else 0.0)
                    al = max(al, d)
                al = max(0.0, min(1.0, al))
                if al <= 0.01:
                    op[x, y] = (255, 255, 255, 0)
                else:
                    fg = [int(max(0, min(255, (c - (1 - al) * bg[i]) / al))) for i, c in enumerate((r, g, b))]
                    op[x, y] = (fg[0], fg[1], fg[2], int(a * al))
    return out


def whiten(im):
    """纯白反白版：形状按透明度保留，颜色全换白；原本近白、低饱和的部分（字的空心、白字）镂空。
    返回 (图, 保留的不透明面积占比)。"""
    lo, hi = WHITE_KNOCK
    src = im.load()
    out = Image.new("RGBA", im.size, (255, 255, 255, 0))
    op = out.load()
    before = after = 0.0
    for y in range(im.height):
        for x in range(im.width):
            r, g, b, a = src[x, y]
            if not a:
                continue
            k = (min(r, g, b) / 255 - lo) / (hi - lo)
            k = max(0.0, min(1.0, k))
            na = a * (1 - k)
            before += a
            after += na
            op[x, y] = (255, 255, 255, int(na))
    return out, (after / before if before else 0.0)


def silhouette(im, color):
    """单色剪影：形状按透明度保留，颜色全换成 color（白字 LOGO 放在浅色页面上用）。"""
    r, g, b = int(color[0:2], 16), int(color[2:4], 16), int(color[4:6], 16)
    out = Image.new("RGBA", im.size, (r, g, b, 0))
    out.putalpha(im.getchannel("A"))
    return out


def merges_detail(im):
    """反白后会不会把 LOGO 内部的明暗区分抹掉（如黄底黑字：黄、黑都变白，字就没了）。
    只看反白后仍保留的像素（近白部分已镂空），其亮度两端（10% 与 90% 分位）对比度不低于 3:1 就算会抹掉。"""
    lo, hi = WHITE_KNOCK
    small = im.copy()
    small.thumbnail((240, 240))
    lums = sorted(_lum("%02X%02X%02X" % p[:3]) for p in small.getdata()
                  if p[3] > 128 and min(p[:3]) / 255 < lo)
    if len(lums) < 20:
        return False
    a, b = lums[len(lums) // 10], lums[len(lums) * 9 // 10]
    return (b + 0.05) / (a + 0.05) >= 3.0


def _ink(im):
    """LOGO 的主要笔画色：不透明、非近白像素的平均色。"""
    small = im.copy()
    small.thumbnail((200, 200))
    opaque = [p for p in small.getdata() if p[3] > 200]
    px = [p for p in opaque if min(p[:3]) < 230] or opaque  # 全是白色笔画（白字标）时就按白色算
    if not px:
        return "000000"
    return "%02X%02X%02X" % tuple(int(sum(p[i] for p in px) / len(px)) for i in range(3))


def _crop(im):
    alpha = im.getchannel("A")
    if alpha.getextrema()[0] < 250:  # 有透明：按透明度裁
        box = alpha.point(lambda a: 255 if a > 12 else 0).getbbox()
    else:  # 不透明：按与白色的差异裁白边
        box = im.convert("L").point(lambda v: 255 if v < 245 else 0).getbbox()
    return im.crop(box) if box else im


def _tmp_png(im, tag):
    fd, out = tempfile.mkstemp(prefix="niuppt_logo_%s_" % tag, suffix=".png")
    os.close(fd)
    im.save(out)
    return out


def prepare(path, out_dir=None):
    """读入 LOGO，按底色分类、抠图、裁边，另存原色版与反白版（本次构建独占的临时文件，绝不覆盖用户原图）。
    返回 dict(path, white, kind, ink, aspect, keep)；路径无效或图片读不了返回 None。"""
    if not path or not os.path.isfile(path):
        return None
    try:
        im = Image.open(path)
        im.load()
        im = im.convert("RGBA")
    except Exception as ex:  # 空文件、损坏或不是图片：提示后跳过 LOGO，不中断生成
        print("提醒：LOGO 文件「%s」无法解码（%s）" % (path, ex.__class__.__name__))
        return None
    if max(im.size) > MAX_SIDE:
        im.thumbnail((MAX_SIDE, MAX_SIDE), Image.LANCZOS)
    kind = classify(im)
    edge = "%02X%02X%02X" % tuple(int(sum(c[i] for c in _corners(im)) / 4) for i in range(3))
    if kind in ("white", "black"):
        im = matte(im)
    im = _crop(im)
    info = {"path": _tmp_png(im, "orig"), "white": None, "kind": kind, "ink": _ink(im), "edge": edge,
            "aspect": im.width / max(1, im.height), "keep": 0.0}
    info["dark"] = None
    if kind != "badge":  # 色块标自带底色，原样放；其余 LOGO 准备反白版和深色剪影，按落点背景挑对比度够的一版
        wim, keep = whiten(im)
        info["keep"] = keep
        alpha = im.getchannel("A")
        fill = sum(1 for a in alpha.getdata() if a > 128) / max(1, im.width * im.height)
        blob = fill >= BLOB_FILL and (1 - keep) < 0.03      # 实心色块且没有镂空：反白成一整块白，信息全丢
        if blob:  # 实心色块（如带透明留边的黄底黑字标）：按色块标处理，原样放
            info["kind"] = "badge"
        else:
            if keep >= WHITE_KEEP:
                info["white"] = _tmp_png(wim, "white")
            info["dark"] = _tmp_png(silhouette(im, "2A3038"), "dark")
    return info


def use(path, out_dir):
    """设定本稿 LOGO；没有 LOGO 时清空。返回是否启用。"""
    info = prepare(path, out_dir) if path else None
    _CUR.update(info or {"path": None, "white": None, "dark": None, "light_ink": False, "kind": None, "ink": "000000",
                         "edge": "FFFFFF", "aspect": 1.0})
    _CUR.setdefault("dark", None)
    _CUR.setdefault("light_ink", False)
    del LAST[:]
    if info:
        print("LOGO：%s，%s" % ({"transparent": "透明底", "white": "白底已抠除", "black": "黑底已抠除", "badge": "色块底，原样使用"}[info["kind"]],
                                "色块标原样放" if info["kind"] == "badge" else
                                ("按背景选原色、反白或深色剪影" if info["white"] else "反白会丢失信息，深色背景放原图")))
    K.TITLE_RESERVE = body_width() + 24 if _CUR["path"] else 0
    return bool(_CUR["path"])


def _size(h, max_w):
    w = h * _CUR["aspect"]
    if w > max_w:
        w, h = max_w, max_w / _CUR["aspect"]
    return w, h


def body_width():
    return _size(BODY_H, BODY_MAXW)[0] if _CUR["path"] else 0


# ---------------- 估算 LOGO 位置的背景色 ----------------
_IMG_CACHE = {}


def _hex(c):
    return tuple(int(c[i:i + 2], 16) for i in (0, 2, 4))


def _clr(node):
    """a:srgbClr 节点 → (r, g, b, alpha 0~1)；不认识的颜色返回 None。"""
    c = node.find(qn("a:srgbClr")) if node.tag != qn("a:srgbClr") else node
    if c is None:
        return None
    al = c.find(qn("a:alpha"))
    return _hex(c.get("val")) + ((int(al.get("val")) / 100000.0) if al is not None else 1.0,)


def _shape_color(sp, u, v):
    """色块在形状内归一化位置 (u, v) 处的颜色 (r, g, b, a)；没有填充返回 None。"""
    pr = sp._element.find(qn("p:spPr"))
    if pr is None:
        return None
    solid = pr.find(qn("a:solidFill"))
    if solid is not None:
        return _clr(solid)
    grad = pr.find(qn("a:gradFill"))
    if grad is None:
        return None
    stops = []
    for gs in grad.iter(qn("a:gs")):
        c = _clr(gs)
        if c:
            stops.append((int(gs.get("pos")) / 100000.0, c))
    if not stops:
        return None
    stops.sort()
    lin = grad.find(qn("a:lin"))
    import math
    ang = math.radians(int(lin.get("ang", "0")) / 60000.0) if lin is not None else 0.0
    ca, sa = math.cos(ang), math.sin(ang)  # DrawingML 角度顺时针、y 向下
    t = ((u - 0.5) * ca + (v - 0.5) * sa) / max(1e-6, abs(ca) + abs(sa)) + 0.5
    t = max(0.0, min(1.0, t))
    if t <= stops[0][0]:
        return stops[0][1]
    for (p0, c0), (p1, c1) in zip(stops, stops[1:]):
        if p0 <= t <= p1:
            k = (t - p0) / max(p1 - p0, 1e-9)
            return tuple(c0[i] + (c1[i] - c0[i]) * k for i in range(4))
    return stops[-1][1]


def _pic_color(sp, u, v):
    try:
        blob = sp.image.blob
    except Exception:
        return None
    key = hash(blob)
    im = _IMG_CACHE.get(key)
    if im is None:
        im = Image.open(__import__("io").BytesIO(blob)).convert("RGBA")
        im.thumbnail((640, 640))
        _IMG_CACHE.clear() if len(_IMG_CACHE) > 16 else None
        _IMG_CACHE[key] = im
    cl, cr, ct, cb = sp.crop_left, sp.crop_right, sp.crop_top, sp.crop_bottom
    ix = (cl + u * (1 - cl - cr)) * (im.width - 1)
    iy = (ct + v * (1 - ct - cb)) * (im.height - 1)
    ix, iy = int(max(0, min(im.width - 1, ix))), int(max(0, min(im.height - 1, iy)))
    r, g, b, a = im.getpixel((ix, iy))
    return (r, g, b, a / 255.0)


def background_at(sl, x, y, w, h, nx=6, ny=3):
    """把页面上已有的照片、色块、渐变蒙版按叠放顺序叠合，取 (x, y, w, h) 区域网格点的平均背景色（十六进制）。
    文字、线条、组合不计；页面底色按白色起算。"""
    pts = [(x + w * (i + 0.5) / nx, y + h * (j + 0.5) / ny) for i in range(nx) for j in range(ny)]
    acc = [0.0, 0.0, 0.0]
    shapes = []
    for sp in sl.shapes:
        if sp.shape_type not in (1, 13) or not sp.width or not sp.height or getattr(sp, "rotation", 0):
            continue  # 1 = 自选图形，13 = 图片
        if sp.shape_type == 1 and sp.has_text_frame and sp.text_frame.text.strip() and \
                sp._element.find(qn("p:spPr")).find(qn("a:solidFill")) is None and \
                sp._element.find(qn("p:spPr")).find(qn("a:gradFill")) is None:
            continue
        shapes.append((sp, sp.left / PX, sp.top / PX, sp.width / PX, sp.height / PX))
    for px_, py_ in pts:
        col = [255.0, 255.0, 255.0]
        for sp, sx, sy, sw, sh in shapes:
            if not (sx <= px_ <= sx + sw and sy <= py_ <= sy + sh):
                continue
            u, v = (px_ - sx) / sw, (py_ - sy) / sh
            c = _pic_color(sp, u, v) if sp.shape_type == 13 else _shape_color(sp, u, v)
            if not c:
                continue
            a = c[3]
            col = [col[i] * (1 - a) + c[i] * a for i in range(3)]
        for i in range(3):
            acc[i] += col[i]
    return "%02X%02X%02X" % tuple(int(round(v / len(pts))) for v in acc)


def _lum(hx):
    def ch(v):
        v /= 255.0
        return v / 12.92 if v <= 0.03928 else ((v + 0.055) / 1.055) ** 2.4
    r, g, b = _hex(hx)
    return 0.2126 * ch(r) + 0.7152 * ch(g) + 0.0722 * ch(b)


def contrast(a, b):
    la, lb = sorted((_lum(a), _lum(b)), reverse=True)
    return (la + 0.05) / (lb + 0.05)


def choose(bg):
    """按落点背景选版本，任何情况都不垫白色底卡。
    色块标自带底色，原样放。其余：浅色背景依次试 原色 → 深色剪影 → 反白，深色背景依次试 反白 → 原色，
    取第一个对比度不低于 MIN_SHOW 的；都不够时取对比度最高的一版。"""
    if _CUR["kind"] == "badge":
        return "original"
    tone = {"original": _CUR["ink"], "white": "FFFFFF", "dark": "2A3038"}
    have = {"original": True, "white": bool(_CUR.get("white")), "dark": bool(_CUR.get("dark"))}
    order = ["original", "dark", "white"] if _lum(bg) >= LIGHT_BG else ["white", "original"]
    order = [v for v in order if have[v]]
    for v in order:
        if contrast(tone[v], bg) >= MIN_SHOW:
            return v
    return max(order, key=lambda v: contrast(tone[v], bg))


def plate_color():
    """底卡颜色：默认白色；LOGO 笔画本身是浅色（如白字标）时改用深色，保证与笔画对比度。"""
    ref = _CUR["edge"] if _CUR["kind"] == "badge" else _CUR["ink"]
    if contrast(ref, "FFFFFF") >= 3.0:
        return "FFFFFF"
    for c in (K.MAIN, "1F2329"):
        if contrast(ref, c) >= 3.0:
            return c
    return "1F2329"


def _pic(sl, path, x, y, w, h):
    sl.shapes.add_picture(path, Emu(int(x * PX)), Emu(int(y * PX)), Emu(int(w * PX)), Emu(int(h * PX)))


def _place(sl, x, y, w, h, where, plate_box=None):
    bg = background_at(sl, x, y, w, h)
    how = choose(bg)
    _pic(sl, {"white": _CUR.get("white"), "dark": _CUR.get("dark")}.get(how) or _CUR["path"], x, y, w, h)
    LAST.append((where, bg, how))
    return how


def corner(sl):
    """结构页右上角：按背景深浅放原色或反白 LOGO，不垫底卡。"""
    if not _CUR["path"]:
        return
    inner_h = PLATE_H - PLATE_PAD * 2
    w, h = _size(inner_h, PLATE_MAXW - PLATE_PAD * 2)
    pw = w + PLATE_PAD * 2
    x = W - PLATE_RIGHT - pw
    return _place(sl, x + PLATE_PAD, PLATE_TOP + (PLATE_H - h) / 2, w, h, "结构页",
                  plate_box=(x, PLATE_TOP, pw, PLATE_H))


def body(sl):
    """正文页页眉右上角，与标题行垂直居中；页底色深时同样换反白版。"""
    if not _CUR["path"]:
        return
    w, h = _size(BODY_H, BODY_MAXW)
    pad = 6
    x, y = W - BODY_RIGHT - w, BODY_TOP + (BODY_H - h) / 2
    return _place(sl, x, y, w, h, "正文页", plate_box=(x - pad, y - pad, w + pad * 2, h + pad * 2))


# ---------------- 从源材料找 LOGO 候选 ----------------
FIND_PAGES = 4  # 只看前几页：封面、扉页、目录页通常有 LOGO


def _pymupdf():
    try:
        import pymupdf
        return pymupdf
    except ImportError:
        try:
            import fitz
            return fitz
        except ImportError:
            return None


def _apply_mask(img_path, mask_path):
    """把 PDF 单独存放的透明度蒙版（SMask）合成为图片的 alpha 通道。"""
    im = Image.open(img_path).convert("RGB")
    mask = Image.open(mask_path).convert("L")
    if mask.size != im.size:
        mask = mask.resize(im.size)
    im = im.convert("RGBA")
    im.putalpha(mask)
    im.save(img_path)


def _extract_pdf_images(pdf, work_dir, pages=FIND_PAGES):
    """抽出 PDF 前几页的内嵌图片到 work_dir（每份 PDF 独占的临时目录），返回文件列表。
    透明底图片的 SMask 合成回 alpha；优先 PyMuPDF，没有时用 poppler 的 pdfimages。"""
    mu = _pymupdf()
    files = []
    if mu is not None:
        with mu.open(pdf) as doc:
            seen = set()
            for pno in range(min(pages, doc.page_count)):
                for info in doc.get_page_images(pno):
                    xref, smask = info[0], info[1]
                    if xref in seen:
                        continue
                    seen.add(xref)
                    try:
                        pix = mu.Pixmap(doc, xref)
                        if pix.colorspace and pix.colorspace.n > 3:  # CMYK 等转 RGB
                            pix = mu.Pixmap(mu.csRGB, pix)
                        f = os.path.join(work_dir, "p%d-%d.png" % (pno + 1, xref))
                        pix.save(f)
                        if smask:
                            # 透明度蒙版（SMask）单独存出来再用 PIL 合成。不用 Pixmap(pix, mask)：
                            # 新版 PyMuPDF 取出的图可能已带一个全不透明的 alpha 通道，合成会报错或丢透明度
                            m = os.path.join(work_dir, "m%d.png" % smask)
                            mu.Pixmap(doc, smask).save(m)
                            _apply_mask(f, m)
                        files.append(f)
                    except Exception:
                        continue
        return files
    if not shutil.which("pdfimages"):
        print("提醒：抽取 PDF 图片需要 PyMuPDF 或 poppler（pdfimages），跳过「%s」" % pdf)
        return []
    prefix = os.path.join(work_dir, "img")
    subprocess.run(["pdfimages", "-png", "-f", "1", "-l", str(pages), pdf, prefix],
                   check=False, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    lst = subprocess.run(["pdfimages", "-list", "-f", "1", "-l", str(pages), pdf],
                         capture_output=True, text=True).stdout.splitlines()[2:]
    kinds = {}
    for row in lst:
        cols = row.split()
        if len(cols) > 2 and cols[1].isdigit():
            kinds[int(cols[1])] = cols[2]
    for n in sorted(kinds):
        f = "%s-%03d.png" % (prefix, n)
        if kinds[n] in ("smask", "mask") or not os.path.isfile(f):
            continue
        m = "%s-%03d.png" % (prefix, n + 1)
        if kinds.get(n + 1) == "smask" and os.path.isfile(m):  # pdfimages 把蒙版单独输出在紧随其后的编号
            _apply_mask(f, m)
        files.append(f)
    return files


def logo_score(path):
    """像 LOGO 的程度（0 表示不像）：有可见前景、尺寸适中、横向或方形、底色干净（白或透明）、颜色少。
    照片、整页扫描图、空白图、细长装饰条得 0。"""
    try:
        im = Image.open(path)
        im.load()
        im = im.convert("RGBA")
    except Exception:
        return 0
    w, h = im.size
    if w < 120 or h < 40 or w > 3000 or h > 2000:
        return 0
    ratio = w / h
    if not 0.8 <= ratio <= 8:
        return 0
    small = im.resize((min(w, 160), max(1, int(min(w, 160) / ratio))), Image.NEAREST)  # 不插值，免得边缘过渡色算成新颜色
    acc = small.load()
    px = [acc[i, j] for j in range(small.height) for i in range(small.width)]
    clean = sum(1 for r, g, b, a in px if a < 30 or (r > 235 and g > 235 and b > 235)) / len(px)
    if clean > 0.97:  # 几乎没有可见前景：空白图、纯白底
        return 0
    if clean < 0.35:  # 底色不干净：照片或深色底图
        return 0
    # 颜色集中度：量化成 8 色后前 4 色的占比（JPEG 噪点会摊薄占比，阈值放宽到 0.8）；LOGO 色块少而集中，照片分散
    flat = Image.new("RGB", small.size, "white")
    flat.paste(small, mask=small.getchannel("A"))
    hist = sorted(flat.quantize(colors=8).getcolors() or [], reverse=True)
    top4 = sum(n for n, _ in hist[:4]) / len(px)
    if top4 < 0.8:
        return 0
    corners = [small.getpixel(c) for c in ((0, 0), (small.width - 1, 0), (0, small.height - 1),
                                           (small.width - 1, small.height - 1))]
    corner_clean = sum(1 for r, g, b, a in corners if a < 30 or min(r, g, b) > 225)
    return round(clean * 2 + corner_clean * 0.5 + (1 if 1.5 <= ratio <= 6 else 0) + top4, 2)


def find(sources, out_dir, keep=5):
    """从 PDF（前几页的内嵌图片）和图片文件里挑 LOGO 候选，转成 PNG 存到 out_dir，按得分从高到低返回
    [(候选路径, 得分, 来源)]。PDF 抽出的中间文件放在本次新建的临时目录里，结束即删；用户给的文件只读不删。"""
    os.makedirs(out_dir, exist_ok=True)
    tmp_root = tempfile.mkdtemp(prefix="niuppt_logofind_")
    try:
        raw = []  # (文件, 来源说明)
        for k, src in enumerate(sources):
            if src.lower().endswith(".pdf"):
                work = os.path.join(tmp_root, "pdf%d" % k)  # 每份 PDF 独占一个目录，同名 PDF 不互相覆盖
                os.makedirs(work)
                raw += [(f, "%s 内嵌图 %s" % (src, os.path.basename(f))) for f in _extract_pdf_images(src, work)]
            elif os.path.isfile(src):
                raw.append((src, src))
        scored = sorted(((logo_score(f), f, where) for f, where in raw), key=lambda t: -t[0])
        picked = []
        for score, f, where in scored:
            if score <= 0 or len(picked) >= keep:
                break
            dst = os.path.join(out_dir, "logo-%d.png" % (len(picked) + 1))
            im = Image.open(f)
            im.convert("RGBA" if "A" in im.getbands() or im.mode == "P" else "RGB").save(dst)  # CMYK 等统一转换
            picked.append((dst, score, where))
        return picked
    finally:
        shutil.rmtree(tmp_root, ignore_errors=True)


if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="从源材料找 LOGO 候选")
    sub = ap.add_subparsers(dest="cmd", required=True)
    f = sub.add_parser("find", help="从 PDF 前几页和图片里挑 LOGO 候选")
    f.add_argument("sources", nargs="+")
    f.add_argument("--out", default="logo-candidates")
    a = ap.parse_args()
    got = find(a.sources, a.out)
    if not got:
        print("源材料里没找到像 LOGO 的图；可联网找该单位官方 LOGO（官网、百科），仍没有就不放 LOGO。")
        sys.exit(1)
    print("LOGO 候选（得分高的在前）。逐张看图，确认是汇报对象本单位的 LOGO（不是子公司、合作方或审计机构），"
          "选一张写进 deck.json 的 settings.logo：")
    for dst, score, src in got:
        print("  %s  得分 %.2f  来自 %s" % (dst, score, src))

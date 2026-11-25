"""装饰层：渐变标题块、渐变数字徽标、渐变描边卡片、大块浅色渐变背景形。

常用参数（按通行的商务演示审美整理）：
- 标题块：圆角矩形，45° 线性渐变，7% 处浅色 → 60% 处主色，白色粗体字
- 数字徽标：圆形，45° 渐变 浅 → 主 → 深，柔和下投影
- 卡片：无填充或白底，描边从上到下 50% → 10% 透明度渐变
- 背景形：整页尺度的箭头 / 圆 / 半圆，主色 5%–15% 透明度渐变，压在所有内容下面
"""
from lxml import etree
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from compose import is_solid, _rgb, box, text, PALETTE, PX

def mix(h, t, k):
    """h 向 t 混合 k（0–1）"""
    a = [int(h[i:i+2], 16) for i in (0, 2, 4)]; b = [int(t[i:i+2], 16) for i in (0, 2, 4)]
    return "".join(f"{round(x + (y - x) * k):02X}" for x, y in zip(a, b))
def light(h, k=0.4): return mix(h, "FFFFFF", k)
def dark(h, k=0.25): return mix(h, "000000", k)

def _clr(parent, hexv, alpha=None):
    c = etree.SubElement(parent, qn("a:srgbClr")); c.set("val", hexv)
    if alpha is not None: etree.SubElement(c, qn("a:alpha")).set("val", str(int(alpha * 1000)))
    return c

def _grad_el(stops, ang):
    g = etree.Element(qn("a:gradFill")); g.set("rotWithShape", "1")
    gl = etree.SubElement(g, qn("a:gsLst"))
    for pos, hexv, alpha in stops:
        gs = etree.SubElement(gl, qn("a:gs")); gs.set("pos", str(int(pos * 1000))); _clr(gs, hexv, alpha)
    lin = etree.SubElement(g, qn("a:lin")); lin.set("ang", str(int(ang * 60000))); lin.set("scaled", "1")
    return g

def grad_fill(shape, stops, ang=45, large_block=False):
    """stops: [(位置%, 颜色, 不透明度% 或 None)]；ang: 度，0=从左到右，90=从上到下"""
    if is_solid(shape) and not large_block:
        _, color, alpha = stops[len(stops) // 2]
        shape.fill.solid()
        shape.fill.fore_color.rgb = _rgb(light(color, 1 - alpha / 100) if alpha is not None else color)
        return
    spPr = shape._element.find(qn("p:spPr"))
    for tag in ("a:noFill", "a:solidFill", "a:gradFill"):
        for e in spPr.findall(qn(tag)): spPr.remove(e)
    geom = spPr.find(qn("a:prstGeom"))
    geom.addnext(_grad_el(stops, ang))

def grad_line(shape, stops, ang=90, w=1):
    spPr = shape._element.find(qn("p:spPr"))
    ln = spPr.find(qn("a:ln"))
    if ln is None: ln = etree.SubElement(spPr, qn("a:ln"))
    for e in list(ln): ln.remove(e)
    ln.set("w", str(int(w * PX)))
    ln.append(_grad_el(stops, ang))

def shadow(shape, blur=6, dist=4, alpha=15, color="000000"):
    spPr = shape._element.find(qn("p:spPr"))
    eff = spPr.find(qn("a:effectLst"))
    if eff is None: eff = etree.SubElement(spPr, qn("a:effectLst"))
    sh = etree.SubElement(eff, qn("a:outerShdw"))
    for k, v in dict(blurRad=int(blur * PX), dist=int(dist * PX), dir=5400000, algn="t", rotWithShape="0").items(): sh.set(k, str(v))
    _clr(sh, color, alpha)

def send_to_back(sl, shape):
    tree = sl.shapes._spTree; el = shape._element
    tree.remove(el); tree.insert(2, el)   # 0,1 是 nvGrpSpPr / grpSpPr

# ---------------- 组件 ----------------
def grad_tab(sl, x, y, w, h, label, color=None, size=16, align="c", radius=0.2):
    """渐变标题块：突出小标题"""
    color = color or PALETTE["navy2"]
    s = box(sl, x, y, w, h, shape=MSO_SHAPE.ROUNDED_RECTANGLE, adj=[radius])
    grad_fill(s, [(0, dark(color, .32), None), (100, color, None)], 0, large_block=True)
    text(sl, x + 14, y, w - 28, h, label, size, "FFFFFF", True, align=align, anchor="m", heading=True)
    return s

def grad_badge(sl, cx, cy, d, label, color=None, size=None):
    """渐变数字徽标"""
    color = color or PALETTE["navy2"]
    s = box(sl, cx - d / 2, cy - d / 2, d, d, shape=MSO_SHAPE.OVAL)
    grad_fill(s, [(18, light(color, 0.45), None), (60, color, None), (100, dark(color, 0.25), None)], 45)
    shadow(s, 6, 4, 18)
    text(sl, cx - d / 2, cy - d / 2, d, d, label, size or d * 0.42, "FFFFFF", True, align="c", anchor="m")
    return s

def soft_card(sl, x, y, w, h, color=None, fill="FFFFFF", radius=0.05):
    """白底卡片 + 上实下虚的渐变描边 + 很淡的投影"""
    color = color or PALETTE["navy2"]
    solid = is_solid(sl)
    if not solid and fill == "FFFFFF":
        fill = light(color, 0.95)
    s = box(sl, x, y, w, h, fill=fill, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
            adj=[radius if solid else max(radius, 0.08)])
    if solid:
        s.line.color.rgb = _rgb(light(color, 0.67))
        s.line.width = int(1.25 * PX)
    else:
        grad_line(s, [(0, color, 55), (100, color, 10)], 90, 1)
    shadow(s, 14, 4, 8, PALETTE["navy"])
    return s

def backdrop(sl, kind, x, y, w, h, color=None, a0=14, a1=3, ang=90):
    """整页尺度的浅色渐变背景形，放在最底层。kind: up_arrow | right_arrow | circle | half_left | half_right | band"""
    color = color or PALETTE["navy2"]
    shp = {"up_arrow": MSO_SHAPE.UP_ARROW, "right_arrow": MSO_SHAPE.RIGHT_ARROW, "circle": MSO_SHAPE.OVAL,
           "half_left": MSO_SHAPE.OVAL, "half_right": MSO_SHAPE.OVAL, "band": MSO_SHAPE.RECTANGLE}[kind]
    s = box(sl, x, y, w, h, shape=shp)
    if kind == "up_arrow": s.adjustments[0] = 0.42; s.adjustments[1] = 0.32
    if kind == "right_arrow": s.adjustments[0] = 0.5; s.adjustments[1] = 0.6
    grad_fill(s, [(0, color, a0), (100, color, a1)], ang)
    send_to_back(sl, s)
    s._element.xpath(".//p:cNvPr")[0].set("descr", "niuppt-background")
    return s

def fade_bar(sl, x, y, w, h, label=None, color=None, size=16, a0=40):
    """渐隐标题条：左浓右淡（decoration-header 的通用参数：0° 渐变，40% → 0% 不透明）"""
    color = color or PALETTE["navy2"]
    s = box(sl, x, y, w, h)
    grad_fill(s, [(0, color, a0), (100, color, 0)], 0, large_block=True)
    box(sl, x, y, 5, h, fill=color)
    if label: text(sl, x + 16, y, w - 20, h, label, size, PALETTE["navy"], True, anchor="m", heading=True)
    return s

def half_stadium(sl, side, x0, y0, x1, y1, color=None, a_outer=4, a_inner=16):
    """半跑道形背景：两块左右对称拼在页面中线，半圆一侧朝向中间，直边贴页面外缘。
    side='left' 表示放在左半页。渐变靠中间浓、外缘淡。"""
    import math
    color = color or PALETTE["navy2"]
    r = (y1 - y0) / 2; cy = (y0 + y1) / 2
    ts = [i * math.pi / 48 for i in range(1, 48)]
    if side == "left":      # 直边在左（x0），半圆在右（x1）
        cx = x1 - r
        pts = [(x0, y0), (cx, y0)] + [(cx + r * math.sin(t), cy - r * math.cos(t)) for t in ts] + [(cx, y1), (x0, y1)]
        ang = 0             # 左→右：外淡内浓
    else:                   # 直边在右（x1），半圆在左（x0）
        cx = x0 + r
        pts = [(x1, y0), (cx, y0)] + [(cx - r * math.sin(t), cy - r * math.cos(t)) for t in ts] + [(cx, y1), (x1, y1)]
        ang = 180
    fb = sl.shapes.build_freeform(int(pts[0][0] * PX), int(pts[0][1] * PX), scale=1.0)
    fb.add_line_segments([(int(x * PX), int(y * PX)) for x, y in pts[1:]], close=True)
    s = fb.convert_to_shape()
    s.line.fill.background()
    spPr = s._element.find(qn("p:spPr"))
    for e in spPr.findall(qn("a:solidFill")) + spPr.findall(qn("a:noFill")): spPr.remove(e)
    st = s._element.find(qn("p:style"))
    if st is not None: s._element.remove(st)
    g = _grad_el([(0, color, a_outer), (45, color, a_outer + 1), (100, color, a_inner)], ang)
    spPr.find(qn("a:custGeom")).addnext(g)
    send_to_back(sl, s)
    s._element.xpath(".//p:cNvPr")[0].set("descr", "niuppt-background")
    return s

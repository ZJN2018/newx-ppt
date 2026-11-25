# -*- coding: utf-8 -*-
"""IconPark（线性）→ PPT 原生形状。
图标数据：@iconify-json/icon-park-outline（ByteDance IconPark，Apache-2.0）。
每个图标转成一个 custGeom 自定义形状：矢量、可在 PPT 里直接改色改线宽。"""
import json
import os
from svgelements import SVG, Shape, Path, Move, Line, Close, CubicBezier, QuadraticBezier, Arc
from pptx.util import Emu, Pt
from pptx.oxml.ns import qn
from pptx.enum.shapes import MSO_SHAPE
from lxml import etree

HERE = os.path.dirname(os.path.abspath(__file__))
_DATA = None
PX = 9525
SCALE = 100  # custGeom 坐标 = svg 坐标 × 100

# 中文语义 → IconPark 图标名（工作汇报常用）。名字写错或词表里没有时，先按英文名前缀/词段近似匹配，仍找不到用通用图标 FALLBACK，
# 不再回退成编号徽标（同一组里混出一个“1”很难看）
KEYWORDS = {
    "增长": "trending-up", "上升": "trending-up", "下降": "trending-down", "趋势": "trend",
    "收入": "paper-money", "成本": "wallet", "利润": "finance", "资金": "bank", "应收": "bill", "回款": "bank-transfer",
    "客户": "customer", "团队": "peoples", "人员": "user-business", "合作": "cooperative-handshake",
    "合同": "file-success", "文件": "file-success", "审计": "audit", "复核": "check-correct", "复查": "check-correct",
    "风险": "caution", "异常": "attention", "预警": "remind", "问题": "error", "关闭": "check-one", "完成": "check-one",
    "安全": "protect", "控制": "protect", "临时控制": "protect", "规则": "balance",
    "仓储": "warehousing", "库房": "warehousing", "运输": "truck", "配送": "delivery", "物流": "transport",
    "交接": "exchange", "装车": "truck", "回单": "file-success", "温控": "thermometer", "冷链": "thermometer",
    "时间": "time", "计划": "plan", "日历": "calendar", "到期": "calendar", "节点": "flag", "目标": "target",
    "复测": "experiment", "测试": "experiment", "流程": "process-line", "循环": "cycle", "改进": "repair",
    "工具": "tool", "设置": "setting", "数据": "data", "分析": "analysis", "市场": "market-analysis",
    "占比": "chart-pie", "结构": "chart-pie", "排名": "ranking", "对比": "balance-two", "柱状": "chart-histogram",
    "城市": "city", "区域": "local", "地图": "area-map", "工厂": "factory-building", "楼宇": "building-one",
    "荣誉": "trophy", "奖项": "gold-medal", "思路": "lamp", "重点": "pushpin", "沟通": "message", "电话": "phone",
    "旺季": "lightning", "效率": "lightning", "速度": "stopwatch", "等待": "hourglass",
    "公司": "building-one", "企业": "building-one", "集团": "building-one", "上市": "stock-market", "股票": "stock-market",
    "股东": "peoples", "业务": "briefcase", "期间": "calendar", "电力": "lightning", "发电": "lightning", "能源": "lightning",
    "水电": "water-level", "研发": "experiment", "创新": "lamp", "海外": "earth", "国际": "earth", "订单": "file-success",
    "现金": "paper-money", "负债": "bill", "投资": "chart-line", "产品": "box", "品牌": "certificate", "政策": "file-success",
}
FALLBACK = "pushpin"


def _data():
    global _DATA
    if _DATA is None:
        with open(os.path.join(os.path.dirname(HERE), "assets", "icon-park-outline.json"), encoding="utf-8") as f:
            _DATA = json.load(f)
    return _DATA


def resolve(name_or_kw):
    """支持 IconPark 英文名或中文关键词。"""
    if not name_or_kw:
        return None
    d = _data()["icons"]
    if name_or_kw in d:
        return name_or_kw
    if name_or_kw in KEYWORDS:
        return KEYWORDS[name_or_kw]
    for k, v in KEYWORDS.items():  # 关键词包含匹配，如“湖滨晚到风险”→ 风险
        if k in name_or_kw:
            return v
    key = name_or_kw.strip().lower().replace(" ", "-").replace("_", "-")
    if key in d:
        return key
    pref = sorted((n for n in d if n.startswith(key + "-")), key=lambda n: (not n.endswith("-one"), len(n), n))
    if pref:                                   # building → building-one
        return pref[0]
    token = sorted((n for n in d if key in n.split("-")), key=lambda n: (len(n), n))
    if token:                                  # stock → stock-market
        return token[0]
    return None


def resolve_or_fallback(name_or_kw):
    return resolve(name_or_kw) or FALLBACK


def _paths(name):
    d = _data()
    body = d["icons"][name]["body"]
    w = d["icons"][name].get("width", d.get("width", 48))
    h = d["icons"][name].get("height", d.get("height", 48))
    svg = SVG.parse_string if hasattr(SVG, "parse_string") else None
    from io import StringIO
    doc = SVG.parse(StringIO(
        f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {w} {h}" width="{w}" height="{h}" '
        f'color="#000000">{body}</svg>'), color="#000000")
    out = []
    sw = 4.0
    for e in doc.elements():
        if not isinstance(e, Shape):
            continue
        p = Path(e)
        p.reify()
        filled = e.fill is not None and e.fill.value is not None and str(e.fill) not in ("none", "None")
        stroked = e.stroke is not None and e.stroke.value is not None and str(e.stroke) not in ("none", "None")
        if stroked and e.stroke_width:
            sw = float(e.stroke_width)
        out.append((p, filled, stroked))
    return out, w, h, sw


def _f(v):
    return str(int(round(v * SCALE)))


def _build_custgeom(paths, w, h):
    A = "http://schemas.openxmlformats.org/drawingml/2006/main"
    cg = etree.Element(qn("a:custGeom"))
    for t in ("a:avLst", "a:gdLst", "a:ahLst", "a:cxnLst"):
        etree.SubElement(cg, qn(t))
    rect = etree.SubElement(cg, qn("a:rect"), l="0", t="0", r="r", b="b")
    pl = etree.SubElement(cg, qn("a:pathLst"))
    for p, filled, stroked in paths:
        pe = etree.SubElement(pl, qn("a:path"), w=_f(w), h=_f(h))
        if not filled:
            pe.set("fill", "none")
        if not stroked:
            pe.set("stroke", "0")

        def pt(parent, x, y):
            etree.SubElement(parent, qn("a:pt"), x=_f(x), y=_f(y))

        for seg in p:
            if isinstance(seg, Move):
                m = etree.SubElement(pe, qn("a:moveTo")); pt(m, seg.end.x, seg.end.y)
            elif isinstance(seg, Close):
                etree.SubElement(pe, qn("a:close"))
            elif isinstance(seg, Line):
                l = etree.SubElement(pe, qn("a:lnTo")); pt(l, seg.end.x, seg.end.y)
            elif isinstance(seg, CubicBezier):
                c = etree.SubElement(pe, qn("a:cubicBezTo"))
                pt(c, seg.control1.x, seg.control1.y); pt(c, seg.control2.x, seg.control2.y); pt(c, seg.end.x, seg.end.y)
            elif isinstance(seg, QuadraticBezier):
                c = etree.SubElement(pe, qn("a:quadBezTo"))
                pt(c, seg.control.x, seg.control.y); pt(c, seg.end.x, seg.end.y)
            elif isinstance(seg, Arc):
                for cb in seg.as_cubic_curves():
                    c = etree.SubElement(pe, qn("a:cubicBezTo"))
                    pt(c, cb.control1.x, cb.control1.y); pt(c, cb.control2.x, cb.control2.y); pt(c, cb.end.x, cb.end.y)
    return cg


def icon(sl, key, x, y, size, color="0E5A78", weight=1.0):
    """在 (x,y) 放一个 size×size 像素的图标；成功返回形状，找不到返回 None。
    weight：线宽倍数（1.0 = IconPark 原始 4/48 比例）。"""
    name = resolve_or_fallback(key)
    paths, w, h, sw = _paths(name)
    s = sl.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(int(x * PX)), Emu(int(y * PX)),
                            Emu(int(size * PX)), Emu(int(size * PX)))
    spPr = s._element.spPr
    prst = spPr.find(qn("a:prstGeom"))
    cg = _build_custgeom(paths, w, h)
    prst.addprevious(cg); spPr.remove(prst)
    # 填充（只对 fill≠none 的子路径生效）
    s.fill.solid(); s.fill.fore_color.rgb = __import__("pptx").dml.color.RGBColor.from_string(color)
    ln = s.line
    ln.color.rgb = __import__("pptx").dml.color.RGBColor.from_string(color)
    ln.width = Pt(size * sw / w * weight * 0.75)
    lnx = spPr.find(qn("a:ln"))
    lnx.set("cap", "rnd")
    if lnx.find(qn("a:round")) is None:
        etree.SubElement(lnx, qn("a:round"))
    st = s._element.find(qn("p:style"))
    if st is not None:
        s._element.remove(st)
    s.name = f"icon:{name}"
    return s

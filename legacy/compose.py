"""底层绘图原语：画布、主题、文本框、形状、表格、横幅。

坐标单位 px（1280×720 画布，1px = 9525 EMU）。字号参数也按 px 传入，内部换算为磅（×0.75）。
"""
import copy, math, re
import os as _os
# 旧名兼容：NIUMA_* 环境变量自动映射为 NIUPPT_*（已设置的新名优先）
for _k in list(_os.environ):
    if _k.startswith("NIUMA_") and "NIUPPT_" + _k[6:] not in _os.environ:
        _os.environ["NIUPPT_" + _k[6:]] = _os.environ[_k]
from lxml import etree
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR, MSO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_TICK_LABEL_POSITION
from pptx.oxml.ns import qn

PX = 9525
W, H, M = 1280, 720, 40

PALETTE = dict(navy="1C2F7A", navy2="2E4BA8", dark="122046", peach="E8A07A", peach_l="FBE9DE",
               ink="1A1A1A", muted="5F6470", line="D5D9E2", tint="EEF1F7", white="FFFFFF")
# 全部使用开源字体（SIL OFL）：正文与数字用思源黑体，标题用思源宋体，避免商用字体授权问题。
FONT_CN, FONT_EN = "Source Han Sans CN", "Source Han Sans CN"

HEADING_FONT = "Source Han Serif CN"
_CJK = r"\u3400-\u9fff"
_HSPACE = r"[^\S\r\n\v]+"
_MIXED_SPACE = re.compile(
    rf"(?<=[{_CJK}]){_HSPACE}(?=[+−-]?[0-9０-９])|"
    rf"(?<=[0-9０-９%％]){_HSPACE}(?=[{_CJK}])|(?<=[0-9０-９]){_HSPACE}(?=[%％])"
)


def compact_text(value):
    # Chart categories can also be dates or numbers in the editable workbook.
    return _MIXED_SPACE.sub("", value) if isinstance(value, str) else value


def compact_runs(runs):
    """Keep rich-text styling while removing a gap split across adjacent runs."""
    combined = "".join(run[0] for run in runs)
    gaps = {i for m in _MIXED_SPACE.finditer(combined) for i in range(*m.span())}
    offset = 0
    for value, *style in runs:
        compact = "".join(c for i, c in enumerate(value, offset) if i not in gaps)
        yield (compact, *style)
        offset += len(value)


# ---------------- 基础 ----------------
def _rgb(h): return RGBColor.from_string(h)
def _mix(h, t, k):
    a = [int(h[i:i+2], 16) for i in (0, 2, 4)]; b = [int(t[i:i+2], 16) for i in (0, 2, 4)]
    return "".join(f"{round(x + (y - x) * k):02X}" for x, y in zip(a, b))
def on_accent():
    from brand import contrast
    if contrast(PALETTE["dark"], PALETTE["peach"]) >= 4.5:
        return PALETTE["dark"]
    return max(("FFFFFF", PALETTE["ink"]), key=lambda color: contrast(color, PALETTE["peach"]))

def alt(i):
    """并列条目的交替色：偶数项主色，奇数项强调色。
    整份报告用到状态色（按期 / 滞后 / 风险）时改为主色与主色浅调交替，强调色留给“滞后”，避免装饰色被读成状态。"""
    if i % 2 == 0:
        return PALETTE["navy"]
    return PALETTE.get("alt") or PALETTE["peach"]

def on_main(k=0.75): return _mix(PALETTE["navy"], "FFFFFF", k)   # 主色底上的次要文字

def _nostyle(sh):
    st = sh._element.find(qn("p:style"))
    if st is not None: sh._element.remove(st)

def is_solid(obj):
    return getattr(obj.part.package.presentation_part, "niuppt_skin", None) == "solid"

def box(sl, x, y, w, h, fill=None, line=None, shape=MSO_SHAPE.RECTANGLE, lw=1, adj=None):
    s = sl.shapes.add_shape(shape, Emu(int(x*PX)), Emu(int(y*PX)), Emu(int(w*PX)), Emu(int(h*PX)))
    if fill: s.fill.solid(); s.fill.fore_color.rgb = _rgb(fill)
    else: s.fill.background()
    if line: s.line.color.rgb = _rgb(line); s.line.width = Emu(int(lw*PX))
    else: s.line.fill.background()
    if adj:
        for i, v in enumerate(adj): s.adjustments[i] = v
    _nostyle(s); return s

def line(sl, x1, y1, x2, y2, color, lw=1):
    c = sl.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Emu(int(x1*PX)), Emu(int(y1*PX)), Emu(int(x2*PX)), Emu(int(y2*PX)))
    c.line.color.rgb = _rgb(color); c.line.width = Emu(int(lw*PX)); _nostyle(c); return c

def set_typeface(run, latin, east_asian):
    run.font.name = latin
    rPr = run._r.get_or_add_rPr(); ea = rPr.find(qn("a:ea"))
    if ea is None: ea = rPr.makeelement(qn("a:ea"), {}); rPr.append(ea)
    ea.set("typeface", east_asian)


def _font(run, size, color, bold, *, heading=False, role=None):
    floors = {"source": 11, "page_number": 12, "title": 20}
    size = max(size, floors.get(role, 16))
    if 26 <= size < 28 and re.fullmatch(r"[\d.,%+−—\-/\s亿元万家个件年月倍]+", run.text.strip()):
        size = 28
    f = run.font; f.size = Pt(size*0.75); f.bold = bold; f.color.rgb = _rgb(color)
    set_typeface(run, HEADING_FONT if heading else FONT_EN, HEADING_FONT if heading else FONT_CN)

def text(sl, x, y, w, h, content, size=14, color="1A1A1A", bold=False, align="l", anchor="t", spacing=1.3, *, heading=False, role=None):
    """content: str | list[str]（多段） | list[list[(text,size,color,bold)]]（多段多片段）"""
    tb = sl.shapes.add_textbox(Emu(int(x*PX)), Emu(int(y*PX)), Emu(int(w*PX)), Emu(int(h*PX)))
    if role:
        tb._element.xpath('.//p:cNvPr')[0].set('descr', 'niuppt-role:' + role)
    tf = tb.text_frame; tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}[anchor]
    paras = content if isinstance(content, list) else [content]
    sizes = [r[1] for para in paras if isinstance(para, list) for r in para] or [size]
    if spacing == 1.3 and max(sizes) >= 22: spacing = 1.0   # 大号数字用单倍行距，避免在 PowerPoint 里下沉
    bp = tf._txBody.find(qn("a:bodyPr"))
    for e in list(bp): bp.remove(e)
    bp.append(bp.makeelement(qn("a:noAutofit"), {}))
    for i, para in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}[align]
        p.line_spacing = spacing
        runs = para if isinstance(para, list) else [(para, size, color, bold)]
        for t, sz, col, b in compact_runs(runs):
            r = p.add_run(); r.text = t; _font(r, sz, col, b, heading=heading, role=role)
    return tb

def metric_color(meaning):
    """箭头只说明数值变化方向，颜色由业务含义决定。"""
    return PALETTE[{"改善": "navy", "承压": "peach", "中性": "muted"}[meaning]]


def page_number(sl, num, color):
    """所有骨架使用同一页码框。单倍行距消除小字号在渲染时的下沉。"""
    x, y, w, h = W - M - 30, 690, 30, 22
    box(sl, x, y, w, h, color, shape=MSO_SHAPE.ROUNDED_RECTANGLE, adj=[0.5])
    text(sl, x, y + 0.75, w, h, str(num), 12, "FFFFFF", True, align="c", anchor="m", spacing=1.0, role="page_number")


def est_lines(s, width, size):
    u = sum(1.0 if ord(c) > 0x2E80 else 0.55 for c in s)
    return max(1, math.ceil(u / max(1, int(width / (size * 1.06)))))

# ---------------- 主题 ----------------
def new_deck():
    prs = Presentation(); prs.slide_width = Emu(W*PX); prs.slide_height = Emu(H*PX)
    # 改写主题色与主题字体，让移植进来的组件自动套用本套配色
    theme_part = [r.target_part for r in prs.slide_masters[0].part.rels.values() if r.reltype.endswith("/theme")][0]
    root = etree.fromstring(theme_part.blob)
    ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
    cmap = {"dk2": PALETTE["dark"], "lt2": PALETTE["tint"], "accent1": PALETTE["navy"], "accent2": PALETTE["peach"],
            "accent3": PALETTE["navy2"], "accent4": "8FA3D9", "accent5": "F3C9AE", "accent6": "6B7280"}
    for k, v in cmap.items():
        el = root.find(f".//a:clrScheme/a:{k}", ns)
        for ch in list(el): el.remove(ch)
        etree.SubElement(el, qn("a:srgbClr")).set("val", v)
    for tag in ("majorFont", "minorFont"):
        fe = root.find(f".//a:fontScheme/a:{tag}", ns)
        fe.find("a:latin", ns).set("typeface", FONT_EN); fe.find("a:ea", ns).set("typeface", FONT_CN)
    theme_part._blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    return prs

# ---------------- 页面框架 ----------------
def banner(sl, y, h, left, right):
    """左深蓝箭头块 + 右杏色块。left/right = {label, metrics:[(name, value, note)]}"""
    split = 640
    box(sl, split - 20, y, W - M - split + 20, h, PALETTE["peach"])
    box(sl, M, y, split - M + 30, h, PALETTE["navy"], shape=MSO_SHAPE.PENTAGON, adj=[0.3])
    for (x0, x1, blk, fg, sub) in ((M + 22, split - 30, left, "FFFFFF", on_main()), (split + 70, W - M - 20, right, on_accent(), on_accent())):
        text(sl, x0, y + 12, x1 - x0, 24, blk["label"], 16, fg, True)
        n = len(blk["metrics"]); cw = (x1 - x0) / n
        for i, (name, val, note) in enumerate(blk["metrics"]):
            cx = x0 + i * cw
            if i: line(sl, cx - 10, y + 48, cx - 10, y + h - 14, sub, 1)
            text(sl, cx, y + 44, cw - 20, 22, name, 14, sub)
            text(sl, cx, y + 64, cw - 20, 40, val, 30, fg, True)
            text(sl, cx, y + 108, cw - 20, 20, note, 12, sub)

def _text_units(v):
    return sum(1.0 if ord(c) > 0x2E80 else 0.55 for c in compact_text(str(v)))

def _auto_col_w(header, rows):
    """按每列最长文字定相对列宽：够放即可，长列多分，短列少分。"""
    out = []
    for c in range(len(header)):
        cells = [header[c]] + [r[c] for r in rows]
        out.append(min(22.0, max(3.0, max(_text_units(v) for v in cells))))
    return out

def table_fit(w, header, rows, col_w=None, size=16):
    """估算表格紧凑尺寸：返回 (列宽列表, 各行高列表, 总高)。行高只比文字略高。"""
    col_w = list(col_w or _auto_col_w(header, rows))
    k = w / sum(col_w); widths = [c * k for c in col_w]
    fs = max(size, 16); line_h = fs * 1.25 * 1.05   # 正文字号下限 16px，_font 会把更小的字号抬到 16
    heights = []
    for r, cells in enumerate([header] + list(rows)):
        lines = max(est_lines(compact_text(str(v)), max(20, cw - 18), fs) for v, cw in zip(cells, widths))
        heights.append(lines * line_h + 12)
    return widths, heights, sum(heights)

def table(sl, x, y, w, h, header, rows, col_w=None, hl_rows=(), size=16, bold_rows=()):
    """原生表格。行高按文字收紧，不再均分 h；返回表格实际底边 y，空出的高度留给结论带或说明。"""
    widths, heights, need = table_fit(w, header, rows, col_w, size)
    if need > h:
        k = h / need; heights = [v * k for v in heights]
        try:
            from skin_solid import WARN
            WARN.append(f"表格：内容需要约 {need:.0f}px，给了 {h:.0f}px，可能挤字；减行或加高")
        except ImportError: pass
    used = min(need, h)
    if h - used >= 40:
        try:
            from skin_solid import WARN
            WARN.append(f"表格：按文字收高后下方空出约 {h - used:.0f}px，请在表格下方（y={y + used + 12:.0f} 起）放结论带、说明或小卡片，不要留白")
        except ImportError: pass
    gt = sl.shapes.add_table(len(rows) + 1, len(header), Emu(int(x*PX)), Emu(int(y*PX)), Emu(int(w*PX)), Emu(int(used*PX)))
    t = gt.table
    for i, cw in enumerate(widths): t.columns[i].width = Emu(int(cw*PX))
    for i, rh in enumerate(heights): t.rows[i].height = Emu(int(rh*PX))
    tblPr = gt._element.graphic.graphicData.tbl.tblPr
    sid = tblPr.find(qn("a:tableStyleId"))
    if sid is not None: sid.text = "{5940675A-B579-460E-94D1-54222C63F5DA}"  # 无样式，仅网格
    for r in range(len(rows) + 1):
        for c in range(len(header)):
            cell = t.cell(r, c); val = header[c] if r == 0 else rows[r-1][c]
            cell.margin_left = cell.margin_right = Emu(8*PX); cell.margin_top = cell.margin_bottom = Emu(3*PX)
            cell.fill.solid()
            if r == 0:
                from deco import _grad_el
                # Adjacent cells continue the same editable gradient across the header.
                far = _mix(PALETTE["navy"], "000000", .08)
                start = _mix(PALETTE["dark"], far, c / len(header))
                end = _mix(PALETTE["dark"], far, (c + 1) / len(header))
                solid = cell._tc.get_or_add_tcPr().find(qn("a:solidFill"))
                solid.getparent().replace(solid, _grad_el([(0,start,None),(100,end,None)],0))
            elif (r - 1) in hl_rows: cell.fill.fore_color.rgb = _rgb(PALETTE["peach_l"])
            else: cell.fill.fore_color.rgb = _rgb("FFFFFF" if r % 2 else PALETTE["tint"])
            tf = cell.text_frame; tf.text = ""; p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT if c == 0 else PP_ALIGN.CENTER
            run = p.add_run(); run.text = compact_text(str(val))
            _font(run, size, "FFFFFF" if r == 0 else PALETTE["ink"], r == 0 or c == 0 or (r - 1) in bold_rows)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    return y + used

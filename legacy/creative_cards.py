"""Geometry, typography and brand checks for the reviewed T1-301 task unit."""
import re

from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_AUTO_SIZE, MSO_ANCHOR
from pptx.util import Pt
from layout_checks import EMU_PT, text_regions


def _units(shapes):
    for shape in shapes:
        if shape.shape_type == 6:
            if shape.name == "PART|T1-301|单元":
                yield shape
            else:
                yield from _units(shape.shapes)


def _parts(unit):
    children = list(unit.shapes)
    slots = {s.name.split("|")[1].strip(): s for s in children
             if s.name.startswith("T1-301 | task_")}
    backgrounds = sorted((s for s in children if not s.text.strip()),
                         key=lambda s: s.width * s.height, reverse=True)
    if len(backgrounds) < 2 or not {"task_detail", "task_title", "task_status"} <= slots.keys():
        raise ValueError("T1-301 缺少背景、编号区域或文字槽")
    return backgrounds[0], backgrounds[1], slots


def _bounds(shape):
    return tuple(v / EMU_PT for v in (shape.left, shape.top, shape.width, shape.height))


def _rgb(shape):
    nodes = shape._element.xpath('./p:spPr/a:solidFill/a:srgbClr')
    return nodes[0].get("val", "").upper() if nodes else ""


def task_issues(slide, palette=None):
    problems = []
    for unit in _units(slide.shapes):
        background, badge, slots = _parts(unit)
        ident = slots["task_status"].text.strip()
        name = f"T1-301 编号 {ident}"
        # The expected fill travels with the editable object and survives reopening.
        prop = background._element.xpath('./p:nvSpPr/p:cNvPr')[0]
        expected = (palette or {}).get("tint") or prop.get("descr", "").removeprefix("niuppt-task-fill:")
        color = _rgb(background)
        if (expected and color != expected) or (not expected and color in ("", "E8E8E8")):
            problems.append(("卡片配色", name, "底色与品牌浅色不符"))
        for role in ("task_title", "task_detail"):
            slot = slots[role]
            tf = slot.text_frame
            if tf.word_wrap is not True or tf.auto_size != MSO_AUTO_SIZE.NONE:
                problems.append(("文本换行", name + " " + role, "文字槽须开启换行并固定边界"))
            if slot.left < badge.left + badge.width or slot.left + slot.width > background.left + background.width:
                problems.append(("卡片边界", name + " " + role, "文字槽侵入编号区或超出卡片"))
            if slot.top < background.top or slot.top + slot.height > background.top + background.height:
                problems.append(("卡片边界", name + " " + role, "文字槽纵向超出卡片"))
            _, errors = text_regions(tf, _bounds(slot), role="body")
            for kind, detail in errors:
                if kind != "字号":
                    problems.append((kind, name + " " + role, detail))
        title, detail = slots["task_title"], slots["task_detail"]
        if title.top + title.height > detail.top:
            problems.append(("卡片边界", name, "标题和正文区域重叠"))
    return problems


def fit_task_units(slide, palette):
    from compose import HEADING_FONT, set_typeface
    from render import preview_font
    from creative_assets import _text

    for unit in _units(slide.shapes):
        bg, badge, slots = _parts(unit)
        bg.fill.solid()
        bg.fill.fore_color.rgb = RGBColor.from_string(palette["tint"])
        bg._element.xpath('./p:nvSpPr/p:cNvPr')[0].set("descr", "niuppt-task-fill:" + palette["tint"])
        left = badge.left + badge.width + Pt(10)
        for role in ("task_title", "task_detail"):
            slot = slots[role]
            slot.left = left
            slot.width = bg.left + bg.width - Pt(10) - left
            tf = slot.text_frame
            tf.word_wrap = True
            tf.auto_size = MSO_AUTO_SIZE.NONE
            tf.vertical_anchor = MSO_ANCHOR.TOP
            if role == "task_detail":
                slot.height = bg.top + bg.height - Pt(5) - slot.top
            else:
                slot.height = slots["task_detail"].top - Pt(1) - slot.top
            # Use real editable paragraphs with explicit line breaks, even in viewers
            # that ignore word_wrap on grouped imported text.
            run_sizes = [r.font.size.pt for p in tf.paragraphs for r in p.runs if r.font.size]
            size = max(run_sizes, default=14)
            font = preview_font(round(size / .75))
            width = (slot.width - tf.margin_left - tf.margin_right) / 9525 * .94
            lines = []
            for paragraph in tf.text.splitlines():
                line = ""
                tokens = re.findall(r"(?:[0-9]+(?:[,.][0-9]+)*(?:%|亿元|万吨|吨|家)?|[A-Za-z]+|.)[，。、；：！？）】]*", paragraph)
                for token in tokens:
                    if line and font.getlength(line + token) > width:
                        lines.append(line)
                        line = ""
                    line += token
                lines.append(line)
            _text(tf, lines)
            for p in tf.paragraphs:
                p.line_spacing = 1.05
                p.space_after = p.space_before = Pt(0)
                if role == "task_title":
                    for run in p.runs:
                        set_typeface(run, HEADING_FONT, HEADING_FONT)
            needed = Pt(len(lines) * size * 1.3) + tf.margin_top + tf.margin_bottom
            slot.height = min(slot.height, needed)
        detail = slots["task_detail"]
        title = slots["task_title"]
        detail.top = title.top + title.height + Pt(3)
        number = slots["task_status"]
        h = max(number.height + Pt(8), detail.top + detail.height + Pt(6) - bg.top)
        bg.height = badge.height = min(bg.height, h)
        number.top = bg.top + (bg.height - number.height) // 2
        unit.shapes._recalculate_extents()
    issues = task_issues(slide, palette)
    if issues:
        raise ValueError("T1-301 文字或底色检查未通过：" + str(issues))

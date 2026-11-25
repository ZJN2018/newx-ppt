"""可读性硬检查：按显式字号、字宽、换行、行距估算实际文字区域。"""
from pathlib import Path
import math
import re
from itertools import combinations
from pptx import Presentation
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN

# pt：正文/栏目标签12；页标题15（20px）；关键数字21（28px）；
# 来源8.25（11px）、页码9（12px）。装饰引号不属于可读正文。
FLOORS = {'body': 12, 'title': 15, 'number': 21, 'section': 12,
          'source': 8.25, 'page_number': 9, 'decoration': 0}
EMU_PT = 12700
EPS = .75  # EMU舍入和字体基线的1px余量，不允许用它容纳额外一行。


def role_of(text, x, y, size, bold=False, tagged=None):
    if tagged in FLOORS:
        return tagged
    if not re.search(r'[\w\u4e00-\u9fff]', text):
        return 'decoration'
    if y >= 510 and text.startswith('来源'):
        return 'source'
    if y >= 510 and re.fullmatch(r'\d{1,3}', text.strip()):
        return 'page_number'
    if y < 55 and size >= 15:
        return 'title'
    if y < 120 and bold and len(text.strip()) <= 12 and size <= 11.25:
        return 'section'
    if size >= 19.5 and re.fullmatch(r'[\d.,%+−—\-/\s亿元万家个件年月倍]+', text.strip()):
        return 'number'
    return 'body'


def _width(char, size):
    if char == '\t':
        return size * 2
    if char.isspace():
        return size * .3
    return size * (1 if ord(char) > 0x2e80 else .55)


def text_regions(frame, bounds, role=None, scale=(1, 1)):
    """返回逐行文字边界及字体下限/溢出问题；长词也会按实际可用宽度换行。"""
    x, y, w, h = bounds
    left, right, top, bottom = [float(getattr(frame, 'margin_' + side, 0) or 0) / EMU_PT
                                for side in ('left', 'right', 'top', 'bottom')]
    sx, sy = scale
    left, right, top, bottom = left*sx, right*sx, top*sy, bottom*sy
    available = max(.1, w - left - right)
    lines, errors = [], []
    current_y = 0
    for paragraph in frame.paragraphs:
        inherited = paragraph.font.size.pt if paragraph.font.size else 12
        runs = [(run.text, run.font.size.pt if run.font.size else inherited,
                 bool(run.font.bold)) for run in paragraph.runs]
        if not runs:
            runs = [(paragraph.text, 12, False)]
        spacing = paragraph.line_spacing
        before = paragraph.space_before.pt*sy if paragraph.space_before else 0
        after = paragraph.space_after.pt*sy if paragraph.space_after else 0
        current_y += before
        width, height, has_text = 0, 0, False

        def flush():
            nonlocal current_y, width, height, has_text
            if isinstance(spacing, float):
                line_h = height * spacing
            elif spacing is not None:
                line_h = max(height, spacing.pt*sy)
            else:
                line_h = height * 1.2
            line_h = max(line_h, height)
            if has_text:
                if width > available + EPS:
                    errors.append(("横向出框", f"文字宽{width:.2f}pt，可用{available:.2f}pt"))
                align = paragraph.alignment
                start = x + left
                if align == PP_ALIGN.CENTER:
                    start += (available - width) / 2
                elif align == PP_ALIGN.RIGHT:
                    start += available - width
                lines.append((start, current_y, width, height))
            current_y += line_h
            width, height, has_text = 0, 0, False

        for text, original_size, bold in runs:
            kind = role_of(text, x, y, original_size, bold, role)
            size = original_size * sy
            visible_size = original_size * min(sx, sy)
            if text.strip() and visible_size < FLOORS[kind] - .01:
                errors.append(('字号', '%s %.2fpt < %.2fpt' % (kind, visible_size, FLOORS[kind])))
            for char in text:
                if char in '\n\v\r':
                    height = max(height, size)
                    flush()
                    continue
                char_w = _width(char, original_size*sx)
                if frame.word_wrap is not False and width + char_w > available and has_text:
                    flush()
                width += char_w
                height = max(height, size)
                has_text = has_text or not char.isspace()
        if height:
            flush()
        current_y += after
    if current_y > h - top - bottom + EPS:
        errors.append(('出框', '预计文字高%.2fpt，可用%.2fpt' % (current_y, h - top - bottom)))
    offset = y + top
    if frame.vertical_anchor == MSO_ANCHOR.MIDDLE:
        offset += max(0, (h - top - bottom - current_y) / 2)
    elif frame.vertical_anchor == MSO_ANCHOR.BOTTOM:
        offset += max(0, h - top - bottom - current_y)
    return [(a, b + offset, c, d) for a, b, c, d in lines], errors


def _intersects(a, b):
    return (min(a[0] + a[2], b[0] + b[2]) - max(a[0], b[0]) > EPS and
            min(a[1] + a[3], b[1] + b[3]) - max(a[1], b[1]) > EPS)


def _contains(a, b):
    return a[0] - EPS <= b[0] and a[1] - EPS <= b[1] and \
        a[0] + a[2] + EPS >= b[0] + b[2] and a[1] + a[3] + EPS >= b[1] + b[3]


def _leaves(shapes, transform=(1, 1, 0, 0), owner=None):
    sx, sy, tx, ty = transform
    for shape in shapes:
        if shape.shape_type == 6:
            mine = owner or (shape.name if shape.name.startswith('creative:') else None)
            xf = shape._element.grpSpPr.xfrm
            gx = sx * xf.ext.cx / max(1, xf.chExt.cx)
            gy = sy * xf.ext.cy / max(1, xf.chExt.cy)
            yield from _leaves(shape.shapes, (gx, gy, tx + sx*xf.off.x - gx*xf.chOff.x,
                                              ty + sy*xf.off.y - gy*xf.chOff.y), mine)
        else:
            nodes = shape._element.xpath('.//p:cNvPr')
            if nodes and nodes[0].get('descr') in ('niuppt-background', 'niuma-background'):
                continue
            yield shape, ((tx+shape.left*sx)/EMU_PT, (ty+shape.top*sy)/EMU_PT,
                          shape.width*sx/EMU_PT, shape.height*sy/EMU_PT), (sx,sy), owner


def slide_issues(slide, slide_width, slide_height):
    errors, texts, graphics = [], [], []
    owners = {}
    for index, (shape, bounds, scale, owner) in enumerate(_leaves(slide.shapes)):
        owners[index] = owner
        frames = []
        if getattr(shape, 'has_chart', False):
            for node in shape.chart._element.xpath('.//a:defRPr[@sz] | .//a:rPr[@sz]'):
                size = int(node.get('sz')) / 100 * min(scale)
                if size < FLOORS['body'] - .01:
                    errors.append(('字号', shape.name + ': 图表标签 %.2fpt < 12pt' % size, '图表'))
        if getattr(shape, 'has_text_frame', False) and shape.text.strip():
            tag = shape._element.xpath('.//p:cNvPr')[0].get('descr', '')
            role = tag.split(':', 1)[1] if tag.startswith(('niuppt-role:', 'niuma-role:')) else None
            frames = [(shape.text_frame, bounds, role, shape.text[:32])]
        elif getattr(shape, 'has_table', False):
            y = bounds[1]
            for row in shape.table.rows:
                x = bounds[0]
                for column, cell in zip(shape.table.columns, row.cells):
                    if not cell.is_spanned:
                        frames.append((cell.text_frame, (x, y, column.width*scale[0] / EMU_PT,
                                                        row.height*scale[1] / EMU_PT), 'body', cell.text[:32]))
                    x += column.width*scale[0] / EMU_PT
                y += row.height*scale[1] / EMU_PT
        else:
            # 背景在文字之前且完全包住文字，是相关容器；后画的遮盖物不享此例外。
            fill = shape._element.xpath('.//a:solidFill | .//a:gradFill')
            alpha = shape._element.xpath('.//a:alpha')
            opaque = bool(fill) and all(int(a.get('val')) == 100000 for a in alpha)
            graphics.append((index, shape.name, bounds, opaque))
        for frame, rect, role, label in frames:
            if role_of(label, rect[0], rect[1], 12, tagged=role) == 'decoration':
                continue
            lines, local = text_regions(frame, rect, role, scale)
            if getattr(shape, 'rotation', 0):
                angle = math.radians(shape.rotation)
                cx, cy = rect[0] + rect[2] / 2, rect[1] + rect[3] / 2
                rotated = []
                for lx, ly, lw, lh in lines:
                    points = [(cx + (px-cx)*math.cos(angle) - (py-cy)*math.sin(angle),
                               cy + (px-cx)*math.sin(angle) + (py-cy)*math.cos(angle))
                              for px, py in ((lx,ly),(lx+lw,ly),(lx,ly+lh),(lx+lw,ly+lh))]
                    xs, ys = zip(*points)
                    rotated.append((min(xs),min(ys),max(xs)-min(xs),max(ys)-min(ys)))
                lines = rotated
            errors.extend((kind, shape.name + ': ' + detail, label) for kind, detail in local)
            for line in lines:
                if not _contains((0, 0, slide_width / EMU_PT, slide_height / EMU_PT), line):
                    errors.append(('出页', shape.name, label))
            texts.append((index, shape.name, lines, label))
    for a, b in combinations(texts, 2):
        if any(_intersects(x, z) for x in a[2] for z in b[2]):
            errors.append(('文字重叠', a[1] + ' / ' + b[1], a[3] + ' / ' + b[3]))
    for index, name, lines, label in texts:
        for graphic_index, graphic_name, rect, _ in graphics:
            if rect[2] < 1 or rect[3] < 1:
                continue
            # 已复核的创意组件里，画在文字之下的装饰属于原设计，不算压图；画在文字之上的仍要报。
            if owners.get(index) and owners.get(index) == owners.get(graphic_index) and graphic_index < index:
                continue
            # 比较可见区域：后画的不透明标题帽/数字圆完全盖住旧图形时，
            # 旧图形并未压住文字；不能把这种正常叠层误判为遮挡。
            if any(_intersects(line, rect) and not (graphic_index < index and _contains(rect, line))
                   and not any(graphic_index < cover_index < index and opaque and _contains(cover, line)
                               for cover_index, _, cover, opaque in graphics)
                   for line in lines):
                errors.append(('文字压图', name + ' / ' + graphic_name, label))
    # Only the repository build includes the optional creative unit inspector.
    if Path(__file__).with_name("creative_cards.py").is_file():
        from creative_cards import task_issues
        errors.extend(issue for issue in task_issues(slide) if issue[0] not in {"出框", "横向出框"})
    return list(dict.fromkeys(errors))


def hard_issues(path):
    prs = Presentation(path)
    return [(page, '[硬检查] ' + kind + ' ' + detail, label)
            for page, slide in enumerate(prs.slides, 1)
            for kind, detail, label in slide_issues(slide, prs.slide_width, prs.slide_height)]

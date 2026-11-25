"""OPC helpers for local creative assets; preserves editable charts and workbooks."""
from copy import deepcopy
import colorsys
from lxml import etree
from compose import compact_text
from pptx.chart.data import CategoryChartData, BubbleChartData, XyChartData
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.packuri import PackURI
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"

def shapes(slide):
    def walk(items):
        for sh in items:
            yield sh
            if sh.shape_type == 6:
                yield from walk(sh.shapes)
    return {sh.shape_id: sh for sh in walk(slide.shapes)}


def _text(tf, value):
    """Preserve original paragraph/run geometry and typography, including empty slots."""
    lines = value.split('\n') if isinstance(value, str) else list(value)
    lines = [compact_text(str(line)) for line in lines] or ['']
    if tf.text == '\n'.join(str(line) for line in lines):
        return
    paras = list(tf.paragraphs)
    while len(paras) < len(lines):
        paras[-1]._p.addnext(deepcopy(paras[-1]._p))
        paras = list(tf.paragraphs)
    for p, line in zip(paras, lines):
        # Numeric values often use a large first run and a smaller unit suffix.
        # Keep that style boundary when a replacement retains the same unit.
        runs = list(p.runs)
        if len(runs)>1 and runs[-1].text.strip() and str(line).endswith(runs[-1].text) and not any(c.isdigit() for c in runs[-1].text):
            suffix = runs[-1].text
            runs[0].text = str(line)[:-len(suffix)]
            for middle in runs[1:-1]:
                p._p.remove(middle._r)
            continue
        run = p.runs[0] if p.runs else p.add_run()
        run.text = str(line)
        for child in list(p._p):
            if child.tag in (f'{{{A}}}r', f'{{{A}}}br', f'{{{A}}}fld') and child is not run._r:
                p._p.remove(child)
    for p in paras[len(lines):]:
        p._p.getparent().remove(p._p)


def _chart(shape, data):
    kind = str(shape.chart.chart_type)
    if 'BUBBLE' in kind or 'XY_' in kind:
        cd = BubbleChartData() if 'BUBBLE' in kind else XyChartData()
        for name, points in data['series']:
            ser = cd.add_series(compact_text(name))
            for point in points:
                ser.add_data_point(*point)
    else:
        cd = CategoryChartData(number_format=data.get('number_format', 'General'))
        cd.categories = [compact_text(cat) for cat in data['categories']]
        for name, values in data['series']:
            if len(values) != len(data['categories']):
                raise ValueError('chart category/series lengths differ')
            cd.add_series(compact_text(name), values)
    shape.chart.replace_data(cd)
    if "legend" in data:
        from pptx.enum.chart import XL_LEGEND_POSITION
        # Explicit template plot boxes may cover the legend added for a new series.
        for layout in shape.chart._chartSpace.xpath(".//c:plotArea/c:layout | .//c:legend/c:layout"):
            layout.getparent().remove(layout)
        shape.chart.has_legend=bool(data["legend"])
        if shape.chart.has_legend:
            shape.chart.legend.position=XL_LEGEND_POSITION.BOTTOM
            shape.chart.legend.include_in_layout=False
    if "labels" in data:
        from pptx.enum.chart import XL_LABEL_POSITION
        plot=shape.chart.plots[0]
        plot.has_data_labels=bool(data["labels"])
        if plot.has_data_labels:
            plot.data_labels.show_value=True
            plot.data_labels.position=XL_LABEL_POSITION.OUTSIDE_END
            plot.data_labels.number_format=data.get("number_format","General")

    # Fixed bounds in examples are not portable to a different company's values.
    for tag in ('min', 'max'):
        for node in shape.chart._chartSpace.xpath(f'.//c:scaling/c:{tag}'):
            node.getparent().remove(node)
    for axis, bounds in data.get('bounds', {}).items():
        ax = getattr(shape.chart, axis)
        ax.minimum_scale, ax.maximum_scale = bounds


def _picture(slide, shape, path):
    # Replace only the image relationship. Preserve crop, z-order and geometry.
    part, rid = slide.part.get_or_add_image_part(str(path))
    blip = shape._element.find(f'.//{{{A}}}blip')
    blip.set(f'{{{R}}}embed', rid)


def _palette_color(value, pal):
    exact = {'005BAC': 'navy', '003A70': 'dark', '4A8BD0': 'navy2',
             'E8F0F9': 'tint', 'E39B2D': 'peach', '262B33': 'ink', '848B96': 'muted'}
    if value.upper() in exact:
        return pal.get(exact[value.upper()], value)
    try:
        h, light, sat = colorsys.rgb_to_hls(*[int(value[i:i+2], 16)/255 for i in (0, 2, 4)])
    except (ValueError, TypeError):
        return value
    if sat < .15:
        return pal.get('ink', value) if light < .28 else (pal.get('muted', value) if .38 < light < .62 else value)
    if 25 <= h*360 <= 65:
        return pal.get('peach', value)
    role = 'dark' if light < .22 else 'navy' if light < .42 else 'navy2' if light < .80 else 'tint'
    return pal.get(role, value)


def bind_component_colors(slide):
    """Detach shape colours from the source master before moving its group.

    Keep alpha/tint/luminance children on each colour and leave phClr alone
    (it is a style placeholder, not a theme swatch). The normal brand palette
    pass can then recolour these explicit source RGBs once, if requested.
    """
    layout = slide.slide_layout
    master = layout.slide_master
    theme = etree.fromstring(master.part.part_related_by(RT.THEME).blob)
    scheme = theme.find(f'.//{{{A}}}clrScheme')
    for owner in (layout, slide):
        for rel in owner.part.rels.values():
            if rel.reltype == RT.THEME_OVERRIDE:
                override = etree.fromstring(rel.target_part.blob).find(f'{{{A}}}clrScheme')
                if override is not None:
                    scheme = override
    colors = {}
    for swatch in scheme:
        value = swatch[0]
        rgb = value.get('lastClr') if value.tag == f'{{{A}}}sysClr' else value.get('val')
        if value.tag in (f'{{{A}}}srgbClr', f'{{{A}}}sysClr') and rgb:
            colors[etree.QName(swatch).localname] = rgb

    master_map = master._element.find(f'{{{P}}}clrMap').attrib
    color_map = master_map
    for owner in (layout, slide):
        override = owner._element.find(f'{{{P}}}clrMapOvr')
        if override is not None and len(override):
            mapping = override[0]
            color_map = mapping.attrib if mapping.tag == f'{{{A}}}overrideClrMapping' else master_map
    for node in slide.shapes._spTree.iter(f'{{{A}}}schemeClr'):
        slot = node.get('val')
        rgb = colors.get(color_map.get(slot, slot))
        if rgb:
            node.tag = f'{{{A}}}srgbClr'
            node.set('val', rgb)


def _clone_slide(dst, src, palette):
    """Copy the reachable OPC graph with typed parts and remapped relationships.

    Native chart caches AND editable embedded workbooks survive. A memo handles
    layout/master cycles; unique names prevent same-filename chart collisions.
    """
    package = dst.part.package
    taken = {str(p.partname) for p in package.iter_parts()}
    memo = {}

    def copy_part(part):
        if part in memo:
            return memo[part]
        old = str(part.partname)
        stem, ext = old.rsplit('.', 1)
        n = 1
        name = f'{stem}_creative{n}.{ext}'
        while name in taken:
            n += 1
            name = f'{stem}_creative{n}.{ext}'
        taken.add(name)
        new = type(part).load(PackURI(name), part.content_type, package, part.blob)
        memo[part] = new
        mappings = {}
        for rel in part.rels.values():
            if rel.reltype in (RT.NOTES_SLIDE, RT.NOTES_MASTER):
                continue
            target = rel.target_ref if rel.is_external else copy_part(rel.target_part)
            mappings[rel.rId] = new.relate_to(target, rel.reltype, is_external=rel.is_external)
        if part.content_type.endswith('+xml'):
            el = getattr(new, '_element', None)
            owned = el is None
            if owned:
                el = etree.fromstring(new.blob)
            for node in el.iter():
                for key, value in list(node.attrib.items()):
                    if key.startswith(f'{{{R}}}') and value in mappings:
                        node.set(key, mappings[value])
                if node.tag == f'{{{A}}}srgbClr':
                    node.set('val', _palette_color(node.get('val', ''), palette))
                elif node.tag == f'{{{A}}}schemeClr' and palette:
                    role = {'accent1':'navy','accent2':'navy2','accent3':'peach','accent4':'tint',
                            'accent5':'navy2','accent6':'peach','dk1':'ink','tx1':'ink','dk2':'dark',
                            'tx2':'dark','lt2':'tint','bg2':'tint'}.get(node.get('val'))
                    color = palette.get(role) if role else ('FFFFFF' if node.get('val') in ('lt1','bg1') else None)
                    if color:
                        node.tag=f'{{{A}}}srgbClr';node.set('val',color)
            for node in list(el.iter(f'{{{P}}}timing', f'{{{P}}}transition')):
                node.getparent().remove(node)
            if owned:
                new._blob = etree.tostring(el, xml_declaration=True, encoding='UTF-8', standalone=True)
        return new

    part = copy_part(src.part)
    # Consumers such as LibreOffice resolve bg1/tx1 and theme fonts via the
    # presentation's registered masters, not only the layout→master relationship.
    from pptx.parts.slide import SlideMasterPart
    masters = dst.part._element.get_or_add_sldMasterIdLst()
    for copied in memo.values():
        if isinstance(copied, SlideMasterPart):
            rid = dst.part.relate_to(copied, RT.SLIDE_MASTER)
            identifier = max([int(e.get('id')) for e in masters] + [2147483647]) + 1
            node = masters._add_sldMasterId(rId=rid)
            node.set('id', str(identifier))
    dst.slides._sldIdLst.add_sldId(dst.part.relate_to(part, RT.SLIDE))
    return dst.slides[-1]

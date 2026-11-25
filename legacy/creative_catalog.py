"""Unified discovery and safe placement of reviewed, local creative objects.

The repository contains metadata only. A missing or changed asset returns None
before touching the destination; the caller then draws its saved stable spec.
"""
from pathlib import Path
import hashlib
import json
import math
import os
import re
from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Pt
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
import layout_combo
from creative_assets import bind_component_colors, shapes, _text, _chart, _picture, _clone_slide, _palette_color, A, P, R

CATALOG=Path(__file__).resolve().parent/'catalog/creative-catalog.json'
ROLE_FALLBACK={'客户伙伴':'并列','案例图集':'并列','荣誉证书':'一组核心指标','地理布局':'并列','多维比较':'对比','层级':'层级','一组核心指标':'一组核心指标'}


CLASSIFIED = CATALOG.with_name("classified-objects.json")

# 组件自带的装饰字（如对比图中间的 V / S）：保留原样，不作为填字槽位，也不参与“文字压图”检查。
DECOR_TEXT = {"C1241": (14, 15), "B7-151": (535, 536, 537, 538)}  # B7-151：背景大字 S/W/O/T
# 主题色槽位 → 配色角色（与 creative_assets._clone_slide 的映射一致）
SCHEME_ROLE = {'accent1':'navy','accent2':'navy2','accent3':'peach','accent4':'tint','accent5':'navy2','accent6':'peach',
               'dk1':'ink','tx1':'ink','dk2':'dark','tx2':'dark','lt2':'tint','bg2':'tint'}

DEFAULT_DIVIDER = "C0485"

def catalog():
    originals = {item["code"]: item for item in json.loads(CATALOG.read_text())["items"]}
    classified = json.loads(CLASSIFIED.read_text())
    entries = []
    for code, group in classified.items():
        item = dict(originals.get(code, {}), **{"code":code,"classification":group})
        family = group["family"]
        item.setdefault("title",group["title"])
        item.setdefault("role",group["category"])
        if family in ("layout","reference") or group.get("rebuild"):
            modes = ["layout"]
        else:
            modes = ["component" if family in ("repeat","marker","header","divider","agenda") else "region"]
        item["modes"] = modes
        item["preferred_reuse"] = modes[0]
        entries.append(item)
    return entries

def entry(code):
    for item in catalog():
        if item["code"] == code:
            return item
    raise ValueError(f"未知创意编号 {code}")

def resolve(code, mode):
    item=entry(code)
    if mode not in item["modes"] or mode=="layout":
        raise ValueError(f"{code} 仅支持 {item['modes']}")
    group=item["classification"]
    if group.get("deferred"):
        raise ValueError(group["deferred"])
    if "file" in group:
        root=library_root("classified")
        asset=group
    else:
        root=library_root("creative")
        source_mode="page" if "body_ids" in group else mode
        asset=json.loads((root/"manifest.json").read_text())["items"][code][source_mode]
    path=(root/asset["file"]).resolve()
    if not path.is_relative_to(root.resolve()):
        raise ValueError("资产路径位于目录之外")
    if hashlib.sha256(path.read_bytes()).hexdigest()!=asset["sha256"]:
        raise ValueError(f"{code} 资产版本与复核记录不同")
    return path

def candidates(role=None, mode=None, family=None):
    results=[]
    for item in catalog():
        kind=item["preferred_reuse"]
        if (role and item["role"]!=role) or (mode and mode!=kind) or (family and item["classification"]["family"]!=family):
            continue
        try:
            if kind!="layout":
                resolve(item["code"],kind)
            item["available"]=True
        except (OSError,ValueError,KeyError) as exc:
            item["available"]=False
            item["unavailable_reasons"]={kind:str(exc)}
        item["availability"]={kind:item["available"]}
        item["fallback_role"]=ROLE_FALLBACK.get(item["role"],item["role"])
        results.append(item)
    missing=library_status()
    if missing and not _WARNED:
        _WARNED.append(1)
        print("[创意] 未找到素材库："+"；".join(missing)+"。依赖文件的组件均不可用，将改画基础图示。")
    return results

_WARNED=[]
_LIBS={"classified":("NIUPPT_COMPONENT_LIBRARY","分类图示库",None),
       "creative":("NIUPPT_CREATIVE_ASSETS","创意版式库（manifest.json）","manifest.json")}

def library_root(kind):
    """素材库位置：环境变量优先，其次是技能自带的脱敏库 skill/assets/library/<kind>。"""
    env,_,_=_LIBS[kind]
    if os.environ.get(env):
        return Path(os.environ[env])
    return Path(__file__).resolve().parents[1]/"assets"/"library"/kind

def library_status():
    """返回缺失的素材库说明；为空表示两个库都在。"""
    out=[]
    for kind,(env,label,probe) in _LIBS.items():
        root=library_root(kind)
        if not (root/probe if probe else root).exists():
            out.append(f"{label}，请设置 {env} 或放到 {root}")
    return out

def _selection(slide, item, mode, part):
    group=item["classification"]
    if "file" in group or "body_ids" in group:
        if part:
            raise ValueError("该条目已按已审阅的单元或功能区切分")
        return set(group["body_ids"]) if "body_ids" in group else {sh.shape_id for sh in slide.shapes}
    parts=item.get("parts",[])
    selected=[p for p in parts if p["name"]==part] if part else [p for p in parts if p["role"]=="unit"]
    if not selected or (mode=="component" and len(selected)!=1):
        raise ValueError("请选择唯一的已复核单元")
    return {p["group_id"] for p in selected}

def _prepare(path,item,mode,part):
    from prepare_creative_assets import clean, slots
    src=Presentation(path)
    slide=src.slides[0]
    keep=_selection(slide,item,mode,part)
    for sh in list(slide.shapes):
        if sh.shape_id not in keep:
            slide.shapes._spTree.remove(sh._element)
    if "file" in item["classification"]:
        clean(src)
        unit_id = item["classification"].get("unit_id")
        if unit_id:
            container = slide.shapes[0]
            unit = shapes(slide)[unit_id]
            # Keep the original group coordinate system while resetting its
            # bounding box to the single reviewed unit.
            for child in list(container.shapes):
                if child.shape_id != unit_id:
                    container.shapes._grpSp.remove(child._element)
            container.shapes._recalculate_extents()
    # The reviewed dual-line template was saved as stacked lines. The
    # classification compares two independent series on the same scale.
    if item["code"]=="C1171":
        for node in shapes(slide).values():
            if node.has_chart:
                for grouping in node.chart._chartSpace.xpath(".//c:lineChart/c:grouping"):
                    grouping.set("val","standard")
    schema=slots(slide)
    decor=[i for i in DECOR_TEXT.get(item["code"],()) if i in shapes(slide)]
    for i in decor:
        shapes(slide)[i]._element.xpath('.//p:cNvPr')[0].set('descr','niuppt-role:decoration')
    schema["text_ids"]=[i for i in schema["text_ids"] if i not in decor]
    schema["decor_text_ids"]=decor
    if "file" in item["classification"]:
        schema["content_image_ids"]=list(schema["image_ids"])
    else:
        original=item if "body_ids" in item["classification"] else item["cut_slots"]
        schema["content_image_ids"]=[i for i in original.get("content_image_ids",[]) if i in schema["image_ids"]]
    schema["entities"]=item.get("entities",[])
    return src,slide,schema

def inspect(code, mode=None, part=None):
    item=entry(code)
    mode=mode or item["preferred_reuse"]
    if mode not in item["modes"]:
        raise ValueError(f"{code} 仅支持 {item['modes']}")
    if mode=="layout":
        return dict(code=code,mode=mode,**item["classification"])
    _,slide,schema=_prepare(resolve(code,mode),item,mode,part)
    d=shapes(slide)
    return dict(code=code,mode=mode,part=part,repeatable=item["classification"]["family"]=="repeat",texts={i:d[i].text for i in schema["text_ids"]},
                charts={int(i):dict(type=str(d[int(i)].chart.chart_type),series=[v.name for v in d[int(i)].chart.series]) for i in schema["chart_slots"]},
                tables={int(i):[[c.text for c in row.cells] for row in d[int(i)].table.rows] for i in schema["table_slots"]},
                images=schema["image_ids"], required_images=schema["content_image_ids"],
                decor_texts={i:d[i].text for i in schema.get("decor_text_ids",[])},
                slots=_slot_info(slide,schema["text_ids"]), vector_icons=_icon_info(slide,schema["text_ids"]))

def _abs_boxes(slide):
    """所有叶子形状在页面上的绝对位置（pt）。"""
    out={}
    def visit(shapes,t=(1,1,0,0)):
        sx,sy,tx,ty=t
        for x in shapes:
            if x.shape_type==6:
                xf=x._element.grpSpPr.xfrm
                gx=sx*xf.ext.cx/max(1,xf.chExt.cx); gy=sy*xf.ext.cy/max(1,xf.chExt.cy)
                out[x.shape_id]=((tx+x.left*sx)/12700,(ty+x.top*sy)/12700,x.width*sx/12700,x.height*sy/12700)
                visit(x.shapes,(gx,gy,tx+sx*xf.off.x-gx*xf.chOff.x,ty+sy*xf.off.y-gy*xf.chOff.y))
            else:
                out[x.shape_id]=((tx+x.left*sx)/12700,(ty+x.top*sy)/12700,x.width*sx/12700,x.height*sy/12700)
    visit(slide.shapes)
    return out

def _slot_info(slide,text_ids):
    """每个文字槽位的角色（取自形状名，如 parallel_heading）、示例文字与位置，便于判断该填什么。"""
    d=shapes(slide); boxes=_abs_boxes(slide); out={}
    for i in text_ids:
        sh=d[i]; parts=[p.strip() for p in sh.name.split('|')]
        role=parts[1] if len(parts)>=3 else ''
        x,y,w,h=boxes.get(i,(0,0,0,0))
        out[i]=dict(role=role,text=sh.text[:30],x=round(x),y=round(y),w=round(w),h=round(h))
    return out

def _icon_info(slide,text_ids):
    """矢量小图标的替换顺序（icons=[...] 按此顺序对应）以及离它最近的文字槽位。"""
    d=shapes(slide); boxes=_abs_boxes(slide); out=[]
    for k,g in enumerate(_vector_icons(slide)):
        x,y,w,h=boxes[g.shape_id]
        near=min(text_ids,key=lambda i:(boxes[i][0]-x)**2+(boxes[i][1]-y)**2) if text_ids else None
        out.append(dict(order=k+1,x=round(x),y=round(y),near_slot=near,near_text=d[near].text[:12] if near else ''))
    return out

def _walk(items):
    for s in items:
        yield s
        if s.shape_type==6:yield from _walk(s.shapes)


def _fill(slide,item,texts,charts,tables,images,logo,source):
    d=shapes(slide)
    maps=[{int(k):v for k,v in (m or {}).items()} for m in (texts,charts,tables,images)]
    texts,charts,tables,images=maps
    for label,expected,supplied in [('texts',set(item['text_ids'])&set(d),set(texts)),
                                    ('charts',set(map(int,item['chart_slots']))&set(d),set(charts)),
                                    ('tables',set(map(int,item['table_slots']))&set(d),set(tables))]:
        if expected!=supplied:raise ValueError(f'{label}: missing {expected-supplied}; unknown {supplied-expected}')
    needed=set(item.get('content_image_ids',[]))&set(d)
    if not needed<=set(images):raise ValueError(f'Replace content images: {needed-set(images)}')
    if set(images)-(set(item['image_ids'])&set(d)):raise ValueError('unknown image slot')
    for i in set(item.get('entities',[]))&set(d):
        if not str(texts[i]).strip():raise ValueError('entity slots require real, nonempty names')
    for i,value in texts.items():
        frame=d[i].text_frame
        widest=max((len(line) for line in frame.text.split('\n')),default=0)
        _text(frame,value)
        # 模板里不换行的文本框，填入更长的行后改为自动换行，避免整行越出组件。
        body=frame._txBody.find(f'{{{A}}}bodyPr')
        if body is not None and body.get('wrap')=='none' and max(len(l) for l in str(value).split('\n'))>widest:
            body.set('wrap','square')
    for i in item.get('source_text_ids',[]):
        if i in d:_text(d[i].text_frame,source)
    for i,data in charts.items():_chart(d[i],data)
    for i,data in tables.items():
        table=d[i].table
        if len(data)!=len(table.rows) or any(len(r)!=len(table.columns) for r in data):raise ValueError(f'table {i}: shape mismatch')
        for r,row in enumerate(data):
            for c,value in enumerate(row):_text(table.cell(r,c).text_frame,value)
    for i,data in images.items():
        if not isinstance(data,dict) or not data.get('source') or not data.get('meaning'):raise ValueError('image needs path, meaning, source')
        _picture(slide,d[i],data['path'])
    for i in set(item.get('logo_ids',[]))&set(d):
        if not logo:raise ValueError('page needs a replacement logo')
        _picture(slide,d[i],logo)
    used={v for el in slide._element.iter() for k,v in el.attrib.items() if k.startswith(f'{{{R}}}')}
    for rel in list(slide.part.rels.values()):
        if rel.reltype in (RT.IMAGE,RT.CHART) and rel.rId not in used:slide.part.drop_rel(rel.rId)


def place(prs,palette,code,*,mode=None,slide=None,area=None,part=None,texts=None,charts=None,tables=None,images=None,icons=None,logo=None,source='',reason='',semantic_role=None,min_text_pt=None,measure_only=False):
    item=entry(code)
    mode=mode or item['preferred_reuse']
    if mode=='layout':raise ValueError(f'{code} 是分区或拼法条目，不能直接放置或试放；用 S.creative_layout("{code}", section, title, regions, source=来源) 借它的分区在新页面上重建，regions 的写法与 S.content_page 相同')
    if mode not in item['modes']:raise ValueError(f'{code}: unsupported mode {mode}')
    try:path=resolve(code,mode)
    except (OSError,ValueError,KeyError) as exc:
        return None,dict(code=code,status='unavailable',reason=str(exc),fallback_role=ROLE_FALLBACK.get(item['role'],item['role']))
    if not source.strip() or not reason.strip():raise ValueError('source and content-based selection reason required')
    if semantic_role!=item['role']:
        raise ValueError(f'{code} 的 semantic_role 必须逐字写成 {item["role"]!r}（当前传入 {semantic_role!r}）；'
                         f'它用来确认你按材料关系选了这个组件，可用 creative_catalog.entry("{code}")["role"] 查到')
    if min_text_pt is not None:
        if not math.isfinite(min_text_pt) or not 1 <= min_text_pt <= 400:
            raise ValueError('min_text_pt requires a component/region and a size between 1 and 400 pt')
    if slide is None or area is None or len(area)!=4:
        raise ValueError("区域调用需要目标 slide 和 area")
    if not any(s.part is slide.part for s in prs.slides):
        raise ValueError("目标页属于另一份演示")
    template,src,slots=_prepare(path,item,mode,part)
    x,y,w,h=map(float,area)
    if not all(math.isfinite(v) for v in (x,y,w,h)) or w<=0 or h<=0 or x<0 or y<0 or x+w>1280.01 or y+h>720.01:raise ValueError('area must fit the 1280×720 canvas')
    if images:
        # {"icon": 关键词} 换成按主色渲染的图标库 PNG
        import icons as _icon_lib
        images={k:(dict(path=_icon_lib.render(v["icon"],(palette or {}).get("navy","1F4E79")),meaning=v.get("meaning",v["icon"]),source="Lucide 图标（ISC）")
                   if isinstance(v,dict) and v.get("icon") else v) for k,v in images.items()}
    icon_slots={i for i in slots.get("image_ids",[]) if i not in set(map(int,(images or {}).keys()))}
    if texts and slots.get("decor_text_ids"):
        texts={k:v for k,v in texts.items() if int(k) not in slots["decor_text_ids"]}   # 装饰字保留原样，旧稿传入的值忽略
    _fill(src,slots,texts,charts,tables,images,logo,source)
    if code=="T1-277":
        _tune_t1_277(src)
    if code=="C1241" and palette and not palette.get("alt"):
        _contrast_right_side(src)
    if icon_slots:
        print(f"[创意] {code} 有 {len(icon_slots)} 个图标槽位未替换（{sorted(icon_slots)}），仍是模板原图标；请用 images={{槽位: {{'icon': 关键词}}}} 换成贴合内容的图标")
    vec=_vector_icons(src)
    if icons:
        if len(icons)!=len(vec):raise ValueError(f'{code} 有 {len(vec)} 个矢量图标（按从上到下、从左到右），icons 给了 {len(icons)} 个')
        import icons as _icons
        mapping=[]
        for k,(group,query) in enumerate(zip(vec,icons)):
            # 图标压在深色块上时用白色，否则用主色，保证看得见
            _swap_icon(src,group,_icons.render(query,_icon_color(src,group,palette)))
            mapping.append(f"{k+1}:{query}→{_icons.pick(query)}")
        print(f"[创意] {code} 图标（从上到下、从左到右）："+"，".join(mapping))
    elif vec:
        print(f"[创意] {code} 含 {len(vec)} 个模板矢量图标，未按内容替换；请用 icons=[关键词, ...]（从上到下、从左到右）换成贴合内容的图标")
    if item["classification"]["family"] == "marker":
        for shape in _walk(src.shapes):
            if shape.has_text_frame and shape.text.strip():
                shape._element.xpath('.//p:cNvPr')[0].set('descr', 'niuppt-role:section')
    bind_component_colors(src)
    # These reviewed slots are headings even when the body text is identical.
    heading_slots = {"C1933": 8, "C1949": 13, "C1952": 13, "C1309": 2}
    if code in heading_slots:
        from compose import HEADING_FONT, set_typeface
        heading = shapes(src)[heading_slots[code]]
        for paragraph in heading.text_frame.paragraphs:
            for run in paragraph.runs:
                set_typeface(run, HEADING_FONT, HEADING_FONT)
        if code in ("C1309", "C1933"):
            from deco import grad_fill
            grad_fill(heading, [(0, palette["dark"], None), (100, palette["navy"], None)], 0, large_block=True)
            for paragraph in heading.text_frame.paragraphs:
                for run in paragraph.runs:
                    run.font.color.rgb = RGBColor.from_string("FFFFFF")
    # Some source units have 10 pt captions. An actual grid may opt into
    # a larger floor after fitting its copy; preserve every larger run.
    if min_text_pt is not None:
        floor = round(min_text_pt * 100)
        for run in src.shapes._spTree.iter(f'{{{A}}}rPr', f'{{{A}}}defRPr', f'{{{A}}}endParaRPr'):
            if run.get('sz') is not None and int(run.get('sz')) < floor:
                run.set('sz', str(floor))
    if code == "C1933" and mode == "component":
        from skin_solid import card_text_height
        from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE
        slots = shapes(src)
        title, detail = slots[8], slots[7]
        for shape, size in ((title, 21), (detail, 18)):
            tf = shape.text_frame
            tf.word_wrap = True
            tf.auto_size = MSO_AUTO_SIZE.NONE
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf.margin_top = tf.margin_bottom = Pt(4.5)
            width = (shape.width - tf.margin_left - tf.margin_right) / 9525
            size = max(size, (min_text_pt or 0) / .75)
            height = card_text_height(tf.text, width, size, 1.2)
            shape.height = round(height * 9525) + tf.margin_top + tf.margin_bottom
            for paragraph in tf.paragraphs:
                paragraph.line_spacing = 1.2
                for run in paragraph.runs:
                    run.font.size = Pt(size * .75)
        detail.top = title.top + title.height
        for group in reversed(list(_walk(src.shapes))):
            if group.shape_type == 6:
                group.shapes._recalculate_extents()
    if code == "T1-301" and mode == "component" and palette:
        from creative_cards import fit_task_units
        fit_task_units(src, palette)
    # Measure the filled, reflowed unit, including nested group transforms.
    # Empty template slots have no visible text and impose no scale floor.
    from layout_checks import _leaves
    gw = max(s.left+s.width for s in src.shapes)-min(s.left for s in src.shapes)
    gh = max(s.top+s.height for s in src.shapes)-min(s.top for s in src.shapes)
    scale = min(w*9525/max(1,gw), h*9525/max(1,gh))
    fill = None
    if item["classification"]["family"] in ("single","repeat"):
        fill = _fit_body_slots(src, slots, scale, item["classification"]["family"])
    effective = []
    for shape, _, (sx,sy), _ in _leaves(src.shapes):
        # Some renderers resize a group without enlarging its editable fonts.
        # A placement must satisfy both the group transform and the stored size.
        local = min(1.0, scale*min(sx,sy))
        frames = [shape.text_frame] if shape.has_text_frame else []
        if shape.has_table:
            frames.extend(c.text_frame for c in shape.table.iter_cells() if not c.is_spanned)
        for frame in frames:
            for para in frame.paragraphs:
                for run in para.runs:
                    if run.text.strip():
                        size = run.font.size or para.font.size
                        effective.append((size.pt if size else 12)*local)
        if shape.has_chart:
            sizes = shape.chart._element.xpath('.//a:defRPr[@sz] | .//a:rPr[@sz]')
            effective.extend((int(node.get('sz'))/100)*local for node in sizes)
            if not sizes: effective.append(12*local)
    minimum = min(effective, default=None)
    floor = min_text_pt or 12  # v2：允许调用方在放不下时按更小下限试放（runner 会提醒）
    if minimum is not None and minimum < floor-.01:
        # 字号随缩放线性变化：按比例反推能放下当前内容的最小区域
        k = floor / minimum
        need_w, need_h_ = gw * scale / 9525 * k, gh * scale / 9525 * k
        raise ValueError(f"{code} 放置后最小字号 {minimum:.2f}pt，低于 {floor:g}pt；"
                         f"按当前内容至少需要约 {need_w:.0f}×{need_h_:.0f}px 的区域（现为 {w:.0f}×{h:.0f}），或换组件")
    if measure_only:
        return None, dict(code=code, status="fits", scale=scale, min_visible_pt=minimum, text_fill=fill)
    # Clone only after all user input and geometry checks succeeded.
    copied=_clone_slide(prs,src,palette)
    # Group first, remap all IDs (including internal connector endpoints),
    # then move the OPC relationships from the temporary copied slide.
    group=copied.shapes.add_group_shape(list(copied.shapes))
    unify_component(copied,group,palette)
    oldw,oldh=group.width,group.height
    group.left=round((x+(w-oldw*scale/9525)/2)*9525);group.top=round((y+(h-oldh*scale/9525)/2)*9525)
    group.width=round(oldw*scale);group.height=round(oldh*scale)
    # Group transforms move/resize text boxes, not their point sizes.
    # Retain run properties; dividing sizes by scale makes enlarged units
    # illegible in LibreOffice and changes their editable font sizes.
    nextid=max((int(n.get('id')) for n in slide._element.iter(f'{{{P}}}cNvPr')),default=0)+1;ids={}
    for node in group._element.iter(f'{{{P}}}cNvPr'):
        old=node.get('id');ids[old]=str(nextid);node.set('id',str(nextid));nextid+=1
    rels={}
    for el in group._element.iter():
        if el.tag in (f'{{{A}}}stCxn',f'{{{A}}}endCxn') and el.get('id') in ids:el.set('id',ids[el.get('id')])
        for key,value in list(el.attrib.items()):
            if key.startswith(f'{{{R}}}'):
                if value not in rels:
                    rel=copied.part.rels[value]
                    rels[value]=slide.part.relate_to(rel.target_ref if rel.is_external else rel.target_part,rel.reltype,is_external=rel.is_external)
                el.set(key,rels[value])
    group.name=f'creative:{code}:{part or mode}'
    slide.shapes._spTree.insert_element_before(group._element,'p:extLst')
    sid=prs.slides._sldIdLst[-1];prs.part.drop_rel(sid.rId);prs.slides._sldIdLst.remove(sid)
    result=slide
    return result,dict(code=code,status='placed',min_visible_pt=minimum,scale=scale,mode=mode,part=part,role=item['role'],source=source,reason=reason,
                       text_fill=fill,group_name=group.name,sha256=hashlib.sha256(path.read_bytes()).hexdigest())


_NUMERIC = re.compile(r"[\d.,%+−—\-/\s亿元万家个件年月倍]+")


def _fit_body_slots(slide, schema, scale, family):
    """组件的正文槽位按实际文字适配，返回正文占文字区的比例（没有正文槽位时为 None）。

    组件放进功能区时只缩放图形，字号不变，短文字会悬在一大块文字区中间。这里对正文槽位：
    去掉没有项目符号的悬挂缩进（否则第二行起错位）、改为自动换行；
    字号在模板字号到 18pt 之间取能放下的最大值（比标题小 2pt 以上）；
    可复制单元里左对齐的正文改为顶对齐，并列单元的正文起始高度一致。"""
    d = shapes(slide); boxes = _abs_boxes(slide)
    found = []
    for i in schema["text_ids"]:
        sh = d.get(i)
        if sh is None or not sh.has_text_frame or not sh.text_frame.text.strip():
            continue
        sizes = [r.font.size.pt for p in sh.text_frame.paragraphs for r in p.runs if r.text.strip() and r.font.size]
        found.append((i, sh, max(sizes) if sizes else None))
    known = sorted({z for _, _, z in found if z})
    if not known:
        return None
    heading = known[-1] if len(known) > 1 else None
    from skin_solid import card_text_height
    from pptx.util import Pt as _Pt
    ratios = []
    for i, sh, size in found:
        value = sh.text_frame.text
        if not size or size > 16 or size == heading or _NUMERIC.fullmatch(value.strip()):
            continue
        if _fill_colour(sh, {}) is not None:      # 有底色的是标签或标题块，不是正文
            continue
        x, y, bw, bh = boxes[i]
        if bh < size * 2.4:          # 单行标签，不是正文
            continue
        body = sh.text_frame._txBody.find(f'{{{A}}}bodyPr')
        ins = [int(body.get(k, dflt)) / 12700 for k, dflt in (("lIns", 91440), ("rIns", 91440), ("tIns", 45720), ("bIns", 45720))]
        avail_w = (bw * scale - ins[0] - ins[1]) / .75
        avail_h = (bh * scale - ins[2] - ins[3]) / .75
        if avail_w <= 20 or avail_h <= 20:
            continue
        body.set("wrap", "square")
        for auto in body.findall(f'{{{A}}}spAutoFit') + body.findall(f'{{{A}}}normAutofit'):
            body.remove(auto)
        if body.find(f'{{{A}}}noAutofit') is None:
            etree.SubElement(body, f'{{{A}}}noAutofit')
        left = True
        for para in sh.text_frame.paragraphs:
            ppr = para._p.find(f'{{{A}}}pPr')
            if ppr is None:
                continue
            if ppr.get("algn") in ("ctr", "r"):
                left = False
            if not any(ppr.find(f'{{{A}}}{tag}') is not None for tag in ("buChar", "buAutoNum", "buBlip")):
                ppr.attrib.pop("marL", None); ppr.attrib.pop("indent", None)
        spacing = max([p.line_spacing for p in sh.text_frame.paragraphs if isinstance(p.line_spacing, float)] or [1.2])
        cap = min(size * 1.5, 18 if heading else 16, (heading - 2) if heading else 99)
        chosen = size
        for pt in range(int(cap), int(size), -1):
            if card_text_height(value, avail_w, pt / .75, spacing) <= avail_h * .96:
                chosen = pt
                break
        if chosen != size:
            for para in sh.text_frame.paragraphs:
                for run in para.runs:
                    run.font.size = _Pt(chosen)
                end = para._p.find(f'{{{A}}}endParaRPr')
                if end is not None and end.get("sz"):
                    end.set("sz", str(int(chosen * 100)))
        if family == "repeat" and left and body.get("anchor") == "ctr":
            body.set("anchor", "t")
        ratios.append(card_text_height(value, avail_w, chosen / .75, spacing) / avail_h)
    return round(min(ratios), 2) if ratios else None


def _tune_t1_277(slide):
    """五列能力：模板图标只有 16pt，放大到 24pt 并把标题下移，图标与标题都落在标题帽内。"""
    d = shapes(slide)
    if not all(k in d for k in (2861, 2862, 2888)):
        return
    cap, heading, icon = d[2861], d[2862], d[2888]
    side = 24 * 12700
    centre = icon.left + icon.width // 2
    icon.left, icon.top, icon.width, icon.height = centre - side // 2, cap.top + 5 * 12700, side, side
    heading.top = icon.top + side + 12700
    heading.height = cap.top + cap.height - 3 * 12700 - heading.top


def _contrast_right_side(slide):
    """VS 对比：右栏标题与右侧字母改用强调色，左右两栏一眼分得开（原模板两栏是同一色系的深浅）。"""
    d = shapes(slide)
    for i in (9, 14):
        if i not in d:
            continue
        for node in d[i]._element.iter(f'{{{A}}}schemeClr'):
            if node.get("val") in ("accent2", "accent5"):
                node.tag = f'{{{A}}}srgbClr'; node.set("val", "E39B2D")   # 模板里的强调色，复制时映射到当前强调色


def _fill_colour(shape, palette):
    """形状复制到成稿后的填充色；没有填充返回 None。"""
    sp = shape._element.find(f'{{{P}}}spPr')
    if sp is None:
        return None
    for child in sp:
        tag = etree.QName(child).localname
        if tag == "noFill":
            return None
        if tag in ("solidFill", "gradFill"):
            node = next(child.iter(f'{{{A}}}srgbClr', f'{{{A}}}schemeClr'), None)
            if node is None:
                return None
            if node.tag.endswith("srgbClr"):
                return _palette_color(node.get("val", ""), palette or {})
            role = SCHEME_ROLE.get(node.get("val"))
            return (palette or {}).get(role) if role else ("FFFFFF" if node.get("val") in ("lt1", "bg1") else None)
    return None


def _icon_color(slide, icon, palette):
    """图标颜色：衬在深色块上用白色，否则用主色。"""
    main = (palette or {}).get("navy", "1F4E79")
    boxes = _abs_boxes(slide)
    if icon.shape_id not in boxes:
        return main
    x, y, w, h = boxes[icon.shape_id]
    cx, cy = x + w / 2, y + h / 2
    backdrop = None
    def leaves(items):
        for s in items:
            if s.shape_type == 6:
                if s.shape_id == icon.shape_id:
                    yield s
                else:
                    yield from leaves(s.shapes)
            else:
                yield s
    for s in leaves(slide.shapes):
        if s.shape_id == icon.shape_id:
            break
        bx = boxes.get(s.shape_id)
        if bx and bx[0] <= cx <= bx[0] + bx[2] and bx[1] <= cy <= bx[1] + bx[3]:
            colour = _fill_colour(s, palette)
            if colour:
                backdrop = colour
    if not backdrop or not re.fullmatch(r"[0-9A-Fa-f]{6}", backdrop):
        return main
    r, g, b = (int(backdrop[i:i + 2], 16) / 255 for i in (0, 2, 4))
    return "FFFFFF" if 0.2126 * r + 0.7152 * g + 0.0722 * b < 0.55 else main


def harmonize_units(groups):
    """同一页上重复放置的同一种单元：同名文字槽位统一取最小字号，避免并列单元字号不一。"""
    table = {}
    for g in groups:
        for sp in g.iter(f'{{{P}}}sp'):
            name = sp.find(f'.//{{{P}}}cNvPr').get("name", "")
            sizes = [int(r.get("sz")) for r in sp.iter(f'{{{A}}}rPr') if r.get("sz")]
            if name and sizes:
                table.setdefault(name, []).append((sp, max(sizes)))
    for name, entries in table.items():
        if len(entries) != len(groups):
            continue
        low = min(size for _, size in entries)
        for sp, size in entries:
            if size != low:
                for r in sp.iter(f'{{{A}}}rPr', f'{{{A}}}endParaRPr'):
                    if r.get("sz"):
                        r.set("sz", str(low))


def _vector_icons(slide, limit=72*12700):
    """模板里由若干自由曲线拼成的小图标组合，按页面上的绝对位置从上到下、从左到右排序（同一行按 20pt 归并）。"""
    found=[]
    def leaves(g):
        for x in g.shapes:
            if x.shape_type==6: yield from leaves(x)
            else: yield x
    def lone_icon(x, w, h):
        # 单个自由曲线画成的小图标（如 T1-277 标题上方的放大镜）：无文字、近似方形、12–72pt
        if x.shape_type!=5 or (x.has_text_frame and x.text_frame.text.strip()): return False
        return 12*12700<=min(w,h) and max(w,h)<=limit and max(w,h)<=2*min(w,h)
    def visit(shapes, t=(1,1,0,0)):
        sx,sy,tx,ty=t
        for x in shapes:
            if x.shape_type!=6:
                if lone_icon(x, x.width*sx, x.height*sy):
                    found.append((ty+x.top*sy, tx+x.left*sx, x))
                continue
            xf=x._element.grpSpPr.xfrm
            kids=list(leaves(x))
            w,h=x.width*sx,x.height*sy
            if kids and all(k.shape_type==5 for k in kids) and max(w,h)<=limit:
                found.append((ty+x.top*sy, tx+x.left*sx, x))
            else:
                gx=sx*xf.ext.cx/max(1,xf.chExt.cx); gy=sy*xf.ext.cy/max(1,xf.chExt.cy)
                visit(x.shapes,(gx,gy,tx+sx*xf.off.x-gx*xf.chOff.x,ty+sy*xf.off.y-gy*xf.chOff.y))
    visit(slide.shapes)
    found.sort(key=lambda f:(round(f[0]/(20*12700)),f[1]))
    return [f[2] for f in found]

def _icon_positions(slide):
    """供 creative_inspect 使用：每个矢量图标的序号与左上角位置（pt）。"""
    out=[]
    for i,g in enumerate(_vector_icons(slide)):
        out.append(dict(index=i, name=g.name))
    return out

def _swap_icon(slide, group, png):
    """把矢量图标组合的内容换成图标库 PNG，保留原位置与大小。"""
    from pptx.oxml.ns import qn
    if group.shape_type!=6:
        # 单个自由曲线图标：原位换成同尺寸居中的图标图片
        sp=group._element
        xf=sp.find(qn('p:spPr')).find(qn('a:xfrm'))
        x0,y0=int(xf.find(qn('a:off')).get('x')),int(xf.find(qn('a:off')).get('y'))
        cw,ch=int(xf.find(qn('a:ext')).get('cx')),int(xf.find(qn('a:ext')).get('cy'))
        side=min(cw,ch)
        pic=slide.shapes.add_picture(png,0,0,side,side)
        el=pic._element; el.getparent().remove(el)
        off=el.find(qn('p:spPr')).find(qn('a:xfrm')).find(qn('a:off'))
        off.set('x',str(x0+(cw-side)//2)); off.set('y',str(y0+(ch-side)//2))
        sp.addprevious(el); sp.getparent().remove(sp)
        pic.name=f"icon {Path(png).stem}"
        return
    xfrm=group._element.find(qn('p:grpSpPr')).find(qn('a:xfrm'))
    choff,chext=xfrm.find(qn('a:chOff')),xfrm.find(qn('a:chExt'))
    cx,cy,cw,ch=(int(choff.get('x')),int(choff.get('y')),int(chext.get('cx')),int(chext.get('cy')))
    side=min(cw,ch)
    pic=slide.shapes.add_picture(png,0,0,side,side)
    el=pic._element; el.getparent().remove(el)
    for child in list(group._element):
        if child.tag in (qn('p:sp'),qn('p:grpSp'),qn('p:pic'),qn('p:cxnSp')):group._element.remove(child)
    group._element.append(el)
    off=el.find(qn('p:spPr')).find(qn('a:xfrm')).find(qn('a:off'))
    off.set('x',str(cx+(cw-side)//2)); off.set('y',str(cy+(ch-side)//2))
    pic.name=f"icon {Path(png).stem}"

def unify_component(slide, group, palette):
    """统一放入组件的字体与图标颜色：正文用整套正文字体，标题类槽位用标题字体；位图图标改为主色。"""
    from compose import FONT_CN, FONT_EN, HEADING_FONT
    import io
    from lxml import etree
    from PIL import Image
    def walk(g):
        for sh in g.shapes:
            if sh.shape_type == 6: yield from walk(sh)
            else: yield sh
    heading_words = ("heading", "title", "topic", "header")
    for sh in walk(group):
        heading = any(w in sh.name.lower() for w in heading_words)
        for rpr in sh._element.iter(f'{{{A}}}rPr', f'{{{A}}}defRPr', f'{{{A}}}endParaRPr'):
            for tag, face in (("latin", HEADING_FONT if heading else FONT_EN), ("ea", HEADING_FONT if heading else FONT_CN), ("cs", FONT_EN)):
                node = rpr.find(f'{{{A}}}{tag}')
                if node is None:
                    node = etree.SubElement(rpr, f'{{{A}}}{tag}')
                    # a:latin / a:ea / a:cs 须排在 a:hlinkClick 等之前
                    for later in ("hlinkClick", "hlinkMouseOver", "rtl", "extLst"):
                        ref = rpr.find(f'{{{A}}}{later}')
                        if ref is not None: ref.addprevious(node); break
                node.set("typeface", face)
                for k in ("panose", "pitchFamily", "charset"):
                    node.attrib.pop(k, None)
        if sh.shape_type == 13 and palette and "icon" in sh.name.lower():
            blip = sh._element.xpath('.//a:blip')[0]
            part = slide.part.related_part(blip.get(f'{{{R}}}embed'))
            im = Image.open(io.BytesIO(part.blob)).convert("RGBA")
            main = tuple(int(palette["navy"][i:i+2], 16) for i in (0, 2, 4))
            solid = Image.new("RGBA", im.size, main + (0,)); solid.putalpha(im.getchannel("A"))
            buf = io.BytesIO(); solid.save(buf, "PNG"); part._blob = buf.getvalue()
            for ext in sh._element.xpath('.//a:extLst'):
                for e in list(ext):
                    if 'svg' in etree.tostring(e).decode()[:400].lower(): ext.remove(e)


DIVIDER_SLOTS = {
    "C0480": {"number": 8, "title": 6, "note": 7},
    "C0485": {"number": 3, "title": 8, "tag": 5, "note": 7},
    "C0504": {"number": 3, "title": 5, "note": 7},
    "C0520": {"number": 3, "title": 7, "note": 8},
    "C0526": {"number": 3, "title": 4, "tag": 5, "note": 7},
}

def divider_texts(code, no, title, sub, note):
    """把章号、章名、目录说明和章节摘要填进库内转场的槽位；没有摘要时用目录说明。"""
    slots = DIVIDER_SLOTS.get(code)
    if not slots:
        raise ValueError(f"{code} 尚未登记转场槽位")
    cn = "零一二三四五六七八九十"
    tag = f"第{cn[int(no)]}部分" if 0 < int(no) <= 10 else f"第{int(no)}部分"
    # 没有摘要也没有目录说明时摘要槽留空，不再把章名再写一遍（会与章名框重叠）
    values = {"number": f"{int(no):02d}", "title": title, "tag": tag, "note": note or sub or ""}
    return {slot: values[key] for key, slot in slots.items()}


def check_creative_limit(log):
    """Count content pages in reading order, once per page even for a grid."""
    frames = {n: frame for n, kind, frame in log if kind == 'frame'}
    content = sorted(n for n, frame in frames.items()
                     if frame in layout_combo.CONTENT_FRAMES or frame == 'creative')
    used = {n for n, kind, value in log if kind.startswith('creative:')
            or (kind == 'frame' and value == 'creative')}
    pages = [n for n in content if n in used]
    limit = len(content) // 3
    if len(pages) > limit:
        raise ValueError(f'创意页 {pages}，共 {len(content)} 张内容页，最多使用 {limit} 张创意页')
    for i in range(2, len(content)):
        triple = content[i - 2:i + 1]
        if all(n in used for n in triple):
            raise ValueError(f'连续三张内容页 {triple} 使用创意版式，最多连续两张')


def attach_catalog(toolbox, logo):
    """为仓库版绑定 code、角色别名和交付检查，保留原数字页码调用。"""
    stable_page = toolbox.for_page
    numbered_page = toolbox.creative

    def for_page(category, *args, **kwargs):
        return stable_page(ROLE_FALLBACK.get(category, category), *args, **kwargs)

    def creative(page, section="", title="", texts=None, images=None, roles=None,
                 source=None, vector_logos=False, **options):
        if not isinstance(page, str) or page.isdigit():
            return numbered_page(page, section, title, texts, images, roles, source, vector_logos, **options)
        if section or title or roles or vector_logos:
            raise ValueError("创意图示在已有功能区内放置，页眉和来源由当前页生成")
        if options.get("area") is not None and hasattr(toolbox,"_flag"):
            toolbox._flag("S.creative", options["area"])
        slide,report=place(toolbox.prs,toolbox.palette,page,texts=texts,images=images,
                           source=source or "",logo=logo,**options)
        if slide is None:
            return slide,report
        mode=report["mode"]
        number=next(i for i,s in enumerate(toolbox.prs.slides,1) if s.part is slide.part)
        group_name=report.pop("group_name","")
        siblings=[el for el in slide.shapes._spTree.iter(f'{{{P}}}grpSp')
                  if el.find(f'{{{P}}}nvGrpSpPr/{{{P}}}cNvPr').get("name")==group_name]
        if len(siblings)>1:
            harmonize_units(siblings)
        fill=report.get("text_fill")
        if fill is not None and fill<0.5:
            message=f"第 {number} 页 创意组件 {page}: 内容偏少，正文只占组件文字区的 {fill:.0%}；把每条说明写充分，或改用基础图示"
            if message not in toolbox.WARN:
                toolbox.WARN.append(message)
        toolbox.log.append([number, "creative:" + page, mode])
        toolbox.creative_records.append(dict(page=number, **report))
        return slide, report

    def creative_layout(code, section, title, regions, *, source, **options):
        item=entry(code)
        group=item["classification"]
        if "layout" not in item["modes"]:
            raise ValueError("该条目是图示或可复制单元")
        options.setdefault("partition",group.get("partition","auto") if len(regions)>1 else "auto")
        sl,areas=toolbox.content_page(section,title,regions,source=source,_layout_reference=True,**options)
        toolbox.bound[-1]["layout_reference"]=code
        return sl,areas

    toolbox.creative_layout=creative_layout
    def use_catalog_styles(*, header=None, marker=None):
        """Bind decoration once, before starting the deck. Structural bodies take explicit slots."""
        if len(toolbox.prs.slides):
            raise ValueError("请在封面之前选择整套样式")
        chosen = {}
        for family,code in (("header",header),("marker",marker)):
            if code is None:
                continue
            info = entry(code)
            if info["classification"]["family"] != family:
                raise ValueError(f"{code} 不属于 {family}")
            slots = inspect(code)
            max_text_slots = 1 if family=="marker" else 0
            if slots["required_images"] or slots["charts"] or slots["tables"] or len(slots["texts"]) > max_text_slots:
                raise ValueError(f"{code} 装饰含额外内容槽")
            chosen[family] = (info, slots)
        def draw(family, slide, area, label=None):
            info, slots = chosen[family]
            texts = {key:label for key in slots["texts"]}
            graphic,report=place(toolbox.prs,toolbox.palette,info["code"],slide=slide,area=area,
                                 texts=texts,semantic_role=info["role"],source="本套版式",
                                 reason="整套统一的"+family)
            if graphic is None:
                raise ValueError(report["reason"])
        if header:
            toolbox.header_decoration=lambda slide,area:draw("header",slide,area)
        if marker:
            def mark(slide, rect, title, index):
                x,y,w,h=rect
                if marker=="C1309":
                    draw("marker",slide,(x,y,min(w,380),41),title)
                    return
                draw("marker",slide,(x,y,46,40),str(index).zfill(2))
                toolbox.text(slide,x+54,y,w-54,42,title,20,toolbox.N,True,anchor="m", heading=True)
            toolbox.heading_decoration=mark
        toolbox.appearance.update({f:code for f,code in (("header",header),("marker",marker)) if code})

    def creative_structure(code, *, texts, source, reason, images=None):
        info=entry(code)
        family=info["classification"]["family"]
        if family not in ("agenda","divider"):
            raise ValueError("请选择目录或转场页样式")
        previous=toolbox.appearance.get("private_"+family)
        if previous and previous!=code:
            raise ValueError("同一份报告使用同一套结构页样式")
        slide=toolbox.prs.slides.add_slide(toolbox.prs.slide_layouts[6])
        try:
            graphic,report=place(toolbox.prs,toolbox.palette,code,slide=slide,area=(16,16,1248,688),
                                 texts=texts,images=images,source=source,reason=reason,semantic_role=info["role"])
            if graphic is None:
                raise ValueError(report["reason"])
        except Exception:
            sid=toolbox.prs.slides._sldIdLst[-1]
            toolbox.prs.part.drop_rel(sid.rId)
            toolbox.prs.slides._sldIdLst.remove(sid)
            raise
        slide._element.cSld.name="niuppt-"+family
        toolbox.appearance["private_"+family]=code
        toolbox.log.append([len(toolbox.prs.slides),"frame",family])
        return slide

    toolbox.use_catalog_styles=use_catalog_styles
    toolbox.creative_structure=creative_structure
    def icon(query, color=None):
        """按关键词从图标库取主色 PNG 路径，可直接给 box 外的图片或创意组件图片槽位使用。"""
        import icons
        return icons.render(query, color or toolbox.palette.get("navy", "1F4E79"))
    toolbox.icon=icon
    toolbox.for_page = for_page
    toolbox.creative = creative
    toolbox.creative_candidates = candidates
    toolbox.creative_inspect = inspect
    toolbox.creative_records = []
    def validate_report():
        from creative_cards import task_issues
        for n, slide in enumerate(toolbox.prs.slides, 1):
            errors = task_issues(slide, toolbox.palette)
            if errors:
                raise ValueError(f"P{n} 创意组件检查未通过：{errors}")
        check_creative_limit(toolbox.log)

    toolbox._check_creative_limit = validate_report

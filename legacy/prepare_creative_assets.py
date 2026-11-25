"""Prepare reviewed local assets from the three trial batches; never upload them.

python prepare_creative_assets.py --trial PATH --out PRIVATE_DIR --catalog JSON
The output includes all 41 pages, 29 cuts, verified hashes and actual slot IDs.
"""
from pathlib import Path
from copy import deepcopy
import argparse
import hashlib
import json
import re
from lxml import etree
from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from creative_assets import shapes, A, P, R


def clean(prs):
    # The source masters contain unrelated template artwork. Keep theme/style
    # definitions, but no master/layout shapes or unused slide relationships.
    for part in list(prs.part.package.iter_parts()):
        el = getattr(part, '_element', None)
        if el is None:
            continue
        name = str(part.partname)
        # Explicit East Asian fonts avoid renderer-dependent +mj-ea/+mn-ea
        # fallback after moving objects between masters. Preserve size/weight.
        for node in el.iter(f'{{{A}}}rPr', f'{{{A}}}defRPr', f'{{{A}}}endParaRPr'):
            ea=node.find(f'{{{A}}}ea')
            if ea is None: ea=etree.SubElement(node,f'{{{A}}}ea')
            ea.set('typeface','Source Han Sans CN')
        if '/slideMasters/' in name or '/slideLayouts/' in name:
            for tree in el.iter(f'{{{P}}}spTree'):
                for node in list(tree)[2:]:
                    tree.remove(node)
        for node in list(el.iter(f'{{{P}}}timing', f'{{{P}}}transition')):
            node.getparent().remove(node)
        used = {v for node in el.iter() for k, v in node.attrib.items() if k.startswith(f'{{{R}}}')}
        for rel in list(part.rels.values()):
            if rel.reltype in (RT.NOTES_SLIDE, RT.NOTES_MASTER) or (rel.reltype in (RT.IMAGE, RT.CHART, RT.HYPERLINK) and rel.rId not in used):
                part.drop_rel(rel.rId)
    for attr in ('author', 'last_modified_by', 'title', 'subject', 'keywords', 'comments'):
        setattr(prs.core_properties, attr, '')


def slots(slide):
    result = dict(text_ids=[], chart_slots={}, table_slots={}, image_ids=[], logo_ids=[], content_image_ids=[])
    for i, sh in shapes(slide).items():
        if getattr(sh, 'has_text_frame', False) and sh.text.strip() and set(sh.text.strip()) != {'·'}:
            result['text_ids'].append(i)
        if getattr(sh, 'has_chart', False):
            result['chart_slots'][str(i)] = dict(type=str(sh.chart.chart_type), series_count=len(sh.chart.series))
        if getattr(sh, 'has_table', False):
            result['table_slots'][str(i)] = [len(sh.table.rows), len(sh.table.columns)]
        if sh.shape_type == 13:
            result['image_ids'].append(i)
            if any(k in sh.name for k in ('scene_photo', 'showcase_image', 'certificate_document', 'case_image')):
                result['content_image_ids'].append(i)
            if sh.left > 1050 * 9525 and sh.top < 100 * 9525:
                result['logo_ids'].append(i)
    return result


def bounds(items, width):
    x=min(s.left for s in items); y=min(s.top for s in items)
    w=max(s.left+s.width for s in items)-x; h=max(s.top+s.height for s in items)-y
    return [int(x),int(y),int(w),int(h)], [round(n*1280/width,2) for n in (x,y,w,h)]


def prepare(trial, out, catalog_path):
    trial, out = Path(trial), Path(out)
    out.mkdir(parents=True, exist_ok=True)
    spec = json.loads((trial/'创意切分-2026-11-05/切分清单.json').read_text())
    cuts = {x['code']:deepcopy(x) for x in spec['items']}
    notes=(trial/'创意版式-41页调用说明-2026-11-05.md').read_text()
    sections={}
    role=''
    for line in notes.splitlines():
        if line.startswith('## '):
            role=line.split('、',1)[-1].split('（')[0].strip()
        m=re.match(r'\*\*([A-Z][A-Z0-9]*-?\d+) (.+?)\*\*', line)
        if m: sections[m[1]]=(m[2], role.split(' / ')[0])
    folders=['用途回补-2026-11-05/第一批/逐页成稿','用途回补-2026-11-05/第二批/逐页成稿',
             '常用图示补缺-2026-11-05/年报试做/用户调整版-2026-11-05/逐页成稿']
    full={'B3-115','B7-151','B7-159','B2-091'}
    manifest={'version':2,'items':{}}
    catalog={'version':2,'canvas':[1280,720],'items':[]}
    for batch,folder in enumerate(folders,1):
        for original in sorted((trial/folder).glob('*.pptx')):
            code=original.stem.split('-',1)[1]
            revised=list((trial/'问题页修订-2026-11-05').glob('*-'+code+'.pptx'))
            source=revised[0] if revised else original
            prs=Presentation(source); sl=prs.slides[0]
            item=dict(code=code, title=sections.get(code,(cuts.get(code,{}).get('title',code),''))[0],
                      role=sections.get(code,('', '客户伙伴'))[1], batch=batch,
                      revised=bool(revised), preferred_reuse=cuts[code]['kind'] if code in cuts else ('page' if code in full else 'reference'),
                      modes=['page'], **slots(sl))
            item['role']={'B1-040':'对比','B7-151':'分类定位','B3-115':'分类定位','B7-159':'对比'}.get(code,item['role'])
            if code=='L02':
                item.update(bars=[124,125,126,127],bar_categories=[137,139,141,143],bar_values=[138,140,142,144])
            # Name walls have exact entity sets, including all 32 K02 names.
            if code.startswith('K'):
                exclude={'K01':{102,104,161},'K02':{5,376},'K03':{2,135,137},'K05':{5,456,458},'K08':{1653,5,1659}}[code]
                item['entities']=[i for i in item['text_ids'] if i not in exclude]
            item['source_text_ids']=[i for i,s in shapes(sl).items() if getattr(s,'has_text_frame',False) and ('资料来源' in s.text or '来源：' in s.text)]
            clean(prs)
            page_file=code+'-page.pptx';prs.save(out/page_file)
            files={'page':dict(file=page_file,sha256=hashlib.sha256((out/page_file).read_bytes()).hexdigest())}
            if code in cuts:
                cut=cuts[code]
                if code=='K02':
                    cut['parts'][0]['ids'] += [351,358,363,364,365,368,369,370]
                if code=='B8-224':
                    cut['parts'][0]['ids']=[6,7]+cut['parts'][0]['ids']
                for part in cut['parts']:
                    part['ids']=list(dict.fromkeys(part['ids']))
                cutprs=Presentation(source);cs=cutprs.slides[0]
                allsh={s.shape_id:s for s in cs.shapes}
                original_order={s.shape_id:n for n,s in enumerate(cs.shapes)}
                keep={i for part in cut['parts'] for i in part['ids']}
                for sh in list(cs.shapes):
                    if sh.shape_id not in keep: cs.shapes._spTree.remove(sh._element)
                parts=[]; groups=[]
                for part in cut['parts']:
                    ids=set(part['ids']); chosen=[s for s in cs.shapes if s.shape_id in ids]
                    if len(chosen)!=len(ids): raise ValueError(f'{code}: missing selected IDs {ids-set(allsh)}')
                    # add_group_shape moves existing elements, retaining original
                    # order and relationship IDs. Each PART has a globally unique ID.
                    group=cs.shapes.add_group_shape(chosen)
                    groups.append((min(original_order[i] for i in ids),group))
                    group.name=f'PART|{code}|{part["name"]}'
                    emu,px=bounds([group],cutprs.slide_width)
                    one=deepcopy(part);one.pop('texts',None)
                    one.update(bbox_emu=emu,bbox_1280=px,group_id=group.shape_id)
                    parts.append(one)
                restore_group_order(cs,groups)
                clean(cutprs)
                cutfile=code+'-'+cut['kind']+'.pptx';cutprs.save(out/cutfile)
                files[cut['kind']]=dict(file=cutfile,sha256=hashlib.sha256((out/cutfile).read_bytes()).hexdigest())
                item['modes'].append(cut['kind'])
                item['parts']=parts
                item['cut_slots']=slots(cs)
                item['repeat']=cut.get('repeat')
                if item['repeat']:
                    if 'pitch_1280' in item['repeat']:
                        item['repeat']['reference_pitch_1280']=item['repeat'].pop('pitch_1280')
                    item['repeat']['note']='原页步距仅供参考；按实际单元宽高、目标区域与间隙计算，数量范围不代表已验证容量。'
                item['layout']=cut.get('layout')
                item['suggested_count']=cut.get('count')
            manifest['items'][code]=files
            catalog['items'].append(item)
    if len(catalog['items'])!=41: raise ValueError('expected exactly 41 source pages')
    (out/'manifest.json').write_text(json.dumps(manifest,ensure_ascii=False,indent=2)+'\n')
    Path(catalog_path).write_text(json.dumps(catalog,ensure_ascii=False,indent=2)+'\n')
    return catalog


def restore_group_order(slide, groups):
    """Separate PART groups must retain the source stacking order.

    Background/decor often appears last in metadata but first on the slide.
    Appending in metadata order would cover C04's lower text and K03's tiles.
    """
    for _,group in groups:slide.shapes._spTree.remove(group._element)
    for _,group in sorted(groups,key=lambda pair:pair[0]):
        slide.shapes._spTree.insert_element_before(group._element,'p:extLst')


if __name__=='__main__':
    ap=argparse.ArgumentParser();ap.add_argument('--trial',required=True);ap.add_argument('--out',required=True);ap.add_argument('--catalog',required=True)
    a=ap.parse_args();print(f'Prepared {len(prepare(a.trial,a.out,a.catalog)["items"])} pages')

"""按论据组织页面功能区。坐标单位与基础图示一致，为 1280 × 720 画布像素。"""
import math
import re
from collections import Counter

from frames import cols, rows

CUT_OVERLAP_LIMIT = 0.40
SKELETON_OVERLAP_LIMIT = 0.70
MIN_COMPARE_PAGES = 6
FIXED_ROLES = ("cover", "agenda", "divider", "closing")
CONTENT_FRAMES = ("standard", "open", "side", "split", "statement")
ALIASES = {"分项分析": "构成/分项", "分项": "构成/分项", "指标概览": "一组核心指标"}

# 类型来自已复核的基础图示清单。皮肤限制独立于区域几何。
REPEAT = frozenset("card card_grid module big_numbers kpi_tiles kpi_rings gauges numbered_rows bar_cards".split())
SINGLE = frozenset("""chevron_flow stairs steps node_row cycle funnel timeline gantt tree pyramid roof hub radial
compare sides_center matrix2x2 venn chart table row_matrix waterfall donut progress_bars panel chart_card
callout hero_left pictogram stat_strip banner""".split())
SOFT_ONLY = frozenset("card kpi_rings radial sides_center pyramid steps waterfall donut bar_cards chart_card node_row".split())
SOLID_ONLY = frozenset("module kpi_tiles panel row_matrix compare hero_left roof banner".split())
WIDE = frozenset("timeline gantt chevron_flow steps node_row stairs stat_strip banner roof".split())

def company_key(name):
    return re.sub(r"\s+", "", (name or "").replace("\ufeff", ""))

def zone_tokens(cut):
    return [s for s in cut.partition("|")[2].split("+") if s and not re.fullmatch(r"c\d+", s)]

def cut_of(frame, zones):
    return frame + ("|" + "+".join(dict.fromkeys(zones)) if zones else "")

def options(skin):
    if skin not in ("soft", "solid"):
        raise ValueError("skin must be soft or solid")
    return (SINGLE | REPEAT) - (SOLID_ONLY if skin == "soft" else SOFT_ONLY)

def plan(role, regions, skin, frame="open", partition="auto", *, layout_reference=False):
    if not regions or frame not in CONTENT_FRAMES:
        raise ValueError("内容页需要功能区和有效骨架")
    names = [r["diagram"] for r in regions]
    invalid = set(names) - options(skin)
    if invalid:
        raise ValueError(f"{skin} 不支持图示 {sorted(invalid)}")
    # 图表配明细表是汇报常用的数据页：图看趋势或结构，表列逐项口径，两者可同页。
    paired_data = sorted(names) == ["chart", "table"]
    if sum(name in SINGLE for name in names) > 1 and not paired_data:
        hit = [n for n in names if n in SINGLE]
        raise ValueError(f"一页最多一个单例图示，本页有 {hit}。单例图示一页只放一个（如时间轴、图表、表格、矩阵、金字塔），"
                         f"另一组材料改用可重复的 {sorted(REPEAT)} 之一；图表配明细表可以同页")
    if partition not in ("auto", "columns", "rows", "left", "top"):
        raise ValueError("未知的功能区划分方式")
    if partition in ("left", "top") and len(regions) < 2:
        raise ValueError("主辅分区至少需要两组材料")
    for region in regions:
        for key in ("weight", "aspect"):
            value = region.get(key, 1)
            if not math.isfinite(value) or value <= 0:
                raise ValueError(f"{key} 必须为正")
    return dict(role=role, frame=frame, zones=names, regions=list(regions), partition=partition,
                variant=partition, cut=cut_of(frame,names), density="按材料", note="")

def regions_in(area, regions, partition="auto"):
    if len(regions) == 1:
        return [area]
    weights = [r.get("weight", 1) for r in regions]
    targets = [r.get("aspect", 4 if r["diagram"] in WIDE else 1.25) for r in regions]
    def split(kind):
        if kind == "columns":
            return cols(area, weights)
        if kind == "rows":
            return rows(area, weights)
        main, rest = (cols if kind == "left" else rows)(area, [weights[0], sum(weights[1:])])
        return [main, *(rows if kind == "left" else cols)(rest, weights[1:])]
    kinds = ["columns", "rows"]
    if len(regions) > 2:
        kinds.extend(["left", "top"])
    layouts = [split(k) for k in (kinds if partition == "auto" else [partition])]
    valid = [v for v in layouts if all(w >= 140 and h >= 64 for _, _, w, h in v)]
    if not valid:
        raise ValueError("功能区过小，请减少分区或拆页")
    return min(valid, key=lambda v: sum(abs(math.log((w/h)/a)) for (_,_,w,h),a in zip(v,targets)))

def unit_boxes(area, count, aspect=2.1, gap=12):
    """按单元数与可用宽高选行列，尾行按实际条数分配宽度。"""
    if count < 1 or not math.isfinite(aspect) or aspect <= 0:
        raise ValueError("重复单元需要正数条目和有效宽高比")
    _, _, width, height = area
    if count > int((width+gap)//(95+gap)) * int((height+gap)//(52+gap)):
        raise ValueError("重复单元过多，请拆页")
    def boxes(ncol):
        bands = rows(area, [1] * math.ceil(count/ncol), gap)
        return [cell for j,band in enumerate(bands)
                for cell in cols(band, [1] * min(ncol,count-j*ncol), gap)]
    choices = [boxes(n) for n in range(1, count+1)]
    valid = [v for v in choices if all(w >= 95 and h >= 52 for _,_,w,h in v)]
    if not valid:
        raise ValueError("重复单元过多，请拆页")
    return min(valid, key=lambda v: sum(abs(math.log(w/h/aspect)) for _,_,w,h in v))

def conformance_warnings(company, skin, frames, bound, cuts_by_page, creative_records=()):
    pages = {}
    for page, cut in sorted(cuts_by_page.items()):
        for z in set(zone_tokens(cut)) & SINGLE:
            pages.setdefault(z, []).append(page)
    # 同一种单例图示占正文页三分之一以上时，整套会显得重复；低于此只作提示
    limit = max(2, len(cuts_by_page) // 3)
    return [f"[版式] 单例图示 {name} 出现 {len(ps)} 页（第 {'、'.join(map(str, ps))} 页）"
            + ("，超过建议上限 %d 页，换用其他图示" % limit if len(ps) > limit else "，复核是否确有需要")
            for name, ps in pages.items() if len(ps) > 1]

def looks_same(cuts_a, cuts_b, min_pages=MIN_COMPARE_PAGES):
    """两套成稿是否雷同。返回 (是否雷同, 明细)。"""
    # 致谢是固定收尾；不要让它改变已有成稿之间的雷同度。
    cuts_a = [cut for cut in cuts_a if cut.split("|", 1)[0] != "closing"]
    cuts_b = [cut for cut in cuts_b if cut.split("|", 1)[0] != "closing"]
    n = min(len(cuts_a), len(cuts_b))
    detail = {"n": n, "cut_same": 0.0, "frame_same": 0.0, "zone_same": 0.0, "reason": ""}
    if n < min_pages:
        return False, detail
    detail["cut_same"] = sum(1 for i in range(n) if cuts_a[i] == cuts_b[i]) / n
    fa = [c.split("|")[0] for c in cuts_a[:n]]
    fb = [c.split("|")[0] for c in cuts_b[:n]]
    detail["frame_same"] = sum(1 for i in range(n) if fa[i] == fb[i]) / n
    detail["zone_same"] = sum(1 for i in range(n) if zone_tokens(cuts_a[i]) == zone_tokens(cuts_b[i])) / n
    if detail["cut_same"] > CUT_OVERLAP_LIMIT:
        detail["reason"] = "cut"
    elif detail["frame_same"] > SKELETON_OVERLAP_LIMIT and detail["zone_same"] > SKELETON_OVERLAP_LIMIT:
        detail["reason"] = "skeleton"
    return bool(detail["reason"]), detail

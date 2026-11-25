# -*- coding: utf-8 -*-
"""旧版素材放置的独立入口：由 scripts/assets.py 在单独进程里调用，避免与 v2 的同名模块冲突。

stdin 传 JSON：
  {"op": "inspect", "code": "B3-052"}
  {"op": "apply", "pptx": "...", "palette": {...}, "jobs": [{"slide": 5, "code": "...", "texts": {...}, "icons": [...],
                                                              "images": {...}, "area": [x, y, w, h], "items": [...]}]}
stdout 输出 JSON 结果。本目录的模块来自 v1（main 分支）的创意素材放置代码，只做路径调整。
"""
import contextlib
import io
import json
import math
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

from pptx import Presentation  # noqa: E402
from lxml import etree  # noqa: E402
import creative_catalog as C  # noqa: E402


def inspect(code):
    it = C.entry(code)
    mode = it["preferred_reuse"]
    out = {"code": code, "title": it.get("title"), "role": it["role"], "mode": mode}
    if mode == "layout":
        out["placeable"] = False
        return out
    info = C.inspect(code, mode=mode)
    out.update(placeable=True, slots=info.get("slots", {}), vector_icons=info.get("vector_icons", []),
               images=info.get("images", []), decor_texts=info.get("decor_texts", {}), repeatable=info.get("repeatable"))
    return out


def _grid(area, n, aspect):
    """把 n 个单元排进 area，单元宽高比接近 aspect（宽/高）。返回 [x, y, w, h]。"""
    x, y, w, h = area
    gap = 16
    best = None
    for cols in range(1, n + 1):
        rows = math.ceil(n / cols)
        cw = (w - gap * (cols - 1)) / cols
        ch = (h - gap * (rows - 1)) / rows
        uw = min(cw, ch * aspect)
        uh = uw / aspect
        if best is None or uw * uh > best[0]:
            best = (uw * uh, cols, rows, uw, uh)
    _, cols, rows, uw, uh = best
    tw = cols * uw + gap * (cols - 1)
    th = rows * uh + gap * (rows - 1)
    x0, y0 = x + (w - tw) / 2, y + (h - th) / 2
    return [[x0 + (i % cols) * (uw + gap), y0 + (i // cols) * (uh + gap), uw, uh] for i in range(n)]


KEEP_ICONS = {"B5-072", "B7-192", "B8-465"}  # 月桂枝、地图点阵/省块是构图本身
KEEP_IMAGES = {"B5-142", "B8-465"}            # 点阵世界地图、地图定位针是构图本身        # 这些素材的“矢量图标”是构图本身（月桂枝），保留原样
AREA = [32, 112, 1216, 566]   # 正文区左右各放宽 16px，下沿到页脚上方
MIN_PT = (12, 10.5, 9.5, 9)   # 先按 12pt 放；放不下再逐级降低下限试放，并提醒（B7-159 原稿正文约 9pt）


def _call(prs, palette, code, mode, sl, box, pl, it, job, icons, min_pt):
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        r = C.place(prs, palette, code, mode=mode, slide=sl, area=tuple(box),
                    texts={str(k): v for k, v in (pl.get("texts") or {}).items()},
                    icons=icons, images=pl.get("images"),
                    source=job.get("source") or "材料", reason=job.get("reason") or "按材料关系选用",
                    semantic_role=it["role"], min_text_pt=min_pt)
    return r, [l for l in buf.getvalue().splitlines() if l.strip()]


def _place_one(prs, palette, code, mode, sl, box, pl, it, job):
    import icons as LI
    icons = None if code in KEEP_ICONS else pl.get("icons")
    notes = []
    if icons:
        fixed = []
        for q in icons:
            try:
                LI.pick(q); fixed.append(q)
            except Exception:
                fixed.append("重点"); notes.append("图标关键词「%s」不在图标库里，已用通用图标；可用 scripts/assets.py 图标 查看可用关键词" % q)
        icons = fixed
    if code in KEEP_IMAGES:
        pl = dict(pl, images=None)
    if pl.get("images"):
        imgs = {}
        for k, v in pl["images"].items():
            if isinstance(v, dict) and v.get("icon"):
                try:
                    LI.pick(v["icon"])
                except Exception:
                    notes.append("图标关键词「%s」不在图标库里，已用通用图标" % v["icon"]); v = dict(v, icon="重点")
            imgs[k] = v
        pl = dict(pl, images=imgs)
    last = ""
    for min_pt in MIN_PT:
        try:
            r, n = _call(prs, palette, code, mode, sl, box, pl, it, job, icons, min_pt)
        except Exception as ex:
            last = str(ex)[:300]
            if "低于" in last:
                continue
            return False, last, notes
        if r and r[0] is not None:
            if min_pt < 12:
                notes.append("按 12pt 放不下，最小字号降到 %.1fpt；文字偏多，考虑精简或换素材" % min_pt)
            if code in KEEP_IMAGES or code in KEEP_ICONS:
                n = [l for l in n if "未替换" not in l and "图标（从上到下" not in l and "模板矢量图标" not in l]
            return True, "", notes + n
        last = str(r[1] if r else "未放置")[:300]
    return False, last, notes


def unify_font(slide, font):
    """素材里的文字统一换成整套字体（模板原稿常混用细体、衬线体）。"""
    A = "http://schemas.openxmlformats.org/drawingml/2006/main"
    for grp in slide.shapes:
        if not grp.name.startswith("creative:"):
            continue
        for tag in ("rPr", "defRPr", "endParaRPr"):
            for r in grp._element.iter("{%s}%s" % (A, tag)):
                for sub in ("latin", "ea", "cs"):
                    old = r.find("{%s}%s" % (A, sub))
                    if old is not None:
                        r.remove(old)
                # schema 顺序：latin/ea/cs 必须在 sym、hlinkClick、hlinkMouseOver、rtl、extLst 之前
                after = [c for c in r if etree.QName(c).localname in ("sym", "hlinkClick", "hlinkMouseOver", "rtl", "extLst")]
                pos = list(r).index(after[0]) if after else len(r)
                for i, sub in enumerate(("latin", "ea", "cs")):
                    el = r.makeelement("{%s}%s" % (A, sub), {"typeface": font})
                    r.insert(pos + i, el)


def apply(pptx, palette, jobs):
    prs = Presentation(pptx)
    results = []
    for job in jobs:
        code = job["code"]
        sl = prs.slides[job["slide"] - 1]
        it = C.entry(code)
        mode = it["preferred_reuse"]
        area = job.get("area") or AREA
        units = job.get("items")
        boxes = [area]
        payloads = [job]
        if units:
            import layout_combo
            part = (it.get("parts") or [{}])[0].get("bbox_1280")
            aspect = (part[2] / part[3]) if part else 2.0
            boxes = layout_combo.unit_boxes(tuple(area), len(units), aspect, gap=14)
            payloads = units
        rec = {"slide": job["slide"], "code": code, "placed": 0, "errors": [], "notes": []}
        for box, pl in zip(boxes, payloads):
            ok, msg, notes = _place_one(prs, palette, code, mode, sl, box, pl, it, job)
            rec["notes"] += notes
            if ok:
                rec["placed"] += 1
            else:
                rec["errors"].append(msg)
        results.append(rec)
        if palette.get("font"):
            unify_font(sl, palette["font"])
    prs.save(pptx)
    return results


def main():
    req = json.load(sys.stdin)
    if req["op"] == "inspect":
        out = inspect(req["code"])
    elif req["op"] == "icons":
        import icons as LI
        out = {"keywords": sorted({w for n in LI.names() for w in LI.KEYWORDS.get(n, "").split()}), "names": LI.names()}
    elif req["op"] == "apply":
        out = apply(req["pptx"], req["palette"], req["jobs"])
    else:
        out = {"error": "未知操作"}
    sys.stdout.write("\n@@RESULT@@" + json.dumps(out, ensure_ascii=False))


if __name__ == "__main__":
    main()

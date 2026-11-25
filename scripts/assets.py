# -*- coding: utf-8 -*-
"""创意素材页：把素材库里的原版创意构图（assets/library/creative）放进 v2 页面。

放置、填字、换色、换图标由 legacy/ 下的 v1 代码完成，在单独进程里运行。

    python3 scripts/assets.py 列表              列出可用的创意素材及用途
    python3 scripts/assets.py 查看 B3-052       列出文字槽位（编号、示例文字、位置）和图标槽位
    python3 scripts/assets.py 图标              列出素材页可用的图标关键词

deck.json 里这样写一页：
    {"section": "...", "title": "...", "separation": "创意型",
     "creative": {"kind": "asset", "code": "B3-052",
                  "texts": {"132": "一体化平台", "20": "线上开户", "21": "说明……", ...},
                  "icons": ["开户", "审批", ...],            # 矢量小图标，按 查看 列出的 order 顺序
                  "images": {"179": {"icon": "钱包"}},       # 图片型图标槽位
                  "items": [{"texts": {...}}, ...]}}         # 可复制单元（K05、B5-072）：每个单元一份
"""
import json
import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
LEGACY = os.path.join(os.path.dirname(HERE), "legacy")
RUNNER = os.path.join(LEGACY, "runner.py")

# 可用素材与用途（用户 2026-11-17 从旧版素材里选定；地图三种待标准地图底图）
USES = {
    "B3-115": "一个总量拆成左右两组，各组 3 项",
    "B7-151": "SWOT 分析",
    "K02": "客户、合作伙伴名单（约 32 个名称）",
    "K05": "客户、合作伙伴名称格（按名单数量重复排开）",
    "K08": "以一个核心对象为中心的上下游、合作方（约 28 个名称）",
    "B7-159": "双向反馈、循环闭环",
    "B8-404": "三级结构：组织架构、总分拆解",
    "B3-052": "一个中心 + 6 个要点",
    "B3-065": "一个平台 + 8 个场景或产品",
    "B3-109": "一个核心 + 两侧 6 个分支（产品、客户分类）",
    "B5-072": "荣誉、获奖（按数量重复排开）",
    "B5-142": "世界地图（点阵）+ 3 个地点标注；宜放在切分树的一格里，旁边格放大数字或说明",
    "B7-192": "世界地图 + 4 张区域卡；宜放在一格里，旁边格放指标栏、底部放说明",
    "B8-465": "中国地图 + 4 个地区标注；宜放在左格，右格放指标面板",
}
NOT_READY = {"B8-304": "证书陈列版式，原素材只能借结构，暂不可直接放置"}
_CACHE = {}


def _run(req, timeout=180):
    p = subprocess.run([sys.executable, RUNNER], input=json.dumps(req, ensure_ascii=False), capture_output=True,
                       text=True, timeout=timeout, cwd=LEGACY)
    out = p.stdout
    if "@@RESULT@@" not in out:
        raise RuntimeError("素材放置进程出错：" + (p.stderr or out)[-800:])
    return json.loads(out.split("@@RESULT@@", 1)[1])


def inspect(code):
    if code not in _CACHE:
        _CACHE[code] = _run({"op": "inspect", "code": code})
    return _CACHE[code]


def palette_for(K):
    """v2 配色 → v1 放置代码使用的配色角色。"""
    return {"navy": K.MAIN, "dark": K.mix(K.MAIN, "000000", .3), "navy2": K.mix(K.MAIN, "FFFFFF", .35),
            "tint": K.mix(K.MAIN, "FFFFFF", .9), "peach": K.mix(K.MAIN, "FFFFFF", .45), "ink": K.INK, "muted": K.MUTED, "font": K.FONT}


def apply(pptx, palette, jobs):
    if not jobs:
        return []
    return _run({"op": "apply", "pptx": os.path.abspath(pptx), "palette": palette, "jobs": jobs}, timeout=600)


def check(cr):
    """规格检查：编号可用、必填文字齐全。返回错误列表。"""
    code = cr.get("code")
    if code in NOT_READY:
        return ["%s 暂不可用：%s" % (code, NOT_READY[code])]
    if code not in USES:
        return ["未知素材编号 %s，可用：%s" % (code, "、".join(USES))]
    try:
        info = inspect(code)
    except Exception as ex:  # 依赖缺失时不在规格检查阶段拦
        return []
    need = set(info.get("slots", {})) - set(info.get("decor_texts", {}))
    err = []
    groups = cr.get("items") if info.get("repeatable") else [cr]
    if info.get("repeatable") and not cr.get("items"):
        return ["%s 是可复制单元，要用 items 按条目各写一份 texts" % code]
    for i, g in enumerate(groups or []):
        miss = sorted(need - set(map(str, (g.get("texts") or {}).keys())), key=lambda s: int(s))
        if miss:
            err.append("%s%s 缺文字槽位 %s（用 scripts/assets.py 查看 %s）" % (code, " 第 %d 个单元" % (i + 1) if info.get("repeatable") else "",
                                                                     "、".join(miss[:12]), code))
    return err


def main(argv=None):
    a = list(argv if argv is not None else sys.argv[1:])
    if not a or a[0] in ("列表", "list"):
        for c, u in USES.items():
            print("%-7s %s" % (c, u))
        for c, u in NOT_READY.items():
            print("%-7s 〔暂不可用〕%s" % (c, u))
        return 0
    if a[0] in ("图标", "icons"):
        r = _run({"op": "icons"})
        print("素材页图标（Lucide）可用中文关键词：" + "、".join(r["keywords"]))
        print("也可直接写英文名：" + ", ".join(r["names"]))
        return 0
    code = a[-1]
    info = inspect(code)
    print("%s %s（%s，%s）" % (code, info.get("title"), info.get("role"), "可复制单元" if info.get("repeatable") else "整块"))
    for k, s in sorted(info.get("slots", {}).items(), key=lambda kv: (kv[1]["y"], kv[1]["x"])):
        print("  文字 %-5s 位置(%4d,%4d) 宽%4d  示例：%s" % (k, s["x"], s["y"], s["w"], s["text"]))
    for ic in info.get("vector_icons", []):
        print("  图标 #%d 位置(%d,%d) 靠近文字 %s" % (ic["order"], ic["x"], ic["y"], ic.get("near_slot")))
    for im in info.get("images", []):
        print("  图片槽位 %s（传 images={\"%s\": {\"icon\": \"关键词\"}}）" % (im, im))
    return 0


if __name__ == "__main__":
    sys.exit(main())

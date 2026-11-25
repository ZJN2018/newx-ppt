# -*- coding: utf-8 -*-
"""渲染后留白检查：按逐页 PNG 量每个格子里实际内容之外的最大空块。

规格检查只看格子尺寸，不知道卡片里的字和图实际占了多少；这里在出图后补上这一步。
每格先找出"有内容"的像素（与该格底色差异明显的点：文字、图标、图表、线条），
按 10px 网格稍作膨胀后，反复取格内最大的全空矩形（不小于 160×100）；空块合计超过格子 35% 就提醒：
先收格、换紧凑组件或并页；仍空的，考虑在挨着正文区一角的位置加一个图片格。

python3 check_blank.py <名>-layout.json [--png-dir 目录]
build.py 渲染后会自动调用，并把结果写入 <名>-留白检查.json。
"""
import argparse
import json
import os
import sys

from PIL import Image

SLIDE_W = 1280
GRID = 10          # 网格边长（幻灯片像素）
INSET = 8          # 去掉卡片描边与阴影
INK = 34           # 与底色的 RGB 距离超过此值算内容
SPREAD = 1         # 内容格向四周膨胀的格数，把同一段文字的行间距连起来
WARN_RATIO = 0.35  # 格内空块合计占格子面积的比例
HOLE_MIN = (160, 100)  # 计入的空块最小宽高（幻灯片像素），行距、窄缝不算
PAGE_RATIO = 0.15  # 整页空块合计占正文区比例，超过就建议考虑加图
HARD_RATIO = 0.20  # 超过这个比例又没有图片格：必须处理（收格到 20% 以下，或加一个挨角的图片格），生成返回失败
SKIP = {"photo", "asset", "bar_h", "donut", "waterfall"}  # 图片和图表自身会铺满格子；表格按内容高度画，下方空出的部分照常计入
BODY_AREA = 1184 * 558


def _mode_color(px, w, h):
    """格子的底色：边缘一圈像素里出现最多的颜色（量化到 8 级，抗渐变）。"""
    counts = {}
    step = max(1, min(w, h) // 60)
    pts = [(x, y) for x in range(0, w, step) for y in (0, h - 1)] + [(x, y) for y in range(0, h, step) for x in (0, w - 1)]
    for x, y in pts:
        r, g, b = px[x, y][:3]
        k = (r // 8, g // 8, b // 8)
        counts[k] = counts.get(k, 0) + 1
    k = max(counts, key=counts.get)
    return (k[0] * 8 + 4, k[1] * 8 + 4, k[2] * 8 + 4)


def _occupancy(img, box, scale):
    x, y, w, h = box
    x0, y0 = int((x + INSET) * scale), int((y + INSET) * scale)
    x1, y1 = int((x + w - INSET) * scale), int((y + h - INSET) * scale)
    crop = img.crop((x0, y0, x1, y1)).convert("RGB")
    cw, ch = crop.size
    if cw < 4 or ch < 4:
        return None
    px = crop.load()
    bg = _mode_color(px, cw, ch)
    g = GRID * scale
    cols, rows = max(1, int(cw / g)), max(1, int(ch / g))
    occ = [[False] * cols for _ in range(rows)]
    for r in range(rows):
        for c in range(cols):
            hit = False
            for yy in range(int(r * g), min(ch, int((r + 1) * g))):
                for xx in range(int(c * g), min(cw, int((c + 1) * g))):
                    p = px[xx, yy]
                    if (p[0] - bg[0]) ** 2 + (p[1] - bg[1]) ** 2 + (p[2] - bg[2]) ** 2 > INK * INK:
                        hit = True
                        break
                if hit:
                    break
            occ[r][c] = hit
    if SPREAD:
        grown = [row[:] for row in occ]
        for r in range(rows):
            for c in range(cols):
                if occ[r][c]:
                    for dr in range(-SPREAD, SPREAD + 1):
                        for dc in range(-SPREAD, SPREAD + 1):
                            if 0 <= r + dr < rows and 0 <= c + dc < cols:
                                grown[r + dr][c + dc] = True
        occ = grown
    return occ


def _largest_empty(occ):
    """最大全空矩形（直方图法），返回 (面积, 行, 列, 高, 宽)，单位为网格。"""
    if not occ:
        return (0, 0, 0, 0, 0)
    cols = len(occ[0])
    hist = [0] * cols
    best = (0, 0, 0, 0, 0)
    for r, row in enumerate(occ):
        for c in range(cols):
            hist[c] = 0 if row[c] else hist[c] + 1
        stack = []
        for c in range(cols + 1):
            hh = hist[c] if c < cols else 0
            start = c
            while stack and stack[-1][1] >= hh:
                s, sh = stack.pop()
                area = sh * (c - s)
                if area > best[0]:
                    best = (area, r - sh + 1, s, sh, c - s)
                start = s
            stack.append((start, hh))
    return best


def _holes(occ):
    """反复取最大全空矩形，只保留够大的，返回 [(行, 列, 高, 宽)]（网格单位）。"""
    occ = [r[:] for r in occ]
    mw, mh = HOLE_MIN[0] // GRID, HOLE_MIN[1] // GRID
    out = []
    for _ in range(12):
        area, r, c, h, w = _largest_empty(occ)
        if area < mw * mh:
            break
        if w >= mw and h >= mh:
            out.append((r, c, h, w))
        for i in range(r, r + h):
            for j in range(c, c + w):
                occ[i][j] = True
    return out


def check_page(png, cells, slide_w=SLIDE_W):
    """cells: [{"box": [x,y,w,h], "name": str, "type": str}]。返回该页空块列表。"""
    img = Image.open(png)
    scale = img.size[0] / slide_w
    found = []
    for cell in cells:
        if cell.get("type") in SKIP:
            continue
        occ = _occupancy(img, cell["box"], scale)
        if occ is None:
            continue
        n = sum(map(len, occ))
        holes = _holes(occ)
        ratio = sum(h * w for _, _, h, w in holes) / n if n else 0
        if ratio >= WARN_RATIO:
            x, y = cell["box"][:2]
            r, c, hh, ww = holes[0]
            found.append({"格子": cell.get("name", ""), "组件": cell.get("type", ""),
                          "最大空块": [round(x + INSET + c * GRID), round(y + INSET + r * GRID), ww * GRID, hh * GRID],
                          "空块面积": round(ratio * (cell["box"][2] - 2 * INSET) * (cell["box"][3] - 2 * INSET)),
                          "占格子": round(ratio, 2)})
    return found


def check(layout, png_dir):
    """layout: build.py 写出的 {"pages": [{"slide": n, "title": ..., "png": ..., "cells": [...], "has_photo": bool}]}"""
    report = []
    for pg in layout["pages"]:
        png = os.path.join(png_dir, pg["png"])
        if not os.path.exists(png):
            continue
        holes = check_page(png, pg["cells"])
        if not holes:
            continue
        total = sum(h["空块面积"] for h in holes) / BODY_AREA
        item = {"页": pg["slide"], "标题": pg["title"], "空块": holes, "空块占正文区": round(total, 2)}
        if any(c.get("type") == "空图片格" for c in pg["cells"]):
            item["必须"] = True
            item["建议"] = "图片格没取到图，留下一块空：换一个中英文关键词或写本机图片路径重跑；确实没有相关的图，就去掉这一格并收格"
        elif total >= HARD_RATIO and not pg.get("has_photo"):
            item["必须"] = True
            item["建议"] = ("必须处理：收格、换紧凑组件或并页，把空白降到正文区 %d%% 以下；内容确实只有这些时，"
                          "加一个挨着正文区一角、与本页相关的图片格（见 components.md 图片）" % (HARD_RATIO * 100))
        elif total >= PAGE_RATIO and not pg.get("has_photo"):
            item["建议"] = "先收格、换紧凑组件或与相邻页合并；内容确实只有这些时，加一个挨着正文区一角的图片格（见 components.md 图片）"
        else:
            item["建议"] = "收紧该格（缩比例、换紧凑组件），或把相邻格内容并进来"
        report.append(item)
    return report


def print_report(report):
    if not report:
        print("[留白] 未发现大块空白")
        return
    for it in report:
        holes = "；".join("「%s」空白约 %d%%（最大一块 %d×%d）" % (h["格子"], h["占格子"] * 100, h["最大空块"][2], h["最大空块"][3])
                         for h in it["空块"])
        print("%s：[留白] 第 %d 页「%s」整页空白约 %d%%，%s → %s" % ("错误" if it.get("必须") else "提醒", it["页"], it["标题"][:18],
              it["空块占正文区"] * 100, holes, it["建议"]))
    hard = [it["页"] for it in report if it.get("必须")]
    if hard:
        print("留白检查未通过：第 %s 页有空图片格，或空白超过正文区 %d%% 且没有图片格。改 deck.json 后重跑。"
              % ("、".join(map(str, hard)), HARD_RATIO * 100))


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("layout")
    ap.add_argument("--png-dir")
    a = ap.parse_args(argv)
    layout = json.load(open(a.layout, encoding="utf-8"))
    rep = check(layout, a.png_dir or os.path.dirname(os.path.abspath(a.layout)))
    print_report(rep)
    return 3 if any(it.get("必须") for it in rep) else 0


if __name__ == "__main__":
    sys.exit(main())

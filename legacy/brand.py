"""从公司 LOGO 推导整套配色：主色取自 LOGO，其余按通行配色规则生成。

规则
1. 主色 = LOGO 中占比最高的彩色（非黑白灰）。
2. 深色 = 主色同色相压暗（HSL 亮度约 20%），用于标题帽、结论带；不另引入新色相。
3. 浅底 = 主色提亮到 93% 左右，用于模块底、卡片底。
4. 强调色（只用一个）：
   - 多色 LOGO 用第二彩色；单色 LOGO 沿用主色或选定的同色相深色。
5. 稳重风大面积用主色与深色；轻盈风用白底、品牌浅底和局部点色。
6. 对比度（WCAG 2.1）：白字压在主色、深色、强调色上 ≥ 4.5:1（大号粗体 ≥ 3:1 视为可用并给出提示）；正文色压在浅底上 ≥ 7:1。不达标时自动把该色压暗直到达标。
7. 灰阶带一点主色色相，不用纯灰。
8. 涨跌色：主色或强调色为红、绿色系时，不用“红涨绿跌”，改用 ▲▼ 箭头区分，避免和品牌色冲突。
"""
import colorsys, warnings
warnings.filterwarnings("ignore", category=DeprecationWarning)
from collections import Counter
from PIL import Image

def _rgb(h): return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))
def _hex(rgb): return "%02X%02X%02X" % tuple(max(0, min(255, int(round(v)))) for v in rgb)
def mix(h, t, k):
    a, b = _rgb(h), _rgb(t); return _hex([x + (y - x) * k for x, y in zip(a, b)])
def hls(h):
    r, g, b = (v / 255 for v in _rgb(h)); return colorsys.rgb_to_hls(r, g, b)
def from_hls(hh, l, s):
    return _hex([v * 255 for v in colorsys.hls_to_rgb(hh % 1, max(0, min(1, l)), max(0, min(1, s)))])

def luminance(h):
    def ch(c):
        c /= 255; return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
    r, g, b = (ch(v) for v in _rgb(h)); return 0.2126 * r + 0.7152 * g + 0.0722 * b
def contrast(a, b):
    la, lb = sorted((luminance(a), luminance(b)), reverse=True); return (la + 0.05) / (lb + 0.05)

def darken_until(h, against="FFFFFF", target=4.5):
    hh, l, s = hls(h)
    while contrast(h, against) < target and l > 0.05:
        l -= 0.01; h = from_hls(hh, l, s)
    return h

def logo_colors(path, n=6):
    """按色相聚类取色：JPEG 压缩和抗锯齿会把同一种颜色打散成许多相近像素，逐像素计数会低估文字类色块，
    所以先按 20° 色相分桶统计面积，再取每桶中饱和度最高的一批像素的中位色作为代表色。"""
    im = Image.open(path).convert("RGBA"); im.thumbnail((400, 400))
    buckets, darks, neutral = {}, Counter(), Counter()
    for r, g, b, a in im.getdata():
        if a < 200 or (r > 235 and g > 235 and b > 235): continue
        hh, l, s = colorsys.rgb_to_hls(r / 255, g / 255, b / 255)
        if s > 0.45 and 0.2 < l < 0.75:
            buckets.setdefault(int(hh * 360) // 20, []).append((s, (r, g, b)))
        elif l < 0.25: darks[(r // 8 * 8, g // 8 * 8, b // 8 * 8)] += 1
        else: neutral[(r // 8 * 8, g // 8 * 8, b // 8 * 8)] += 1
    # 相邻桶合并（如 180° 与 200° 同属蓝色）
    keys = sorted(buckets); groups = []
    for k in keys:
        if groups and k - groups[-1][-1] <= 1: groups[-1].append(k)
        else: groups.append([k])
    if len(groups) > 1 and groups[0][0] == 0 and groups[-1][-1] == 17: groups[0] = groups.pop() + groups[0]
    out = []
    for g in groups:
        px = [p for k in g for p in buckets[k]]
        px.sort(key=lambda t: -t[0]); top = [c for _, c in px[: max(10, len(px) // 3)]]
        med = tuple(sorted(ch)[len(ch) // 2] for ch in zip(*top))
        out.append((_hex(med), len(px), "chroma"))
    out.sort(key=lambda t: -t[1])
    if darks: out.append((_hex(darks.most_common(1)[0][0]), sum(darks.values()), "dark"))
    if neutral: out.append((_hex(neutral.most_common(1)[0][0]), sum(neutral.values()), "neutral"))
    return out[:n]

def logo_color_details(path):
    """取色的位置与彩色面积占比；仅在内嵌、同高多笔画特征明确时判断文字。

    占比的分母是全部彩色像素，排除透明、白底和中性色。文字判断是保守的
    几何提示，不是 OCR；无法判断时只写图形/色块，仍由配色确认供人核对。
    """
    im = Image.open(path).convert("RGBA")
    im.thumbnail((400, 400))
    colors = [c for c, _, kind in logo_colors(path) if kind == "chroma"]
    pixels = {color: set() for color in colors}
    for y in range(im.height):
        for x in range(im.width):
            r, g, b, a = im.getpixel((x, y))
            hh, light, saturation = colorsys.rgb_to_hls(r / 255, g / 255, b / 255)
            if a < 200 or saturation <= .45 or not .2 < light < .75:
                continue
            if colors:
                color = min(colors, key=lambda c: hue_delta(c, _hex((r, g, b))))
                pixels[color].add((x, y))
    total = sum(map(len, pixels.values()))

    def bounds(points):
        return (min(x for x, y in points), min(y for x, y in points),
                max(x for x, y in points) + 1, max(y for x, y in points) + 1)

    def components(points):
        remaining = set(points)
        found = []
        while remaining:
            seed = remaining.pop()
            stack, group = [seed], {seed}
            while stack:
                x, y = stack.pop()
                for neighbor in ((x-1, y), (x+1, y), (x, y-1), (x, y+1)):
                    if neighbor in remaining:
                        remaining.remove(neighbor)
                        group.add(neighbor)
                        stack.append(neighbor)
            if len(group) >= max(15, len(points) * .08):
                found.append((len(group), bounds(group)))
        return found

    result = []
    for color, points in pixels.items():
        if not points:
            continue
        box = bounds(points)
        others = set().union(*(p for c, p in pixels.items() if c != color))
        parts = components(points)
        text = False
        if others and len(parts) >= 2:
            outer = bounds(others)
            ow, oh = outer[2]-outer[0], outer[3]-outer[1]
            nested = (box[0] > outer[0]+ow*.06 and box[2] < outer[2]-ow*.06
                      and box[1] > outer[1]+oh*.06 and box[3] < outer[3]-oh*.06)
            heights = [b[3]-b[1] for _, b in parts]
            bottoms = [b[3] for _, b in parts]
            aligned = min(heights) >= max(heights)*.65 and max(bottoms)-min(bottoms) <= max(heights)*.2
            text = nested and aligned and sum(n for n, _ in parts) >= len(points)*.65
        cx = sum(x for x, y in points) / len(points) / im.width
        cy = sum(y for x, y in points) / len(points) / im.height
        horizontal = "左部" if cx < .4 else "右部" if cx > .6 else "中部"
        vertical = "上部" if cy < .4 else "下部" if cy > .6 else "中部"
        position = "中央" if horizontal == vertical == "中部" else vertical + "、" + horizontal
        result.append(dict(color=color, area_share=round(len(points)/total, 4),
                           area_basis="LOGO 彩色像素", position=position,
                           bbox=[round(v / (im.width if i % 2 == 0 else im.height), 4) for i, v in enumerate(box)],
                           primary_origin="来自文字" if text else "图形/色块",
                           origin_method="内嵌同高笔画几何判断" if text else "未识别为文字"))
    return result


def family(h):
    deg = hls(h)[0] * 360
    if deg >= 330 or deg < 20: return "red"
    if deg < 70: return "orange"
    if deg < 170: return "green"
    return "blue"   # 蓝 / 青 / 紫

def auto_accent(main):
    f = family(main)
    if f == "red": return from_hls(40 / 360, 0.46, 0.72)      # 金
    if f == "blue": return from_hls(22 / 360, 0.66, 0.70)     # 杏橙
    if f == "green": return from_hls(28 / 360, 0.55, 0.80)    # 橙
    return from_hls(222 / 360, 0.28, 0.60)                    # 藏蓝

def _norm_hex(h):
    s = str(h).strip().lstrip("#").upper()
    if len(s) != 6 or any(c not in "0123456789ABCDEF" for c in s):
        raise ValueError("颜色须为 6 位十六进制，例如 0088D0")
    return s

def hue_delta(a, b):
    d = abs(hls(a)[0] - hls(b)[0]) * 360 % 360
    return min(d, 360 - d)

def palette_from_logo(path, accent=None, main=None):
    cols = logo_colors(path)
    chroma = [c for c, _, k in cols if k == "chroma"]
    base = main or chroma[0]
    main = darken_until(base, target=4.5)
    hh, l, s = hls(main)
    dark = from_hls(hh, 0.20, min(1, s * 0.85))
    acc, _, _ = _accent_from_logo(chroma, base, accent)
    # 单色 LOGO 也要有强调色：按主流配色从主色派生（红配金、蓝配杏橙、绿配橙），不与主色相同
    acc = acc or auto_accent(base)
    acc_text = darken_until(acc, target=3.0)            # 强调色上压白色大号数字
    tint = from_hls(hh, 0.955, min(0.55, s))
    pal = dict(
        navy=main, navy2=from_hls(hh, min(0.62, l + 0.10), s), dark=dark,
        peach=acc_text, peach_l=from_hls(hls(acc)[0], 0.94, 0.6),
        ink="1A1A1A", muted=mix("6A6A6A", main, 0.06), line=mix("DCDCDC", main, 0.05),
        tint=tint, white="FFFFFF",
    )
    pal["chart"] = [main, acc_text, mix("8C8C8C", main, 0.08), from_hls(hh, 0.78, s * 0.7)]
    pal["updown"] = "arrow" if "red" in (family(main), family(acc)) or "green" in (family(main), family(acc)) else "color"
    return pal, cols

def _accent_from_logo(chroma, base, override=None):
    """多色 LOGO 的辅助色用其第二色，不因色相接近就改走自动强调色。

    第二色按取色面积：chroma[0] 最大，chroma[1] 为第二色。用户指定的主色若正好是第二色，
    则改用面积最大的另一彩色，仍然来自 LOGO，不调用 auto_accent。
    """
    if override:
        return _norm_hex(override), "指定强调色", len(chroma) >= 2
    if len(chroma) < 2:
        return None, None, False
    ordered = [c.upper() for c in chroma]
    base_u = base.upper()
    if ordered[1] != base_u:
        return ordered[1], "LOGO 第二色", True
    if ordered[0] != base_u:
        return ordered[0], "LOGO 另一彩色", True
    for c in ordered[2:]:
        if c != base_u:
            return c, "LOGO 第二色", True
    return None, None, False

def palette_pair(path=None, main=None, alt_main=None, accent=None):
    """确认（1）的全部配色候选；保留 A/B 字段兼容旧调用。

    多色时每种彩色各作一次主色，图形色先于可识别的文字色；推荐最大图形色。
    只有一种彩色时，方案 A 的强调色按主流配色从主色派生；方案 B 的主色不变，强调色改用同色相深色，不引入新色相。
    没有 LOGO 时，方案 B 用调用方传入的另一示例主色（alt_main）。
    """
    cols = logo_colors(path) if path else []
    chroma = [c for c, _, kind in cols if kind == "chroma"]
    if not path and not main:
        raise ValueError("需要 LOGO，或指定主色")
    if len(chroma) >= 2:
        details = logo_color_details(path)
        details.sort(key=lambda d: (d["primary_origin"] == "来自文字", -d["area_share"]))
        if main:
            chosen = _norm_hex(main)
            specified = next((d for d in details if d["color"] == chosen), None)
            if specified is None:
                specified = dict(color=chosen, area_share=None, area_basis="指定色", position="用户指定",
                                 primary_origin="指定主色", origin_method="用户指定")
            details = [specified] + [d for d in details if d["color"] != chosen]
        candidates = []
        for i, detail in enumerate(details):
            key = chr(ord("A") + i)
            info = swatch_palette(path, main=detail["color"], accent=accent if i == 0 else None)
            info.update(detail)
            info["key"] = key
            share = "指定色" if detail["area_share"] is None else "彩色面积 %.1f%%" % (100*detail["area_share"])
            info["label"] = "方案 %s：%s；%s；%s" % (key, detail["primary_origin"], detail["position"], share)
            candidates.append(info)
        recommended = next((d["key"] for d in candidates if d["primary_origin"] == "图形/色块"), candidates[0]["key"])
        for item in candidates:
            item["recommended"] = item["key"] == recommended
            if item["recommended"]:
                item["label"] += "（推荐）"
        pair = {d["key"]: d for d in candidates}
        pair.update(candidates=candidates, recommended=recommended,
                    a_label=candidates[0]["label"], b_label=candidates[1]["label"])
        return pair
    first = swatch_palette(path, main=main, accent=accent)
    second = None
    second_label = ""
    if len(chroma) >= 2:
        primary_now = first["primary_source"].upper()
        other = next((c for c in chroma if c.upper() != primary_now), None)
        if other:
            second = swatch_palette(path, main=other)
            second_label = "方案 B：主色改用 LOGO 的另一彩色，强调色仍取自 LOGO"
    if second is None and alt_main and _norm_hex(alt_main) != first["primary_source"].upper():
        second = swatch_palette(None, main=alt_main)
        second_label = "方案 B：没有可对调的 LOGO 第二色，这是另一套技能里已有的示例主色"
    if second is None:
        second = swatch_palette(path, main=first["primary_source"], accent=first["dark"])
        second["accent_from"] = "同色相深色"
        second["multi"] = False
        second_label = "方案 B：主色与方案 A 相同，强调色改用同色相深色，不引入新色相"
    if main:
        first_label = "方案 A：按指定主色"
    elif path and chroma:
        first_label = "方案 A：面积最大的彩色作主色"
    else:
        first_label = "方案 A：示例主色（不是从 LOGO 取出）"
    details = logo_color_details(path) if path and chroma else []
    for key, info, label in (("A", first, first_label), ("B", second, second_label)):
        detail = next((d for d in details if d["color"] == info["primary_source"]), None)
        if detail:
            info.update(detail)
        else:
            info.update(primary_origin="指定主色", position="用户指定", area_share=None, area_basis="指定色")
        info.update(key=key, label=label, recommended=key == "A")
    return {"A": first, "B": second, "a_label": first_label, "b_label": second_label,
            "candidates": [first, second], "recommended": "A"}


def swatch_palette(path=None, main=None, accent=None):
    """一套四色：主色、深色、浅底、强调色。确认时用 palette_pair() 出全部候选。

    深色、浅底与 palette_from_logo 同一套公式。多色 LOGO 的强调色是第二色本身
    （对比度不够时只在同色相内压暗）。单色 LOGO 沿用主色，无 LOGO 时可使用示例强调色。
    写成稿时把这里的主色来源和强调色来源传给 start(..., main=, accent=)，
    这样不会再按色相差被自动强调色换掉。
    """
    cols = logo_colors(path) if path else []
    chroma = [c for c, _, k in cols if k == "chroma"]
    if main:
        base = _norm_hex(main)
    elif chroma:
        base = chroma[0]
    else:
        raise ValueError("需要 LOGO，或指定主色")
    primary = darken_until(base, target=4.5)
    hh, l, s = hls(primary)
    dark = from_hls(hh, 0.20, min(1, s * 0.85))
    light = from_hls(hh, 0.955, min(0.55, s))
    acc_src, acc_from, multi = _accent_from_logo(chroma, base, accent)
    if acc_src is None:
        acc_src = auto_accent(primary)
        acc_from = "单色 LOGO，按主流配色派生" if path else "无 LOGO，冷暖对比"
        multi = False
    acc = darken_until(acc_src, target=3.0)
    close = bool(multi and hue_delta(base, acc_src) <= 60)
    return dict(
        primary=primary, dark=dark, light=light, accent=acc,
        primary_source=base, accent_source=acc_src, accent_from=acc_from,
        multi=multi, close_hues=close, chroma=chroma, logo_colors=cols,
    )

def report(pal):
    rows = [("白字 / 主色", contrast("FFFFFF", pal["navy"]), 4.5),
            ("白字 / 深色", contrast("FFFFFF", pal["dark"]), 4.5),
            ("白字 / 强调色（大号粗体）", contrast("FFFFFF", pal["peach"]), 3.0),
            ("正文 / 浅底", contrast(pal["ink"], pal["tint"]), 7.0),
            ("主色字 / 白底", contrast(pal["navy"], "FFFFFF"), 4.5),
            ("次要字 / 白底", contrast(pal["muted"], "FFFFFF"), 4.5)]
    return [(n, round(v, 2), t, "达标" if v >= t else "不达标") for n, v, t in rows]

if __name__ == "__main__":
    import sys
    p, cols = palette_from_logo(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else None)
    print("LOGO 取色：", cols)
    for k, v in p.items(): print(f"  {k:8s} {v}")
    for r in report(p): print("  ", r)

def white_logo(path, out=None, tol=40):
    """扣出 LOGO 主体，变成白色、背景透明（用于深色底封面）。
    有透明通道且四角透明：按原透明度；否则取四角颜色当背景，按与背景的色差算透明度（边缘抗锯齿平滑过渡）。"""
    import os, tempfile
    im = Image.open(path).convert("RGBA"); w, h = im.size
    px = im.load()
    corners = [px[0, 0], px[w - 1, 0], px[0, h - 1], px[w - 1, h - 1]]
    use_alpha = all(c[3] < 20 for c in corners)
    bg = tuple(sorted(c[i] for c in corners)[1] for i in range(3))
    outim = Image.new("RGBA", (w, h))
    op = outim.load()
    for y in range(h):
        for x in range(w):
            r, g, b, a = px[x, y]
            if use_alpha:
                al = a
            else:
                d = max(abs(r - bg[0]), abs(g - bg[1]), abs(b - bg[2]))
                al = 0 if d < tol * 0.5 else min(255, int((d - tol * 0.5) / tol * 255)) * a // 255
            op[x, y] = (255, 255, 255, al)
    bbox = outim.getbbox()
    if bbox: outim = outim.crop(bbox)
    out = out or os.path.join(tempfile.mkdtemp(), "logo-white.png")
    outim.save(out)
    return out

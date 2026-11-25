# -*- coding: utf-8 -*-
"""由主色推导整套配色（规律取自 42 套参考配色的实测色值，见《配色规律说明》）。

用法：
    import palette
    pal = palette.derive("0E5A78")                 # 只给主色
    pal = palette.derive("C8102E", secondary="..."，scheme="A", corner="round")
    palette.apply(deckkit_module, pal)             # 写入引擎的全局颜色
"""
import colorsys


def _rgb(h):
    return [int(h[i:i + 2], 16) / 255 for i in (0, 2, 4)]


def _hex(r, g, b):
    return "%02X%02X%02X" % tuple(max(0, min(255, round(v * 255))) for v in (r, g, b))


def hls(h):
    return colorsys.rgb_to_hls(*_rgb(h))


def from_hls(h, l, s):
    return _hex(*colorsys.hls_to_rgb(h % 1.0, max(0, min(1, l)), max(0, min(1, s))))


def mix(a, b, k):
    A, B = _rgb(a), _rgb(b)
    return _hex(*[x + (y - x) * k for x, y in zip(A, B)])


def _lum(hx):
    def ch(c):
        return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
    r, g, b = (ch(v) for v in _rgb(hx))
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def contrast(a, b):
    la, lb = sorted((_lum(a), _lum(b)), reverse=True)
    return (la + 0.05) / (lb + 0.05)


def family(h):
    """按色相分三类：冷色（蓝/青/蓝紫）、暖色（红/橙/黄/玫紫）、绿色。"""
    d = h * 360
    if 170 <= d < 300:          # 300°–330° 的玫紫、梅紫归暖色，浅底用米白，避免整页发粉
        return "cool"
    if 75 <= d < 170:
        return "green"
    return "warm"


def derive(main, secondary=None, scheme=None, corner="square"):
    """
    main       主色（取自 LOGO）
    secondary  LOGO 自带的第二色；没有则按色相族自动推强调色
    scheme     A = 浅色页底 + 白卡片；B = 白页底 + 浅色卡片；None 自动（冷色 A，暖色/绿色 B）
    corner     square / round，整套统一
    """
    h, l, s = hls(main)
    fam = family(h)
    if scheme is None:
        scheme = "A" if fam == "cool" else "B"

    # 1) 过于荧光的主色先降一点饱和，再压暗：顺序反过来会让压暗后的颜色又变亮，白字对比度跌回 3 以下
    if s > 0.9 and l > 0.35:
        main = from_hls(h, l, 0.85)
        h, l, s = hls(main)
    # 主色过浅时逐步压暗，直到白字对比度 ≥ 3（色条上的白色粗体字可读；参考稿主色亮度集中在 0.15–0.50）
    while contrast(main, "FFFFFF") < 3.0 and l > 0.1:
        l -= 0.02
        main = from_hls(h, l, s)
    h, l, s = hls(main)

    # 2) 浅色底：冷色系用主色同色相的冷白；暖色系不用粉色，改用米白/暖灰；绿色系用带一点主色的灰白
    if fam == "cool":
        tint_h, tint_s = h, 0.38
    elif fam == "warm":
        tint_h, tint_s = 40 / 360, 0.22
    else:
        tint_h, tint_s = h, 0.10
    page_tint = from_hls(tint_h, 0.955, tint_s)       # 实测 0.94–0.96
    card_tint = from_hls(tint_h, 0.935, tint_s)       # 实测 0.91–0.97
    border = from_hls(tint_h, 0.885, tint_s * 0.8)

    if scheme == "A":
        bg, card_a, card_b = page_tint, "FFFFFF", "FFFFFF"
    else:
        bg, card_a, card_b = "FFFFFF", card_tint, from_hls(tint_h, 0.965, tint_s)

    # 3) 强调色：有 LOGO 第二色用第二色；否则冷色配暖橙，暖色/绿色配金色
    if secondary:
        acc = secondary
    elif fam == "cool":
        acc = from_hls(24 / 360, 0.56, 0.70)       # 暖橙
    else:
        acc = from_hls(40 / 360, 0.52, 0.55)       # 浅金
    # 主色本身是橙/黄时，强调色改用低饱和灰蓝（只作"需关注/对比组"，不作重点），
    # 不用深墨色：深墨和正文同色，标出来像"黑掉"而不是被强调
    if fam == "warm" and 15 / 360 < h < 60 / 360:
        acc = from_hls(205 / 360, 0.46, 0.23)

    # 4) 文字灰阶：带一点主色色相的近黑与灰（实测正文色多为带色相的深灰）
    ink = from_hls(h, 0.14, 0.22)
    sub = from_hls(h, 0.30, 0.12)
    muted = from_hls(h, 0.50, 0.08)
    line = from_hls(tint_h, 0.86, tint_s * 0.7)

    # 5) 深色块与渐变：主色 → 同色相略亮（幅度小，接近平涂）
    dark_end = from_hls(h + 0.02, min(l + 0.12, 0.55), s * 0.9)
    on_dark_sub = from_hls(h, 0.86, 0.30)          # 深色块上的说明文字
    track = from_hls(tint_h, 0.90, tint_s)         # 条形图底槽

    # 重点 = 主色，非重点 = 带主色色相的灰（明度 0.72）；环形图其余扇区按灰阶由深到浅
    norm = from_hls(h, 0.72, 0.10)
    grays = [from_hls(h, l_, 0.08) for l_ in (0.52, 0.72, 0.83, 0.90)]
    chart = [main] + grays
    return dict(main=main, accent=acc, scheme=scheme, corner=corner, family=fam,
                bg=bg, card_a=card_a, card_b=card_b, border=border,
                ink=ink, sub=sub, muted=muted, line=line,
                dark=(main, dark_end, 315), head=(main, dark_end, 0),
                on_dark_sub=on_dark_sub, track=track, chart=chart, norm=norm, grays=grays,
                stripe=mix(card_tint, "FFFFFF", .5))


def apply(K, pal):
    """把配色写入 deckkit 的全局变量（组件在运行时读取）。"""
    K.MAIN, K.ACC = pal["main"], pal["accent"]
    K.INK, K.SUB, K.MUTED, K.LINE = pal["ink"], pal["sub"], pal["muted"], pal["line"]
    K.BG = pal["bg"]
    K.DARK_G, K.HEAD_G = pal["dark"], pal["head"]
    K.CELL_G = (pal["card_a"], pal["card_b"], 0)
    K.CARD_BORDER = pal["border"]
    K.ONDARK = pal["on_dark_sub"]
    K.TRACK = pal["track"]
    K.STRIPE = pal["stripe"]
    K.NORM, K.GRAYS = pal["norm"], pal["grays"]
    K.CORNER = pal["corner"]
    K.PALETTE = pal


# ================= 预设色系（不让用户自选主色，由 SKILL 指定） =================
# main 取自参考稿实测值附近，略作统一；accent=None 表示按规则自动推
PRESETS = {
    "藏蓝":   dict(main="0B2F6B", accent="C9A35B", scheme="A", use="金融、银行、政务、通用汇报（默认）"),
    "商务蓝": dict(main="1F5AA6", accent="E07B39", scheme="A", use="通用商务、制造、央国企"),
    "科技蓝": dict(main="0E6FB8", accent="F08A3C", scheme="A", use="科技、互联网、医疗"),
    "青蓝":   dict(main="0E5A78", accent="DD7F40", scheme="A", use="物流、能源、水务、医药"),
    "青绿":   dict(main="0B7A75", accent="C9A35B", scheme="B", use="环保、医疗、党建（青绿）"),
    "墨绿":   dict(main="1F5C4A", accent="C9A35B", scheme="B", use="农业、林业、地产、绿色金融"),
    "政务红": dict(main="B3171E", accent="C9A35B", scheme="B", use="党政、党建、国企总结"),
    "酒红":   dict(main="7A1E2B", accent="C9A35B", scheme="B", use="高端商务、白酒、消费品"),
    "活力橙": dict(main="E2671A", accent="5B7A90", scheme="B", use="零售、营销、互联网活动"),
    "典雅紫": dict(main="4B3A8C", accent="D9A441", scheme="A", use="医疗、教育、美妆"),
    "炭灰":   dict(main="2E3A46", accent="E07B39", scheme="B", use="中性商务、咨询、工业（无品牌色时）"),
}


# ================= 扩充色库：没有 LOGO 时按行业多给几种主色，避免套套都是同一个蓝 =================
# accent / scheme 不写时按 derive() 规则自动推；主色亮度控制在 0.2–0.5、饱和度 0.3–0.8 之间，保证白字可读、不发荧光
EXTRA = {
    "深海蓝": dict(main="0A3D62", use="金融、银行、科技、能源、政务"),
    "湖蓝":   dict(main="1B6CA8", use="科技、互联网、教育、通用商务"),
    "靛青":   dict(main="2C3E8C", use="金融、银行、教育、政务、咨询"),
    "钢蓝":   dict(main="3A6186", use="制造、能源、工业、央国企、物流"),
    "孔雀蓝": dict(main="0F6E8C", use="医药、医疗、物流、水务、互联网"),
    "松石绿": dict(main="138A7E", use="医疗、医药、环保、新能源、乳业"),
    "森林绿": dict(main="2E6B3F", use="农业、林业、环保、绿色金融"),
    "橄榄绿": dict(main="5B6B2E", use="农业、食品、乳业、文旅"),
    "砖红":   dict(main="A23B2A", use="制造、地产、建筑、国企总结"),
    "朱红":   dict(main="C0392B", use="党建、消费品、零售、文化、食品"),
    "枣红":   dict(main="8E2A3B", use="白酒、高端商务、消费品、酒店"),
    "琥珀橙": dict(main="D97B1E", use="零售、食品、乳业、营销"),
    "赭石":   dict(main="A0522D", use="文化、文旅、地产、建筑"),
    "金棕":   dict(main="9C6B2F", use="高端商务、文旅、酒店、白酒"),
    "葡萄紫": dict(main="5E3A87", use="美妆、教育、医疗、文化"),
    "梅紫":   dict(main="7B2D5E", use="美妆、文化、消费品、营销"),
    "石墨蓝": dict(main="34495E", use="咨询、工业、中性商务、律所"),
}
COLORS = {**PRESETS, **EXTRA}


def _keys(use):
    return [k for k in use.replace("（", "、").replace("）", "").replace("，", "、").split("、") if k]


def preset(name, corner="square"):
    """按色系名或主色值（"#1E50A2" / "1E50A2"）推导整套配色。"""
    if name in COLORS:
        p = COLORS[name]
        pal = derive(p["main"], secondary=p.get("accent"), scheme=p.get("scheme"), corner=corner)
        pal["preset"] = name
        return pal
    hx = str(name).lstrip("#").upper()
    if len(hx) == 6 and all(c in "0123456789ABCDEF" for c in hx):
        return from_color(hx, corner=corner)
    raise KeyError(name)


def is_hex(x):
    hx = str(x).lstrip("#")
    return len(hx) == 6 and all(c in "0123456789ABCDEFabcdef" for c in hx)


def usable(hx):
    """主色能不能直接用：太灰、太黑、太白的 LOGO 色不直接用。"""
    h, l, s = hls(hx)
    return s >= 0.15 and 0.08 <= l <= 0.92


def from_color(hx, corner="square"):
    """直接用一个主色（通常是 LOGO 原色）推导整套配色；太灰/太黑/太白时退回最接近的预设。"""
    hx = hx.lstrip("#").upper()
    if not usable(hx):
        pal = preset(nearest_preset(hx), corner=corner)
        pal["basis_note"] = "LOGO 色 #%s 过灰或过深，改用 %s" % (hx, pal["preset"])
        return pal
    pal = derive(hx, corner=corner)
    pal["preset"] = "#" + hx
    return pal


def main_of(choice):
    """色系名或色值 → 主色（显示用）。色值会先过 derive() 的压暗规则。"""
    return preset(choice)["main"]


def label(choice):
    return ("LOGO原色 " + choice) if str(choice).startswith("#") else choice


def option(choice):
    """选择框里一项的写法：色系名＋主色；LOGO 原色被压暗过时注明调整后的色值。"""
    m = main_of(choice)
    if str(choice).startswith("#"):
        raw = choice.lstrip("#").upper()
        return "LOGO原色 #%s" % raw if m.upper() == raw else "LOGO原色 #%s（为白字可读微调为 #%s）" % (raw, m)
    return "%s #%s" % (choice, m)


def by_industry(industry, seed=""):
    """按行业关键词列出候选色（预设在前、扩充在后），再按 seed（一般是标题）稳定地挑一个打头。"""
    import hashlib
    cands = [n for n, p in COLORS.items() if any(k in industry for k in _keys(p["use"]))]
    if not cands:
        return []
    i = int(hashlib.md5((seed or industry).encode("utf-8")).hexdigest(), 16) % len(cands)
    return cands[i:] + cands[:i]


def alternatives(name, k=2, industry="", seed=""):
    """给推荐色配 k 个备选：同行业的其他候选优先，不够时补色相相近的。"""
    m0 = main_of(name)
    h0, l0, s0 = hls(m0)

    def same(n):  # 与推荐色几乎同色的不当备选（如推荐 LOGO 原色 #1B6CA8 时不再给“湖蓝”）
        h, l, s = hls(main_of(n))
        return min(abs(h - h0), 1 - abs(h - h0)) * 360 < 8 and abs(l - l0) < 0.04 and abs(s - s0) < 0.12
    out = [n for n in by_industry(industry, seed) if n != name and not same(n)][:k] if industry else []
    ds = []
    for n in COLORS:
        if n in out or n == name or n == "炭灰" or same(n):
            continue
        h, l, s = hls(COLORS[n]["main"])
        ds.append((min(abs(h - h0), 1 - abs(h - h0)) * 360 / 30 + abs(l - l0) * 2.5, n))
    out += [n for _, n in sorted(ds)][:k - len(out)]
    return out[:k]


def nearest_preset(logo_hex):
    """LOGO 颜色 → 最接近的预设色系（按色相为主、亮度饱和度为辅）。低饱和的 LOGO 归到炭灰。"""
    h, l, s = hls(logo_hex)
    if s < 0.15 or l < 0.08:
        return "炭灰"
    if 225 / 360 <= h < 250 / 360:      # 靛蓝归藏蓝，避免被判成紫色
        return "藏蓝"
    best, bd = None, 9
    for name, p in PRESETS.items():
        if name == "炭灰":
            continue
        ph, pl, ps = hls(p["main"])
        dh = min(abs(h - ph), 1 - abs(h - ph)) * 360 / 30
        d = dh + abs(l - pl) * 2.5 + abs(s - ps) * 0.8
        if d < bd:
            best, bd = name, d
    return best

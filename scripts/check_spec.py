# -*- coding: utf-8 -*-
"""deck.json 规格检查：在出图前拦住违反版面切分规格、组件容量和数字出处的页面。

python3 check_spec.py deck.json [--evidence evidence.md]
退出码：0 通过；1 有错误；2 文件格式错误。规则对应 references/split.md 与 components.md。
"""
import argparse
import json
import math
import os
import re
import sys
from collections import Counter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

BODY = (48, 116, 1184, 558)
SEPS = ("纯卡片", "纯线条", "线卡混用", "创意型")
TYPES = {"statement", "points", "para", "feature", "tags", "tip", "split_note", "steps", "chain", "compare",
         "icon_steps", "num_pair", "stack", "kpi_row", "table", "bar_h", "column", "line", "donut", "waterfall", "bar_cards", "photo", "asset"}
# 暂停使用的组件：规格检查直接报错并给出替代写法
DISABLED = {"waterfall": "瀑布图暂停使用（合计柱在部分渲染环境下画不出来）。增减拆分各项都为正时用 bar_h，"
                         "按贡献从大到小排序，起止合计写进标题或 kpi_row；只要有一项为负（减少、拖累），用 table 列出正负项"}
CARD_STYLES = {None, "plain", "header", "dark", "number", "label"}
CAP = {"points": 4, "steps": 5, "chain": 3, "compare": 4, "icon_steps": 5, "num_pair": 3, "kpi_row": 5,
       "tags": 8, "bar_h": 8, "column": 10, "line": 12, "donut": 6, "waterfall": 7, "bar_cards": 6}
MIN_W, MIN_H = 240, 130
# 整稿版式变化（正文页 ≥6 页时检查）：两格页上限、同一骨架上限、4 格以上页下限（≥8 页时）
TWO_CELL_MAX = 0.4
SKELETON_MAX = 0.3
FOUR_CELL_MIN = 0.25
GRID_MAX = 0.2          # 田字格（一刀两分、两边各再一刀两分）占正文页上限
MAIN_BLOCK_MIN = 0.25   # 至少这么多正文页有一个不再细分的大主块
MAIN_BLOCK_AREA = 0.4   # 大主块：单格面积占正文区的比例
EVEN_CUT_MAX = 0.5      # 两分刀里接近对半（小的一段 ≥45%）的上限
CELL_GAP = {"纯卡片": 20}  # 与 build.GAP 一致：纯卡片页格子之间留 20 px，线条页为 0
GRIDS = ("v2(h2(x,x),h2(x,x))", "h2(v2(x,x),v2(x,x))")
# 格内注释写成口径、来源说明时，应进页脚，不要在格子底部用细线隔出一个小格子来放
SOURCE_NOTE = re.compile(r"^(口径|来源|资料来源|数据来源|注[:：])|未经审计|管理口径|时点数|按[^，。；]{0,6}统计")


def _layout(node, x, y, w, h, gap, leaves, depth=0, first=None):
    if "children" not in node:
        leaves.append((node, (x, y, w, h)))
        return
    rs = node["ratio"]
    if len(rs) != len(node["children"]):
        raise ValueError("ratio 与 children 数量不一致")
    tot = (w if node["cut"] == "v" else h) - gap * (len(rs) - 1)
    pos = x if node["cut"] == "v" else y
    for r, ch in zip(rs, node["children"]):
        size = tot * r / sum(rs)
        if node["cut"] == "v":
            _layout(ch, pos, y, size, h, gap, leaves, depth + 1)
        else:
            _layout(ch, x, pos, w, size, gap, leaves, depth + 1)
        pos += size + gap


def _is_photo(node):
    return "children" not in node and (node.get("element") or {}).get("type") == "photo"


def _bad_cuts(node):
    """切分树里不合法的 cut 值（只能是 v 或 h）。"""
    if not isinstance(node, dict) or "children" not in node:
        return []
    own = [] if node.get("cut") in ("v", "h") else [str(node.get("cut"))]
    return own + [c for ch in node["children"] for c in _bad_cuts(ch)]


def _cuts(node):
    """分隔线条数。与图片格相邻的那一刀不画线（图片本身就是分隔），不计入。"""
    if "children" not in node:
        return 0
    kids = node["children"]
    own = sum(1 for a, b in zip(kids, kids[1:]) if not (_is_photo(a) or _is_photo(b)))
    return own + sum(_cuts(c) for c in kids)


def _even_cuts(node, out):
    """收集每一刀两分的较小一段占比；等分三四格是有意等分，不计。"""
    if "children" not in node:
        return
    rs = node.get("ratio") or []
    if len(rs) == 2 and sum(rs) > 0:
        out.append(min(rs) / float(sum(rs)))
    for c in node["children"]:
        _even_cuts(c, out)


def _signature(node):
    if "children" not in node:
        return "x"
    return node["cut"] + "[" + ",".join(str(r) for r in node["ratio"]) + "](" + ",".join(_signature(c) for c in node["children"]) + ")"


def _shape(node, typed=True):
    """不看比例的结构指纹：切法、格数和（可选）每格组件类型。比例微调不算换了版面。"""
    if "children" not in node:
        if not typed:
            return "x"
        el = node.get("element") or {}
        return el.get("type", "?")
    return node["cut"] + str(len(node["children"])) + "(" + ",".join(_shape(c, typed) for c in node["children"]) + ")"


# 页面上不能出现的过程用语：这些是写底稿、做核对时的词，不是汇报内容
META_WORDS = ("核查依据", "核查边界", "事实与核查", "数据依据", "论据", "主证据", "辅证据", "底稿", "待核查")
TEXT_ONLY = {"points", "para", "statement", "tags", "tip"}


def _labels(part):
    """页面上看得见的标签：格子小标题、要点导语、结论句。"""
    out = []

    def walk(n):
        if "children" in n:
            for c in n["children"]:
                walk(c)
            return
        if n.get("title"):
            out.append(n["title"])
        for e in _elements(n.get("element") or {}):
            if e.get("type") == "points":
                out.extend(it[0] for it in e.get("items", []) if it)
            if e.get("type") == "statement":
                out.append(e.get("text", ""))
    walk(part)
    return out


def _has_icon(el):
    if el.get("type") in ("points", "steps", "icon_steps"):
        return any(len(it) > 2 and it[2] for it in el.get("items", []))
    return el.get("type") in ("feature", "tip")


def _elements(el):
    yield el
    if el.get("type") == "stack":
        for sub, _ in el.get("items", []):
            yield from _elements(sub)
    if el.get("type") == "split_note" and el.get("main"):
        yield from _elements(el["main"])


def _texts(obj):
    if isinstance(obj, str):
        yield obj
    elif isinstance(obj, dict):
        for k, v in obj.items():
            if k not in ("type", "icon", "topic", "anchor", "kind", "card", "fill", "role", "evidence", "cut"):
                yield from _texts(v)
    elif isinstance(obj, (list, tuple)):
        for v in obj:
            yield from _texts(v)


NUM = re.compile(r"\d[\d,]*\.?\d*")


def _num_ok(n, ev):
    raw = n.replace(",", "")
    if len(raw.replace(".", "")) < 2:
        return True
    plain = ev.replace(",", "")
    return raw in plain


def _touches_corner(box):
    """图片格至少挨着正文区的一个角：左右必有一边贴正文区边缘，上下也必有一边贴边。"""
    x, y, w, h = box
    bx, by, bw, bh = BODY
    side = abs(x - bx) < 1.5 or abs(x + w - bx - bw) < 1.5
    edge = abs(y - by) < 1.5 or abs(y + h - by - bh) < 1.5
    return side and edge


def _full_strip(box):
    """线卡混用的大卡片要占可切割区完整的一整条：通高贴左或右边，或通宽贴上或下边，即同时挨着两个角。"""
    x, y, w, h = box
    bx, by, bw, bh = BODY
    left, right = abs(x - bx) < 1.5, abs(x + w - bx - bw) < 1.5
    top, bottom = abs(y - by) < 1.5, abs(y + h - by - bh) < 1.5
    return (top and bottom and (left or right)) or (left and right and (top or bottom))


def _photo_ok(topic, online):
    """离线时图片关键词必须在自带图库里找得到相关的图；联网时交给生成阶段检索。"""
    if online:
        return True
    try:
        import photos
    except Exception:  # 缺 Pillow 等依赖时不在这里拦
        return True
    return photos.pick(topic) is not None


def check_page(spec, i, ev=None, online=False):
    err, warn = [], []
    tag = "第 %d 个正文页「%s」" % (i, spec.get("title", "")[:18])
    if not spec.get("title"):
        err.append(tag + "：缺页标题")
    sep = spec.get("separation")
    if sep not in SEPS:
        err.append(tag + "：separation 只能是 " + "/".join(SEPS))
        return err, warn
    if sep == "创意型" or "creative" in spec:
        cr = spec.get("creative") or {}
        if cr.get("kind") == "asset":
            import assets
            err += [tag + "：" + e for e in assets.check(cr)]
        elif cr.get("kind") == "radial":
            err.append(tag + "：中心辐射（radial）暂停使用（四张小卡悬在大片留白里，版面松散）。"
                       "并列的四项改用纯卡片切分：先横切标题区或结论条，再竖切四栏或切成 2×2；"
                       "确需一个中心带多个要点时用原版创意素材 B3-052（中心 + 六要点）")
        elif cr.get("kind") != "timeline":
            err.append(tag + "：创意型只支持 timeline、asset（原版创意素材，见 scripts/assets.py 列表）")
        return err, warn
    part = spec.get("partition")
    if not part:
        err.append(tag + "：缺 partition")
        return err, warn
    bad = _bad_cuts(part)
    if bad:
        err.append(tag + "：切分方向 cut 只能是 v（左右切）或 h（上下切），现在写的是 %s" % "、".join(sorted(set(bad))))
        return err, warn
    leaves = []
    try:
        _layout(part, *BODY, {"纯卡片": 20}.get(sep, 0), leaves)
    except (KeyError, ValueError) as ex:
        err.append(tag + "：切分树格式错误（%s）" % ex)
        return err, warn
    area = BODY[2] * BODY[3]
    if not 2 <= len(leaves) <= 8:
        err.append(tag + "：一页 2～8 格，现在 %d 格" % len(leaves))
    darks = [b for lf, b in leaves if lf.get("fill") == "重点" or lf.get("card") == "dark"]
    if len(darks) > 1:
        err.append(tag + "：深色重点卡每页最多 1 个")
    photos = 0
    for lf, (x, y, w, h) in leaves:
        el = lf.get("element") or {}
        name = lf.get("title") or el.get("type", "")
        if lf.get("card") not in CARD_STYLES:
            err.append(tag + "：未知卡片样式 %s" % lf.get("card"))
        if w < MIN_W or h < MIN_H:
            err.append(tag + "：格子「%s」太小 %.0f×%.0f（最小 %d×%d）" % (name, w, h, MIN_W, MIN_H))
        for e in _elements(el):
            t = e.get("type")
            if t in DISABLED:
                err.append(tag + "：「%s」%s" % (name, DISABLED[t]))
                continue
            if t not in TYPES:
                err.append(tag + "：未知组件 %s" % t)
                continue
            items = e.get("items") or e.get("rows") or e.get("cats") or e.get("steps") or []
            if t in CAP and len(items) > CAP[t]:
                err.append(tag + "：「%s」的 %s 有 %d 项，超过上限 %d，请拆格或拆页" % (name, t, len(items), CAP[t]))
            if t in ("column", "line", "bar_h"):
                ser = e.get("series") or [{"vals": e.get("vals") or []}]
                if t == "line" and len(ser) > 3:
                    err.append(tag + "：折线图「%s」有 %d 条线，最多 3 条（第一条为重点线）" % (name, len(ser)))
                for s_ in ser:
                    if len(s_.get("vals") or []) != len(e.get("cats") or []):
                        err.append(tag + "：图表「%s」的数值个数与类别个数不一致" % name)
                        break
                if t in ("bar_h", "column"):
                    neg = [v for s_ in ser for v in (s_.get("vals") or []) if isinstance(v, (int, float)) and v < 0]
                    if neg:
                        err.append(tag + "：图表「%s」有负值 %s；条形图和柱形图的数值轴从 0 起，负值画不出来。"
                                   "含负值（减少、拖累）的拆分改用 table 列出正负项" % (name, "、".join("%g" % v for v in neg[:3])))
            if t in ("table",) and h < 200:
                err.append(tag + "：表格格子高度 %.0f 不足 200" % h)
            if t == "asset":
                import assets
                err += [tag + "：" + m for m in assets.check(e)]
                if e is not el:
                    err.append(tag + "：原版素材只能单独占一格")
                if w < 400 or h < 260:
                    err.append(tag + "：素材格子太小 %.0f×%.0f（至少 400×260）" % (w, h))
            if t == "photo" and e.get("caption"):
                err.append(tag + "：正文配图不加文字说明（caption「%s」），删掉" % e["caption"])
            note = e.get("note") if t in ("num_pair", "split_note", "compare") else None
            if note and SOURCE_NOTE.search(note):
                err.append(tag + "：「%s」的注释是口径/来源说明（%s…）；写进本页 footer，不要占格子底部" % (name, note[:16]))
            if t == "photo" and e is not el:
                err.append(tag + "：图片只能单独占一格，不要放进 stack、split_note")
            if t == "photo":
                photos += 1
                if w * h > area / 3 + 1:
                    err.append(tag + "：图片格面积超过正文区 1/3")
                if not _touches_corner((x, y, w, h)):
                    err.append(tag + "：图片格要挨着正文区的一个角（贴左或右边缘，同时贴上或下边缘），不要悬在页面中间")
                if not e.get("topic"):
                    err.append(tag + "：图片格缺 topic（写图库文件名、本机图片路径或中文关键词）")
                elif not _photo_ok(e["topic"], online):
                    err.append(tag + "：自带图库里没有与「%s」相关的图；用 scripts/photos.py 搜索 换关键词、直接写文件名，"
                               "或开启联网取图；都没有就不要放图片格" % e["topic"])
    if photos > 1:
        err.append(tag + "：每页最多 1 个图片格")
    if sep == "纯线条":
        if darks or any(lf.get("card") for lf, _ in leaves):
            err.append(tag + "：纯线条页不能有卡片，改用线卡混用")
        if _cuts(part) > 2:
            err.append(tag + "：纯线条页最多 2 条分隔线，现在 %d 条" % _cuts(part))
        rs = part["ratio"]
        if len(rs) == 2 and max(rs) / max(min(rs), 1e-6) < 1.2:
            err.append(tag + "：纯线条页第一刀要不对称（如 58:42、2:1）")
    if sep == "线卡混用":
        big = [b for b in darks if b[2] * b[3] >= area / 3 - 1]
        cards = [lf for lf, _ in leaves if lf.get("fill") == "重点" or lf.get("card")]
        if not big:
            err.append(tag + "：线卡混用页必须有 1 张深色大卡片，面积不小于正文区 1/3")
        elif not any(_full_strip(b) for b in big):
            err.append(tag + "：线卡混用页的大卡片要占可切割区完整的一整条（通高靠左/右，或通宽靠上/下），同时挨着两个角；"
                       "把大卡片放在第一刀切出的那一段，不要再横切或竖切它所在的一侧")
        if len(cards) > 2:
            err.append(tag + "：线卡混用页除大卡片外最多 1 张小卡片")
    # 过程用语与重复标签：页面是汇报，不是核对记录
    labels = _labels(part)
    meta = sorted({w for t in labels for w in META_WORDS if w in t})
    if meta:
        err.append(tag + "：页面上出现过程用语「%s」；格子小标题和要点导语要写成本页的判断或事项，不写核对过程" % "、".join(meta))
    leads = [t for t in labels if t]
    dup = sorted({t for t in leads if leads.count(t) > 1})
    if dup:
        err.append(tag + "：同一页里重复的小标题或导语「%s」；每条导语写出各自的内容" % "、".join(d[:12] for d in dup))
    # 整页只有文字、没有图标和关系图示，看起来像 Word
    els = [e for lf, _ in leaves for e in _elements(lf.get("element") or {})]
    flat = [e for e in els if e.get("type") not in ("stack", "split_note")]
    try:
        import icons as _I
        unknown = []
        for e in flat:
            for sub in _elements(e):
                names = [it[2] for it in sub.get("items", []) if sub.get("type") in ("points", "steps", "icon_steps")
                         and isinstance(it, (list, tuple)) and len(it) > 2 and it[2]]
                if sub.get("type") == "feature" and sub.get("icon"):
                    names.append(sub["icon"])
                unknown += [n for n in names if isinstance(n, str) and not _I.resolve(n)]
        if unknown:
            warn.append(tag + "：图标「%s」在 IconPark 和中文词表里都找不到，已用通用图标代替；换成贴合内容的 IconPark 名"
                        "（如 building-one、stock-market、calendar）或中文关键词" % "、".join(dict.fromkeys(unknown)))
    except Exception:
        pass
    if flat and all(e.get("type") in TEXT_ONLY for e in flat) and not any(_has_icon(e) for e in flat) \
            and not any(lf.get("fill") == "重点" or lf.get("card") in ("dark", "number", "label") for lf, _ in leaves):
        err.append(tag + "：整页只有文字段落，没有图标、关系图示、数字或重点卡；至少把一格换成 steps / icon_steps / chain / "
                   "num_pair / kpi_row / 深色结论卡，或给要点加图标")
    # 表格只需要放下文字的高度，格子明显更高时把空间让给别的组件
    for lf, (x, y, w, h) in leaves:
        el = lf.get("element") or {}
        if el.get("type") == "table":
            need = (len(el.get("rows", [])) + 1) * 36 + (50 if lf.get("title") else 0) + 16
            if h - need > 80:
                err.append(tag + "：表格「%s」只需约 %d px 高，格子有 %.0f px，下方会空一大块；调小这一格的 ratio，"
                           "把空间给相邻格" % (lf.get("title") or "", need, h))
            elif h - need > 45:
                warn.append(tag + "：表格「%s」只需约 %d px 高，格子有 %.0f px；表格按内容高度画，下方会空出来，"
                            "收格并把空间给指标带、结论条或关系图示" % (lf.get("title") or "", need, h))
    # 条形图按条数取高（每条约 44px），格子明显更高时下面会空
    for lf, (x, y, w, h) in leaves:
        el = lf.get("element") or {}
        if el.get("type") == "bar_h":
            need = len(el.get("cats", [])) * 44 + 24 + (50 if lf.get("title") else 0) + 30
            if h - need > 90:
                err.append(tag + "：条形图「%s」%d 条只需约 %d px 高，格子有 %.0f px；条形图不拉高，"
                           "调小这一格的 ratio，空出的位置放挨角的图片格或其他组件" % (lf.get("title") or "", len(el.get("cats", [])), need, h))
    # 内容量与切法：正文过少、只切一刀等分，都会让页面发空
    chars = len(re.sub(r"\s", "", "".join(_texts(part))))
    if chars < 160:
        warn.append(tag + "：正文只有 %d 字（不少于 160 字）；补论据、加指标带或结论条，或与相邻页合并" % chars)
    rs = part.get("ratio") or []
    if _cuts(part) == 1 and len(rs) == 2 and rs[0] == rs[1]:
        warn.append(tag + "：只切一刀且左右/上下等分；除前后对比、方案对比外，按主次切成不等分，"
                    "并用指标带、结论条或图表再切一刀（见 references/split.md 常用骨架）")
    foot = spec.get("footer") or ""
    if len(foot) > 150:
        warn.append(tag + "：页脚 %d 字，一行放不下；只留材料名和章节，口径说明精简" % len(foot))
    if ev:
        bad = sorted({n for s in _texts(spec.get("partition")) for n in NUM.findall(s) if not _num_ok(n, ev)})
        if bad:
            warn.append(tag + "：这些数字在论据底稿里找不到，请核对出处：" + "、".join(bad[:8]))
    return err, warn


def check_deck(deck, evidence_text=None):
    err, warn = [], []
    pages = deck.get("pages") or []
    if not pages:
        return ["deck.json 没有 pages"], []
    st = deck.get("settings", {})
    if not st.get("title"):
        err.append("settings.title（封面标题）不能为空")
    elif st.get("org"):
        short = re.split(r"[｜|]", st["org"])[0]
        short = re.sub(r"(股份)?有限(责任)?公司|集团|控股|科技|技术|实业|发展", "", short).strip()
        if short and short[:2] not in st["title"]:
            err.append("封面主标题「%s」里没有单位名；把「%s」写进 settings.title（如「%s2025年度工作汇报」）"
                       % (st["title"].replace("\n", " "), short, short))
    if not deck.get("sections"):
        warn.append("没有 sections：将不生成目录和章节转场")
    names = {s.get("name") for s in deck.get("sections", [])}
    import structure
    sty = st.get("styles") or {}
    if sty and st.get("styles_from") != "user":
        warn.append("settings.styles 没有标 \"styles_from\": \"user\"，构建时会被忽略、改为按标题随机抽版式；"
                    "首页、目录、转场版式由脚本抽，只有用户明确要求某种版式时才写 styles")
        sty = {}  # 会被忽略的版式不再校验，免得旧值（如已删除的 dark）挡住构建
    for key, allowed in (("cover", structure.COVER_STYLES), ("agenda", structure.AGENDA_STYLES), ("transition", structure.TRANSITION_STYLES)):
        if sty.get(key) and sty[key] not in allowed:
            err.append("settings.styles.%s 只能是 %s" % (key, "、".join(allowed)))
        elif sty.get(key) and sty[key] not in ("band", "panel", "half"):
            need = (st.get("photos") or {}).get("cover")
            if key == "transition" and not all(s.get("photo") for s in deck.get("sections", [])):
                warn.append("转场版式 %s 需要每章配图（sections[].photo），没配图的章节会退回原版式" % sty[key])
            if key in ("cover", "agenda") and not need:
                warn.append("%s版式 %s 需要封面照片（settings.photos.cover），没有时退回原版式" % ("封面" if key == "cover" else "目录", sty[key]))
    cl = deck.get("closing") or {}
    if cl.get("words") and cl["words"] != "感谢聆听":
        warn.append("结束页大字默认「感谢聆听」，现在写的是「%s」；不是用户要求的就删掉 closing" % cl["words"])
    if not st.get("reporter"):
        warn.append("没有 settings.reporter：结束页将列单位，不列汇报人；材料或用户给了汇报人就补上")
    if not st.get("logo"):
        warn.append("没有 settings.logo：整套不放 LOGO。汇报对象是有 LOGO 的单位时，先用 scripts/logo.py find 从源材料"
                    "（年报封面和前几页）找，找不到再联网找官方 LOGO")
    elif not os.path.isfile(st["logo"]):
        warn.append("settings.logo「%s」不存在，整套不会放 LOGO" % st["logo"])
    ph = st.get("photos") or {}
    import photos as _PH
    online = _PH.online_default(ph)
    wanted = [("封面", ph.get("cover"))] + [("转场「%s」" % s.get("name"), s.get("photo")) for s in deck.get("sections", [])]
    if not any(t for _, t in wanted):
        warn.append("封面和章节转场都没有配图：先用 scripts/photos.py 搜索 找相关的图（或开启联网取图）；确实没有相关的图可以不放")
    for where, t in wanted:
        if t and not _photo_ok(t, online):
            warn.append("%s的图片「%s」在自带图库里找不到相关的图，该页将不放图；换关键词、写文件名或开启联网取图" % (where, t))
    sigs, seps, creative, shapes, skels, types, titles, noicon, npoints = [], [], 0, [], [], set(), {}, 0, 0
    mains, halves = [], []   # 有大主块的页；所有两分刀的较小一段占比
    for i, sp in enumerate(pages, 1):
        e, w = check_page(sp, i, evidence_text, online)
        err += e
        warn += w
        if names and sp.get("section") not in names:
            warn.append("第 %d 个正文页的 section「%s」不在 sections 里" % (i, sp.get("section")))
        sigs.append(_signature(sp["partition"]) if sp.get("partition") else "creative")
        if sp.get("partition") and sp.get("separation") != "创意型" and not _bad_cuts(sp["partition"]):
            shapes.append((i, _shape(sp["partition"])))
            skels.append(_shape(sp["partition"], typed=False))
            lv = []
            try:
                _layout(sp["partition"], *BODY, 0, lv)
            except (KeyError, ValueError):
                lv = []
            cells = []  # 按实际排版（纯卡片页扣掉卡片间距）算格子尺寸，大主块面积以此为准
            try:
                _layout(sp["partition"], *BODY, CELL_GAP.get(sp.get("separation"), 0), cells)
            except (KeyError, ValueError):
                cells = []
            if any(r[2] * r[3] >= BODY[2] * BODY[3] * MAIN_BLOCK_AREA for lf, r in cells if not _is_photo(lf)):
                mains.append(i)
            _even_cuts(sp["partition"], halves)
            for lf, _ in lv:
                if lf.get("title") and not (lf.get("fill") == "重点" or lf.get("card") == "dark"):
                    titles.setdefault(lf["title"], []).append(i)
                for e in _elements(lf.get("element") or {}):
                    types.add(e.get("type"))
                    if e.get("type") == "points":
                        npoints += 1
                        noicon += 0 if _has_icon(e) else 1
        else:
            types.add("创意型")
        seps.append(sp.get("separation"))
        creative += 1 if sp.get("separation") == "创意型" else 0
    for i in range(1, len(sigs)):
        if sigs[i] == sigs[i - 1] and sigs[i] != "creative":
            warn.append("第 %d、%d 个正文页切分完全相同，建议换一种切法" % (i, i + 1))
    # 整稿版面单一度：比例微调不算换版面，按"切法＋每格组件"的结构指纹计
    for k in range(1, len(shapes)):
        (a, sa), (b, sb) = shapes[k - 1], shapes[k]
        if b == a + 1 and sa == sb:
            err.append("第 %d、%d 个正文页结构相同（%s），只改了比例；换切法或换组件" % (a, b, sa))
    n = len(shapes)
    nbody = len(sigs)  # 正文总页数，含创意页
    if n >= 6:
        top, cnt = Counter(s for _, s in shapes).most_common(1)[0]
        cap = max(2, int(n * 0.2))
        if cnt > cap:
            pages_of = [str(i) for i, s in shapes if s == top]
            err.append("整稿 %d 个正文页里有 %d 页是同一种结构 %s（第 %s 页），上限 %d 页；版面要逐页按材料关系决定，"
                       "不要用同一个模板批量生成" % (n, cnt, top, "、".join(pages_of), cap))
        one = sum(1 for s in skels if s == "v2(x,x)" or s == "h2(x,x)")
        if one > n * TWO_CELL_MAX:
            err.append("整稿 %d 个正文页里 %d 页只切一刀分两格，上限 %d%%；其余页要再切"
                       "（指标带三明治、英雄区＋卫星、主栏里再横切、竖排三四格等，见 references/split.md）"
                       % (n, one, TWO_CELL_MAX * 100))
        # 田字格：v2(h2,h2) 和 h2(v2,v2) 指纹不同，看上去却都是四个差不多大的格子，合并计数
        grid = [i for i, s in zip((i for i, _ in shapes), skels) if s in GRIDS]
        grid_cap = max(1, int(n * GRID_MAX))
        if len(grid) > grid_cap:
            err.append("整稿 %d 个正文页里 %d 页是田字格（先切两半、两边再各切两半，第 %s 页），上限 %d 页；"
                       "改成主栏＋侧栏、英雄区＋卫星、竖排三栏、指标带三明治等有主次的切法"
                       % (n, len(grid), "、".join(map(str, grid)), grid_cap))
        need_main = math.ceil(n * MAIN_BLOCK_MIN)
        if len(mains) < need_main:
            err.append("整稿 %d 个正文页里只有 %d 页有大主块（一个不再细分、占正文区 %d%% 以上的格子，图片格不算），至少 %d 页；"
                       "主证据整块放大，其余辅证据拆小格，页面才有主次" % (n, len(mains), MAIN_BLOCK_AREA * 100, need_main))
        if halves:
            even = sum(1 for r in halves if r >= 0.45)
            if even > len(halves) * EVEN_CUT_MAX:
                err.append("整稿两分的刀里有 %d/%d 刀接近对半（45:55 以内），上限一半；按分量切，多用 2:1、3:1、62:38 这类有主次的比例"
                           "（左右对照这类确需对半的可以保留）" % (even, len(halves)))
        # 套内版式变化：骨架（不看组件）也不能集中在少数几种，要有切得更细的页、第一刀也要有变化
        bare_top, bare_cnt = Counter(skels).most_common(1)[0]
        bare_cap = max(2, int(n * SKELETON_MAX))
        if bare_cnt > bare_cap:
            err.append("整稿 %d 个正文页里 %d 页是同一种切法骨架 %s（上限 %d 页）；换组件不算换版式，"
                       "换一种切法（先横后竖、先竖后横、等分三四格、英雄区＋卫星等）" % (n, bare_cnt, bare_top, bare_cap))
        if nbody >= 8:  # 按正文总页数（含创意页）触发，再统计符合条件的切分页
            many = sum(1 for s in skels if s.count("x") >= 4)
            if many < n * FOUR_CELL_MIN:
                err.append("整稿 %d 个正文页里只有 %d 页切到 4 格以上，至少 %d 页；一格不必对应一个中观点，"
                           "可以只放一组指标、一条关键论据或一个小图示" % (n, many, math.ceil(n * FOUR_CELL_MIN)))
            three = sum(1 for s in skels if re.match(r"[vh][3-9]", s))
            if three < 2:
                err.append("整稿 %d 个正文页（可切分 %d 页）里只有 %d 页第一刀切成 3 段以上（竖排三栏、上中下三段等），至少 2 页"
                           % (nbody, n, three))
            fine = sum(1 for s in skels if s.count("x") >= 5)
            if fine < 2:
                err.append("整稿 %d 个正文页（可切分 %d 页）里只有 %d 页切到 5 格以上，至少 2 页（如主栏＋侧栏三格、指标带＋三栏、英雄区＋三卫星）"
                           % (nbody, n, fine))
            firsts = {m.group(0) for m in (re.match(r"[vh]\d+", s) for s in skels) if m}
            if len(firsts) < 3:
                warn.append("整稿正文页的第一刀只有 %s 这几种；第一刀至少用 3 种（如左右两分、上下两分、竖排三四等分、"
                           "上中下三段），不规定具体切法" % "、".join(sorted(firsts)))
        need = min(7, 3 + n // 4)
        used = types - {"创意型", None}
        if len(used) < need:
            err.append("整稿只用了 %d 种组件（%s），%d 个正文页至少要用 %d 种；把部分要点换成关系图示、指标带、数字对或结论卡"
                       % (len(used), "、".join(sorted(used)), n, need))
    for t, ps in titles.items():
        if len(ps) >= 3:
            err.append("格子小标题「%s」在第 %s 页重复出现；小标题要写本格的具体判断" % (t, "、".join(map(str, ps[:8]))))
    if npoints >= 4 and noicon > npoints / 2:
        warn.append("%d 个要点列表里有 %d 个没有图标；要点默认给每条配图标（第三项写 IconPark 名或中文关键词）" % (npoints, noicon))
    for i in range(3, len(seps)):
        if len(set(seps[i - 3:i + 1])) == 1:
            warn.append("第 %d～%d 个正文页连续 4 页都是「%s」" % (i - 2, i + 1, seps[i]))
            break
    if creative > max(1, len(pages) // 5):
        warn.append("创意型页面 %d 页，超过整稿的 1/5" % creative)
    return err, warn


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("deck")
    ap.add_argument("--evidence")
    a = ap.parse_args(argv)
    try:
        deck = json.load(open(a.deck, encoding="utf-8"))
    except Exception as ex:
        print("deck.json 读取失败：%s" % ex)
        return 2
    ev = open(a.evidence, encoding="utf-8").read() if a.evidence else None
    err, warn = check_deck(deck, ev)
    for w in warn:
        print("提醒：" + w)
    for e in err:
        print("错误：" + e)
    print("通过" if not err else "未通过：%d 个错误" % len(err))
    return 1 if err else 0


if __name__ == "__main__":
    sys.exit(main())

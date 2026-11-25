"""v2 确认流程：故事线、配色、补充材料三项，**一个选择框一次问完**。
默认不读标准输入，待确认返回 10。

  python3 confirm.py ask --outline 故事线.md --checklist materials.md --industry 医疗   # 一次列出三道题
  python3 confirm.py answer --outline 故事线.md --checklist materials.md \
          --storyline 确认 --palette 科技蓝 --materials 不补                         # 一次记录三项答复
  （以下为单项命令，改大纲后重问或兼容旧流程时用）
  python3 confirm.py next                                   # 看下一步停在哪里
  python3 confirm.py storyline --choice confirm --outline 故事线.md
  python3 confirm.py palette --logo logo.png | --color C8102E | --industry 银行
  python3 confirm.py materials --choice 不需要 --checklist materials.md
  python3 confirm.py defaults --reason 你定 --outline 故事线.md [--materials 不补充]
  python3 confirm.py check                                  # 三项都确认才返回 0
"""
import argparse
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

import palette as P

EXIT_OK = 0
EXIT_STOP = 10
EXIT_USAGE = 2

STATE_VERSION = 2
DEFAULT_REASONS = ("你定", "用户不在")
GATES = ("storyline", "palette", "materials")
GATE_NAMES = {"storyline": "故事线", "palette": "配色", "materials": "询问是否补充材料"}

# 三次确认都用选择框问用户：有选择题工具（AskUserQuestion、ask_user_input 等）就调出选择框，
# 选项写在下面；只有当前环境没有选择题工具时，才把选项编号写成文字让用户回一个数字。
CHOICE_HINT = "用选择框问用户（有选择题工具就调出来，没有时把选项编号列出）："

ASK_HINT = ("一次问完：调出**一个**选择框，把还没确认的几项各作为一道题放进去（AskUserQuestion 一次可放 1–4 题），"
            "不要分几轮问；环境只支持一次一题时，改成一条文字消息列出全部题目和编号，让用户一次回复，如“1①2②3②”。")

STOP_TEXT = {
    "storyline": (
        "停在确认：故事线还没确认。先写完大纲、evidence.md 和 materials.md，再运行\n"
        "  confirm.py ask --outline 故事线.md --checklist materials.md --industry 行业（或 --logo / --color）\n"
        "把故事线、配色、补充材料三道题放进同一个选择框一次问完。"
    ),
    "palette": (
        "停在确认：配色还没确认。运行 confirm.py ask … 得到推荐色系和 2 个备选，和其他未确认项放进同一个选择框；\n"
        "每个选项写色系名和主色色值，不让用户自己输入色值。"
    ),
    "materials": (
        "停在确认：补充材料还没答复。运行 confirm.py ask … 把材料清单的缺口和选项（我来补充 / 不补 / 你定）\n"
        "与其他未确认项放进同一个选择框；用户选“我来补充”时等材料到了、更新底稿后再运行\n"
        "  confirm.py materials --choice 已补充 --checklist materials.md"
    ),
}


def default_state_path():
    return os.path.abspath(".niuppt-confirm.json")


def enable_auto_for_matrix(mock):
    """给 tests/matrix.py 用。mock 为真时打开自动模式，确认门不再等待。

    #11 合入后，在 run_matrix 里 use_mock 确定后加一行：
        enable_auto_for_matrix(use_mock)
    当前矩阵不调用本函数也能跑完，因为它不进入确认门。
    """
    if mock:
        os.environ["NIUPPT_AUTO"] = "1"
    return bool(mock)


def auto_enabled(argv=None, config_path=None, flag=False):
    """自动模式只认显式开关。正常做稿时这三项都应该是关的。

    argv 只在调用方显式传入时才检查，不读进程自己的 sys.argv，
    避免测试进程或其他命令行里的 --auto 被误当成确认开关。
    """
    if flag:
        return True
    env = os.environ.get("NIUPPT_AUTO", "").strip().lower()
    if env in {"1", "true", "yes", "on"}:
        return True
    if argv is not None and "--auto" in list(argv):
        return True
    path = config_path or os.environ.get("NIUPPT_AUTO_CONFIG") or ""
    if not path:
        path = next((p for p in ("niuppt.auto.json", "niuma.auto.json") if os.path.isfile(p)), None)   # 兼容旧名
    if path and os.path.isfile(path):
        try:
            with open(path, encoding="utf-8") as handle:
                data = json.load(handle)
        except (OSError, json.JSONDecodeError):
            return False
        return bool(isinstance(data, dict) and data.get("auto") is True)
    return False




def load_state(path):
    if not os.path.isfile(path):
        return {"version": STATE_VERSION}
    with open(path, encoding="utf-8") as handle:
        data = json.load(handle)
    return data if isinstance(data, dict) else {"version": STATE_VERSION}


def save_state(path, state):
    parent = os.path.dirname(path)
    if parent:
        os.makedirs(parent, exist_ok=True)
    state["version"] = STATE_VERSION
    with open(path, "w", encoding="utf-8") as handle:
        json.dump(state, handle, ensure_ascii=False, indent=2)
        handle.write("\n")


def missing_gate(state):
    for gate in GATES:
        item = state.get(gate)
        if not isinstance(item, dict) or not item.get("source"):
            return gate
    return None


def ready(state):
    return missing_gate(state) is None


def _story_record(choice, source, reason, outline, industry=""):
    return {"choice": choice, "source": source, "reason": reason or "",
            "outline": os.path.abspath(outline) if outline else "", "industry": industry or "",
            "depth": "argument", "alternatives": 1}


def _materials_record(choice, source, reason, checklist):
    return {"choice": choice or "adopt", "source": source, "reason": reason or "",
            "checklist": os.path.abspath(checklist) if checklist else "", "invented": False}


_EMPTY_MID = {"", "—", "-", "－", "无", "无。", "待补充", "待定", "（无）", "（待补充）"}
_MID_LABELS = ("页内中标题或中观点", "页内中观点", "页内中标题")
_ARG_LABELS = ("中观点论据", "论据")


def _substantive(text):
    value = (text or "").strip().strip("|").strip()
    if value in _EMPTY_MID:
        return False
    compact = "".join(value.split())
    return len(compact) >= 2 and compact not in _EMPTY_MID


def _coverish(section):
    value = (section or "").strip()
    return value in {"封面", "目录", "—", "-", "－"} or value.startswith(("封面", "目录", "致谢", "转场"))


def _field_value(line, label):
    if label not in line:
        return None
    rest = line[line.find(label) + len(label):].lstrip("）)").lstrip()
    if rest.startswith(("：", ":")):
        return rest[1:].strip()
    return None


def _table_pages(text):
    """文件里可能有多张表：逐张读，取第一张表头含“页标题观点”和“页内中观点”的作大纲。"""
    tables, rows = [], []
    for line in text.splitlines() + [""]:
        raw = line.strip()
        if not raw.startswith("|"):
            if rows: tables.append(rows); rows = []
            continue
        cells = [cell.strip() for cell in raw.strip("|").split("|")]
        if not cells or all(not cell or set(cell) <= set("-: ") for cell in cells):
            continue
        rows.append(cells)
    for rows in tables:
        pages = _pages_from_rows(rows)
        if pages:
            return pages
    return []


def _pages_from_rows(rows):
    header = rows[0]

    def column(names):
        for index, cell in enumerate(header):
            if any(name in cell for name in names):
                return index
        return None

    i_section = column(("部分",))
    i_title = column(("页标题观点", "每页观点", "页标题"))
    i_mid = column(_MID_LABELS)
    i_arg = column(_ARG_LABELS)
    if i_mid is None or i_title is None:
        return []
    pages = []
    for cells in rows[1:]:
        def take(index):
            if index is None or index >= len(cells):
                return ""
            return cells[index]

        section = take(i_section)
        title = take(i_title)
        mid = take(i_mid)
        if section in {"部分"} or "页标题" in title:
            continue
        pages.append({"section": section, "title": title, "mid": mid, "arg": take(i_arg)})
    return pages


def _indent_pages(text):
    pages = []
    current = None
    section = ""
    for line in text.splitlines():
        raw = line.strip().lstrip("-*").strip()
        if not raw:
            continue
        found = _field_value(raw, "部分")
        if found is not None and "页标题" not in raw and "页内中" not in raw:
            section = found
            continue
        title = _field_value(raw, "页标题观点")
        if title is None:
            title = _field_value(raw, "每页观点")
        if title is not None:
            if current:
                pages.append(current)
            current = {"section": section, "title": title, "mid": "", "arg": ""}
            continue
        if current is None:
            continue
        arg = None
        for label in _ARG_LABELS:
            arg = _field_value(raw, label)
            if arg is not None:
                current["arg"] = (current["arg"] + "；" + arg).strip("；")
                break
        if arg is not None:
            continue
        for label in _MID_LABELS:
            mid = _field_value(raw, label)
            if mid is not None:
                current["mid"] = mid
                break
    if current:
        pages.append(current)
    return pages


def outline_pages(text):
    """从四列表或三层缩进里取出页面。同一份里两种都写了，就都收进来。"""
    return _table_pages(text or "") + _indent_pages(text or "")


def validate_story_outline(text):
    """确认故事线要写到中观点的论据：部分 → 页标题观点 → 页内中观点 → 中观点论据。

    论据写要点即可（每个中观点 2 条左右，带关键数字）；口径、比较基准和出处留到论据底稿。
    封面、目录可以没有页内中观点和论据。
    """
    if not text or not str(text).strip():
        raise ValueError("大纲是空的。内容页要写到页内中观点，才能确认并进入补充材料。")
    pages = outline_pages(text)
    content = [page for page in pages if not _coverish(page.get("section"))]
    if not content:
        raise ValueError("大纲没有内容页。确认需要写到：部分 → 每页观点 → 页内中观点。")
    missing = [
        page for page in content
        if not _substantive(page.get("title")) or not _substantive(page.get("mid"))
    ]
    if missing:
        raise ValueError(
            "内容页缺页内中观点或页标题。封面和目录可以不写中观点。"
        )
    no_arg = [page.get("title", "")[:16] for page in content if not _substantive(page.get("arg"))]
    if no_arg:
        raise ValueError(
            "故事线要写到中观点的论据（“中观点论据”一列，每个中观点 2 条左右论据要点，带关键数字）。"
            "缺论据的页：" + "、".join(no_arg[:6]) + ("等" if len(no_arg) > 6 else "")
        )
    return content




# ---------------- 配色：直接指定 ----------------
def assign_palette(logo=None, color=None, industry="", seed=""):
    """按 LOGO 图片、LOGO 主色或行业指定配色。返回 (色系名或 "#色值", 依据)。
    有 LOGO 直接用 LOGO 原色推导整套配色（过灰过深时退回最接近的预设）；
    没有 LOGO 按行业在候选色里按 seed（大纲标题）稳定地挑一个，同一份稿子重跑结果不变。"""
    if logo and os.path.isfile(logo):
        import brand
        colors = brand.logo_colors(logo)
        chroma = [c for c, _, k in colors if k == "chroma"]
        details = brand.logo_color_details(logo) if chroma else []
        shape = [d["color"] for d in details if d.get("primary_origin") == "图形/色块"]
        # 深海军蓝、墨绿这类深色品牌色亮度低，取色时归在 dark 一类而不是 chroma；只要有足够饱和度也照用
        deep = [c for c, _, k in colors if k in ("dark", "neutral") and P.usable(c.lstrip("#"))]
        pick = (shape or chroma or deep or [None])[0]
        if pick:
            pick = pick.lstrip("#").upper()
            if P.usable(pick):
                return "#" + pick, "取自 LOGO %s" % ("图形" if shape else "主色")
            return P.nearest_preset(pick), "LOGO 色 #%s 过灰或过深，取最接近的预设" % pick
        return "炭灰", "LOGO 为黑白或低饱和"
    if color:
        c = color.lstrip("#").upper()
        if P.usable(c):
            return "#" + c, "取自 LOGO 主色"
        return P.nearest_preset(c), "指定色 #%s 过灰或过深，取最接近的预设" % c
    if industry:
        cands = P.by_industry(industry, seed)
        if cands:
            return cands[0], "行业：%s（候选 %s）" % (industry, "、".join(cands))
    return "藏蓝", "默认"


def _seed(outline):
    """取大纲第一行标题作随机种子；读不到就用空串。"""
    try:
        with open(outline, encoding="utf-8") as f:
            for line in f:
                if line.strip():
                    return line.strip("# \n")
    except (OSError, TypeError):
        pass
    return ""


def _print_stop(gate):
    print(STOP_TEXT[gate])


def _print_ready(state):
    pal = state.get("palette") or {}
    print("三项确认已完成，可以编写 deck.json。配色：%s（%s）。" % (pal.get("preset", "藏蓝"), pal.get("basis", "默认")))


def delivery_lines(state):
    lines = []
    for gate, title in (("storyline", "故事线"), ("materials", "补充材料")):
        item = state.get(gate) or {}
        if item.get("source") == "default":
            lines.append("%s：模型代选（%s）" % (title, item.get("reason")))
    pal = state.get("palette") or {}
    if pal:
        lines.append("配色：%s，依据 %s" % (pal.get("preset"), pal.get("basis")))
    return lines


# ---------------- 命令 ----------------
def cmd_next(args):
    state = load_state(args.state)
    gate = missing_gate(state)
    if gate is None:
        _print_ready(state)
        return EXIT_OK
    _print_stop(gate)
    return EXIT_STOP


def cmd_storyline(args):
    state = load_state(args.state)
    key = str(args.choice or "").strip().lower()
    if key not in ("confirm", "确认", "adopt", "采用"):
        _print_stop("storyline")
        return EXIT_STOP
    if not args.outline or not os.path.isfile(args.outline):
        raise ValueError("确认故事线需要大纲文件（--outline）。内容页要写到页内中观点。")
    with open(args.outline, encoding="utf-8") as handle:
        validate_story_outline(handle.read())
    source = "default" if args.reason else "user"
    state["storyline"] = _story_record("adopt" if args.reason else "confirm", source, args.reason, args.outline, args.industry)
    state["web_search"] = False
    save_state(args.state, state)
    print("已记录一条故事线。还没确认的项用 confirm.py ask 列出，放进同一个选择框一次问完。")
    return EXIT_OK


def cmd_palette(args):
    """不带 --choice：给出推荐色系和 2 个备选，停下来用选择框问用户；带 --choice：记录用户选的色系。"""
    state = load_state(args.state)
    choice = (args.choice or args.preset or "").strip()
    if not choice:
        name, basis = assign_palette(args.logo, args.color, args.industry, _seed(getattr(args, "outline", None)))
        alts = P.alternatives(name, 2, args.industry, _seed(getattr(args, "outline", None)))
        state["palette_offer"] = {"recommend": name, "basis": basis, "alternatives": alts}
        save_state(args.state, state)
        print("推荐配色：%s（主色 #%s），依据：%s。" % (P.label(name), P.main_of(name), basis))
        print(CHOICE_HINT)
        print("  ① %s（推荐）" % P.option(name))
        for i, a in enumerate(alts, 2):
            print("  %s %s（主色 #%s，%s）" % ("②③④"[i - 2], a, P.main_of(a), P.COLORS[a]["use"].split("、")[0]))
        print("  %s 你定" % "②③④⑤"[len(alts)])
        print("用户选好后：confirm.py palette --choice 色系名（选“你定”就传 %s）。" % name)
        return EXIT_STOP
    offer = state.get("palette_offer") or {}
    if choice in ("你定", "推荐", "recommend"):
        if not offer:
            name, basis = assign_palette(args.logo, args.color, args.industry, _seed(getattr(args, "outline", None)))
        else:
            name, basis = offer["recommend"], offer["basis"]
        source = "user"
    else:
        if choice.startswith("LOGO原色"):
            choice = choice.replace("LOGO原色", "").strip()
        if choice.startswith("#") and P.is_hex(choice):
            choice = "#" + choice.lstrip("#").upper()
        if choice not in P.COLORS and not (choice.startswith("#") and P.is_hex(choice)):
            raise ValueError("没有这个色系：%s。可选：%s，或 LOGO 原色值（如 #1E50A2）" % (choice, "、".join(P.COLORS)))
        name = choice
        basis = offer.get("basis", "") if offer.get("recommend") == name else "用户选择"
        source = "user"
    state["palette"] = {"preset": name, "basis": basis, "corner": args.corner, "source": source}
    save_state(args.state, state)
    p = P.preset(name)
    print("配色：%s（主色 #%s，%s 组合），依据：%s。把“%s”写入 deck.json 的 settings.preset。" % (P.label(name), p["main"], p["scheme"], basis, name))
    return EXIT_OK


def cmd_materials(args):
    state = load_state(args.state)
    if "storyline" not in state:
        raise ValueError("尚未确认故事线，不能先确认补充材料。")
    if not args.choice:
        _print_stop("materials")
        return EXIT_STOP
    state["materials"] = _materials_record(args.choice.strip(), "user", "", args.checklist)
    save_state(args.state, state)
    print("已记录补充材料的答复。更新论据底稿并通过 check_evidence.py 后编写 deck.json。")
    return EXIT_OK


def cmd_defaults(args):
    if args.reason not in DEFAULT_REASONS:
        raise ValueError("defaults 只用于用户说「你定」或不在。--reason 只能是「你定」或「用户不在」。")
    state = load_state(args.state)
    if "storyline" not in state:
        if not args.outline or not os.path.isfile(args.outline):
            raise ValueError("代选故事线需要大纲文件：defaults --outline 故事线.md。")
        with open(args.outline, encoding="utf-8") as handle:
            validate_story_outline(handle.read())
        state["storyline"] = _story_record("adopt", "default", args.reason, args.outline, args.industry)
    if "palette" not in state:
        name, basis = assign_palette(args.logo, None, args.industry, _seed(getattr(args, "outline", None)))
        state["palette"] = {"preset": name, "basis": basis, "corner": "square", "source": "assigned"}
    if "materials" not in state:
        if args.materials:
            state["materials"] = _materials_record(args.materials, "default", "%s；用户事先答复：%s" % (args.reason, args.materials), None)
        elif args.reason == "你定":
            state["materials"] = _materials_record("adopt", "default", args.reason, None)
    state["web_search"] = False
    save_state(args.state, state)
    if not ready(state):
        _print_stop(missing_gate(state))
        return EXIT_STOP
    print("已按「%s」补齐确认，交付时必须说明。" % args.reason)
    _print_ready(state)
    return EXIT_OK


def cmd_auto(args):
    if not auto_enabled(flag=args.auto, config_path=args.config):
        print("未打开自动模式。正常使用请停下来等用户，不要跳过确认。")
        return EXIT_STOP
    state = load_state(args.state)
    state.setdefault("storyline", _story_record("adopt", "auto", "自动测试", None, args.industry))
    state.setdefault("materials", _materials_record("adopt", "auto", "自动测试", None))
    if "palette" not in state:
        name, basis = assign_palette(args.logo, None, args.industry, _seed(getattr(args, "outline", None)))
        state["palette"] = {"preset": name, "basis": basis, "corner": "square", "source": "assigned"}
    save_state(args.state, state)
    print("自动模式：确认已通过，未读取输入。")
    _print_ready(state)
    return EXIT_OK


def cmd_check(args):
    state = load_state(args.state)
    if ready(state):
        _print_ready(state)
        for line in delivery_lines(state):
            print("交付说明：" + line)
        return EXIT_OK
    _print_stop(missing_gate(state))
    return EXIT_STOP


# ---------------- 一次问完 ----------------
STORY_OK = ("confirm", "确认", "adopt", "采用", "1", "①")
STORY_EDIT = ("修改", "我要修改", "edit", "2", "②")
MAT_NO = ("不补", "不补充", "不需要", "按现有材料成稿", "不补，按现有材料成稿", "no")
MAT_YES = ("补充", "我来补充", "yes")


def _gap_count(checklist):
    """数材料清单里的缺口条数（以 - / * / 数字. 开头的行）；没有清单返回 None。"""
    if not checklist or not os.path.isfile(checklist):
        return None
    import re
    n = 0
    for line in open(checklist, encoding="utf-8"):
        if re.match(r"\s*([-*]|\d+[.、])\s+\S", line) and not re.search(r"无(待补|缺口)|没有缺口", line):
            n += 1
    return n


def cmd_ask(args):
    """列出还没确认的几项，提示用一个选择框一次问完。三项都已确认时返回 0。"""
    state = load_state(args.state)
    todo = [g for g in GATES if not (isinstance(state.get(g), dict) and state[g].get("source"))]
    if not todo:
        _print_ready(state)
        return EXIT_OK
    if "storyline" in todo:
        if not args.outline or not os.path.isfile(args.outline):
            raise ValueError("一次问完需要先写好大纲文件（--outline 故事线.md），内容页写到中观点论据。")
        with open(args.outline, encoding="utf-8") as handle:
            validate_story_outline(handle.read())
    if "materials" in todo and not (args.checklist and os.path.isfile(args.checklist)):
        raise ValueError("一次问完前先写好 evidence.md 并运行 check_evidence.py --checklist materials.md，"
                         "把缺口清单和故事线一起给用户看（--checklist materials.md）。")
    print(ASK_HINT)
    print("同一条消息里先给出：大纲全文" + ("、材料缺口清单" if "materials" in todo else "") + "；然后调出选择框：")
    q = 1
    if "storyline" in todo:
        print("  题%d 故事线：① 确认这条故事线（推荐） ② 我要修改（改动写在“其他”里） ③ 你定" % q)
        q += 1
    if "palette" in todo:
        name, basis = assign_palette(args.logo, args.color, args.industry, _seed(args.outline))
        alts = P.alternatives(name, 2, args.industry, _seed(args.outline))
        state["palette_offer"] = {"recommend": name, "basis": basis, "alternatives": alts}
        save_state(args.state, state)
        opts = ["① %s（推荐，%s）" % (P.option(name), basis)]
        opts += ["%s %s" % ("②③"[i], P.option(a)) for i, a in enumerate(alts)]
        opts.append("%s 你定" % "②③④"[len(alts)])
        print("  题%d 配色：%s" % (q, " ".join(opts)))
        q += 1
    if "materials" in todo:
        n = _gap_count(args.checklist)
        gap = "清单列了 %d 处缺口" % n if n else "清单没有缺口"
        rec = "② 不补，按现有材料成稿（推荐）" if not n else "② 不补，按现有材料成稿"
        print("  题%d 补充材料（%s）：① 我来补充（文字或图片） %s ③ 你定" % (q, gap, rec))
    print("用户选完后一次记录：confirm.py answer --outline 故事线.md --checklist materials.md "
          "--storyline 确认|修改|你定 --palette 色系名|你定 --materials 补充|不补|你定")
    return EXIT_STOP


def cmd_answer(args):
    """一次记录选择框里的答复。故事线选“修改”时：配色照记，故事线和补充材料改完后再用 ask 重问（只问这两题）。"""
    state = load_state(args.state)
    msgs = []
    if args.palette:
        pa = argparse.Namespace(**vars(args))
        pa.choice, pa.preset = args.palette, None
        cmd_palette(pa)
        state = load_state(args.state)
    story = str(args.storyline or "").strip().lower()
    if story:
        if story in STORY_OK or story == "你定":
            if not args.outline or not os.path.isfile(args.outline):
                raise ValueError("记录故事线需要大纲文件（--outline）。")
            with open(args.outline, encoding="utf-8") as handle:
                validate_story_outline(handle.read())
            reason = "你定" if story == "你定" else None
            state["storyline"] = _story_record("adopt" if reason else "confirm", "user", reason, args.outline, args.industry)
            state["web_search"] = False
        elif story in STORY_EDIT:
            msgs.append("故事线要修改：按用户意见改大纲，同步改 evidence.md、重跑 check_evidence.py 生成新清单，"
                        "再运行 confirm.py ask，只重问故事线和补充材料（配色已记录，不再问）。")
            state.pop("storyline", None)  # 旧大纲作废，改完要重新确认；配色保留
            state.pop("materials", None)
            save_state(args.state, state)
            for m in msgs:
                print(m)
            return EXIT_STOP
        else:
            raise ValueError("--storyline 只能是 确认 / 修改 / 你定")
    mat = str(args.materials or "").strip()
    if mat:
        if "storyline" not in state:
            raise ValueError("故事线还没确认，补充材料的答复等故事线定了再记录。")
        if mat.lower() in MAT_YES or mat.startswith("我来补"):
            msgs.append("用户要补充材料：等材料（文字或图片）到了，纳入 evidence.md 并通过 check_evidence.py，"
                        "再运行 confirm.py materials --choice 已补充 --checklist materials.md。")
        elif mat in MAT_NO or mat.startswith("不补") or mat == "你定":
            state["materials"] = _materials_record(mat, "user", "你定" if mat == "你定" else "", args.checklist)
        else:  # 用户在“其他”里直接写了补充内容
            state["materials"] = _materials_record(mat, "user", "", args.checklist)
    save_state(args.state, state)
    for m in msgs:
        print(m)
    if ready(state):
        _print_ready(state)
        return EXIT_OK
    if not msgs:
        _print_stop(missing_gate(state))
    return EXIT_STOP


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    common = argparse.ArgumentParser(add_help=False)
    common.add_argument("--state", default=default_state_path())
    common.add_argument("--industry", default="")
    common.add_argument("--logo", default="")
    common.add_argument("--auto", action="store_true")
    common.add_argument("--config")
    sub = ap.add_subparsers(dest="cmd", required=True)
    sub.add_parser("next", parents=[common])
    sub.add_parser("check", parents=[common])
    sub.add_parser("auto", parents=[common])
    k = sub.add_parser("ask", parents=[common], help="列出未确认项，用一个选择框一次问完")
    k.add_argument("--outline"); k.add_argument("--checklist"); k.add_argument("--color")
    w = sub.add_parser("answer", parents=[common], help="一次记录选择框里的答复")
    w.add_argument("--outline"); w.add_argument("--checklist"); w.add_argument("--color")
    w.add_argument("--storyline"); w.add_argument("--palette"); w.add_argument("--materials")
    w.add_argument("--corner", default="square", choices=["square", "round"])
    s = sub.add_parser("storyline", parents=[common])
    s.add_argument("--choice"); s.add_argument("--outline"); s.add_argument("--reason", choices=DEFAULT_REASONS)
    p = sub.add_parser("palette", parents=[common])
    p.add_argument("--color", help="LOGO 主色，六位色值")
    p.add_argument("--choice", help="用户在选择框里选的色系名，或“你定”")
    p.add_argument("--preset", help="同 --choice（兼容旧写法）")
    p.add_argument("--corner", default="square", choices=["square", "round"])
    p.add_argument("--outline", help="大纲文件，取标题作行业挑色的种子，与一次确认的结果一致")
    m = sub.add_parser("materials", parents=[common])
    m.add_argument("--choice"); m.add_argument("--checklist")
    d = sub.add_parser("defaults", parents=[common])
    d.add_argument("--reason", required=True); d.add_argument("--outline"); d.add_argument("--materials")
    a = ap.parse_args(argv)
    try:
        return {"next": cmd_next, "check": cmd_check, "auto": cmd_auto, "storyline": cmd_storyline,
                "palette": cmd_palette, "materials": cmd_materials, "defaults": cmd_defaults,
                "ask": cmd_ask, "answer": cmd_answer}[a.cmd](a)
    except ValueError as ex:
        print(ex)
        return EXIT_USAGE


if __name__ == "__main__":
    sys.exit(main())

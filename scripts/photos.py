# -*- coding: utf-8 -*-
"""图片：自带 CC0 图库（50 张，来自 Hugging Face 数据集 KoalaAI/StockImages-CC0），可选联网从多个开源图库取图。
按单张图的中文匹配词挑图（也可直接写文件名），铺满指定格子（等比裁切），正文页内的图片先降饱和、向主色调色；封面和转场的全幅图叠主色渐变蒙版。
联网取到的每张图都记入输出目录的 图片来源.json，供用户核查授权。网络不可用时退回自带图库；没有相关的图就不放。"""
import json
import os
import tempfile
from PIL import Image
from pptx.util import Emu
from pptx.oxml.ns import qn
from pptx.enum.shapes import MSO_SHAPE
from lxml import etree

HERE = os.path.dirname(os.path.abspath(__file__))
DIR = os.path.join(os.path.dirname(HERE), "assets", "photos")
PX = 9525
_LIB = None

# 中文主题 → 类别
TOPIC = {"城市": "城市楼宇", "楼宇": "城市楼宇", "总部": "城市楼宇", "区域": "城市楼宇", "经营": "城市楼宇",
         "物流": "物流交通", "运输": "物流交通", "冷链": "物流交通", "配送": "物流交通", "港口": "物流交通", "交通": "物流交通",
         "办公": "办公会议", "会议": "办公会议", "团队": "办公会议", "管理": "办公会议", "客户": "办公会议",
         "工业": "工业科技", "仓储": "工业科技", "库房": "工业科技", "能源": "工业科技", "科技": "工业科技", "制造": "工业科技",
         "自然": "自然意境", "展望": "自然意境", "愿景": "自然意境",
         "生鲜": "生鲜食品", "水产": "生鲜食品", "冻品": "生鲜食品", "鲜果": "生鲜食品", "鲜蔬": "生鲜食品",
         "乳品": "生鲜食品", "食品": "生鲜食品", "业务结构": "生鲜食品", "货运": "物流交通", "集装箱": "物流交通"}


def lib():
    global _LIB
    if _LIB is None:
        with open(os.path.join(DIR, "index.json"), encoding="utf-8") as f:
            _LIB = json.load(f)["images"]
    return _LIB


def _words(topic):
    import re
    return [w for w in re.split(r"[、，,；;/\s]+", topic or "") if w]


def _local_file(topic):
    """topic 是图库文件名或本机图片路径时，直接返回该文件。"""
    if not topic:
        return None
    if os.path.isfile(topic):
        return os.path.abspath(topic)
    if any(im["file"] == topic for im in lib()):
        return os.path.join(DIR, topic)
    return None


def candidates(topic, used=None, include_avoid=False):
    """按中文关键词给图库打分，返回 [(分数, 图片记录)]，分数高的在前。
    关键词命中 +3，画面描述命中 +1，大类对应 +1；不看大类顺序，只看单张图的画面。"""
    words = _words(topic)
    out = []
    for im in lib():
        if used is not None and im["file"] in used:
            continue
        if im.get("avoid") and not include_avoid:
            continue
        kws, desc = im.get("keywords", []), im.get("desc", "")
        score = 0
        for w in words:
            if w in kws:
                score += 3
            elif any(w in k or k in w for k in kws if len(k) > 1):
                score += 2
            if w in desc:
                score += 1
            if TOPIC.get(w) == im["category"]:
                score += 1
        if score:
            out.append((score, im))
    out.sort(key=lambda t: -t[0])
    return out


def pick(topic=None, tone=None, used=None, n=0):
    """按主题挑一张相关的图：topic 可以是图库文件名、本机图片路径，或中文关键词（多个用顿号隔开）。
    关键词至少命中一张图的匹配词才算相关；没有相关的图返回 None，由调用方决定不放图，不再随便拿一张顶上。
    used 为集合，挑中后自动加入，保证同一套 PPT 不重复。"""
    f = _local_file(topic)
    if f:
        if used is not None:
            used.add(os.path.basename(f))
        return f
    cands = [(sc, im) for sc, im in candidates(topic, used) if sc >= 3 and (not tone or im["tone"] == tone)]
    if not cands:
        return None
    im = cands[min(n, len(cands) - 1)][1]
    if used is not None:
        used.add(im["file"])
    return os.path.join(DIR, im["file"])


# 全图蒙版版式（首页、目录、转场整幅铺图、上面压蒙版写字）的背景模糊：按 1280 宽画布折算的像素半径
BG_BLUR = 7


def blur(path, radius=BG_BLUR):
    """背景照片做高斯模糊，半径按照片实际宽度折算到 1280 宽画布，不同分辨率的图模糊程度一致。返回新文件路径。"""
    from PIL import ImageFilter
    im = Image.open(path).convert("RGB")
    r = radius * im.width / 1280
    out_dir = os.path.join(tempfile.gettempdir(), "niuppt_blur")
    os.makedirs(out_dir, exist_ok=True)
    out = os.path.join(out_dir, "%s_b%d.jpg" % (os.path.splitext(os.path.basename(path))[0], radius))
    im.filter(ImageFilter.GaussianBlur(r)).save(out, quality=90)
    return out


def place(sl, path, x, y, w, h, tone_main=None, anchor="center", sat=0.25, duo=0.55, blurred=False):
    """等比铺满 (x,y,w,h)，超出部分用 PPT 原生裁切（在 PPT 里仍可调整裁切）。
    正文页内的图片传 tone_main=主色，先降饱和并向主色调色；封面、转场的全幅图靠蒙版统一色调，可不传。"""
    if tone_main:
        path = tone(path, tone_main, sat=sat, duo=duo)
    if blurred:
        path = blur(path)
    iw, ih = Image.open(path).size
    pic = sl.shapes.add_picture(path, Emu(int(x * PX)), Emu(int(y * PX)), Emu(int(w * PX)), Emu(int(h * PX)))
    r_box, r_img = w / h, iw / ih
    if r_img > r_box:
        cut = 1 - r_box / r_img
        k = {"left": 0, "right": 1}.get(anchor, 0.5)
        pic.crop_left = cut * k; pic.crop_right = cut * (1 - k)
    else:
        cut = 1 - r_img / r_box
        k = {"top": 0, "bottom": 1}.get(anchor, 0.5)
        pic.crop_top = cut * k; pic.crop_bottom = cut * (1 - k)
    return pic


def place_faded(sl, path, x, y, w, h, bg, stops, side="left", tone_main=None, sat=0.88, duo=0.10):
    """照片的渐隐直接做进图片里：从 side 一侧由页底色 bg 渐变到原图。
    stops = [(位置0~1, 底色不透明度), ...]，位置从 side 那一侧量起。
    图片边缘就是纯底色，PPT 和各种渲染器里都看不到照片边，也不依赖多档渐变蒙版。"""
    if tone_main:
        path = tone(path, tone_main, sat=sat, duo=duo)
    im = Image.open(path).convert("RGB")
    iw, ih = im.size
    r_box = w / h
    if iw / ih > r_box:
        nw = int(ih * r_box); im = im.crop(((iw - nw) // 2, 0, (iw - nw) // 2 + nw, ih))
    else:
        nh = int(iw / r_box); im = im.crop((0, (ih - nh) // 2, iw, (ih - nh) // 2 + nh))
    W_ = min(2400, im.size[0])
    im = im.resize((W_, max(1, int(W_ / r_box))), Image.LANCZOS)
    iw, ih = im.size
    horizontal = side in ("left", "right")
    n = iw if horizontal else ih
    ramp = []
    pts = sorted(stops)
    for i in range(n):
        t = i / max(n - 1, 1)
        if side in ("right", "bottom"):
            t = 1 - t
        a = pts[-1][1] if t >= pts[-1][0] else pts[0][1]
        for (p0, a0), (p1, a1) in zip(pts, pts[1:]):
            if p0 <= t <= p1:
                k = (t - p0) / max(p1 - p0, 1e-9)
                k = k * k * (3 - 2 * k)  # 平滑插值，档与档之间没有折痕
                a = a0 + (a1 - a0) * k
                break
        ramp.append(int(round(max(0.0, min(1.0, a)) * 255)))
    line = Image.new("L", (n, 1)); line.putdata(ramp)
    mask = line.resize((iw, ih)) if horizontal else line.transpose(Image.ROTATE_90).transpose(Image.FLIP_TOP_BOTTOM).resize((iw, ih))
    solid = Image.new("RGB", (iw, ih), tuple(int(bg[i:i + 2], 16) for i in (0, 2, 4)))
    out = Image.composite(solid, im, mask)
    out_dir = os.path.join(tempfile.gettempdir(), "niuppt_toned")
    os.makedirs(out_dir, exist_ok=True)
    fn = os.path.join(out_dir, "%s_fade_%s_%s_%d.jpg" % (os.path.splitext(os.path.basename(path))[0], side, bg, int(w)))
    out.save(fn, quality=93)
    return sl.shapes.add_picture(fn, Emu(int(x * PX)), Emu(int(y * PX)), Emu(int(w * PX)), Emu(int(h * PX)))


def overlay(sl, x, y, w, h, color, a0=0.92, a1=0.35, ang=0):
    """主色渐变蒙版：a0→a1 为不透明度，ang 为角度（0 = 从左到右）。"""
    s = sl.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(int(x * PX)), Emu(int(y * PX)), Emu(int(w * PX)), Emu(int(h * PX)))
    s.fill.gradient(); s.fill.gradient_angle = ang
    st = s.fill.gradient_stops
    from pptx.dml.color import RGBColor
    for stop, a, pos in ((st[0], a0, 0.0), (st[1], a1, 1.0)):
        stop.color.rgb = RGBColor.from_string(color); stop.position = pos
        clr = stop._gs.find(qn("a:srgbClr"))
        al = etree.SubElement(clr, qn("a:alpha")); al.set("val", str(int(a * 100000)))
    s.line.fill.background()
    stl = s._element.find(qn("p:style"))
    if stl is not None:
        s._element.remove(stl)
    return s


def tone(path, main="0E5A78", sat=0.25, duo=0.55, out_dir=None):
    """页内图片调色：降饱和 + 向主色做双色调混合，避免与主色冲突。
    sat：保留的原始饱和度（0=黑白）；duo：主色双色调占比。返回处理后的文件路径。"""
    from PIL import ImageEnhance, ImageOps
    im = Image.open(path).convert("RGB")
    low = ImageEnhance.Color(im).enhance(sat)
    low = ImageEnhance.Contrast(low).enhance(0.92)
    # 冷色主色直接向主色调色；暖色、绿色主色调色后照片会发粉或发绿，改向"主色混灰"的中性色调色
    import colorsys as _cs
    _h, _l, _s = _cs.rgb_to_hls(*(int(main[i:i + 2], 16) / 255 for i in (0, 2, 4)))
    if not (170 / 360 <= _h < 330 / 360):
        mm = [int(main[i:i + 2], 16) for i in (0, 2, 4)]
        main = "".join("%02X" % round(c + (110 - c) * 0.7) for c in mm)
    r, g, b = (int(main[i:i + 2], 16) for i in (0, 2, 4))
    dark = tuple(int(c * 0.55) for c in (r, g, b))
    light = tuple(int(c + (255 - c) * 0.92) for c in (r, g, b))
    duotone = ImageOps.colorize(ImageOps.grayscale(im), black=dark, white=light)
    res = Image.blend(low, duotone, duo)
    out_dir = out_dir or os.path.join(tempfile.gettempdir(), "niuppt_toned")
    os.makedirs(out_dir, exist_ok=True)
    out = os.path.join(out_dir, f"{os.path.splitext(os.path.basename(path))[0]}_{main}.jpg")
    res.save(out, quality=88)
    return out


# ================= 联网取图（可选） =================
# 中文主题 → 英文检索词。主题里本身有英文时直接用英文检索。
EN = {"城市": "city skyline", "楼宇": "office building", "总部": "corporate building", "物流": "logistics truck",
      "运输": "freight transport", "冷链": "refrigerated truck", "配送": "delivery van", "港口": "container port",
      "办公": "office desk", "会议": "business meeting room", "团队": "teamwork office", "工业": "factory production line",
      "仓储": "warehouse shelves", "库房": "warehouse", "能源": "solar power plant", "科技": "technology data center",
      "制造": "manufacturing", "自然": "mountain lake", "展望": "sunrise horizon", "生鲜": "fresh vegetables market",
      "水产": "seafood market", "乳品": "milk dairy", "农业": "farmland", "金融": "financial district",
      "银行": "bank building", "医疗": "hospital", "教育": "classroom", "零售": "retail store",
      "图书馆": "library bookshelves", "图书": "books bookshelf", "阅读": "reading books", "书": "books",
      "社区": "community neighborhood", "培训": "training workshop", "学校": "school campus", "校园": "university campus",
      "政务": "government service hall", "窗口": "service counter", "客服": "customer service", "柜台": "service counter",
      "文件": "documents paperwork", "合同": "contract signing", "数据": "data analytics dashboard", "电脑": "laptop work",
      "包装": "packaging boxes", "周转箱": "plastic crates", "纸箱": "cardboard boxes", "快递": "parcel delivery",
      "工地": "construction site", "建筑": "construction", "医院": "hospital corridor", "药品": "pharmacy",
      "农村": "rural village", "乡村": "countryside village", "车辆": "fleet vehicles", "公交": "bus", "汽车": "car"}

# 依次尝试的开源图库，都不需要申请密钥。
# 两个都是聚合图库，每张图授权不同（公共领域、CC0、CC BY、CC BY-SA 等），授权与来源逐张记入 图片来源.json，
# 是否可用、是否需要署名由使用者自行核查。
SOURCES = ("wikimedia", "openverse")  # 维基共享资源直接给 1600px 缩略图，下载快，放第一
TIMEOUT = 5
MAX_SIDE = 1920
MAX_BYTES = 25 * 1024 * 1024  # 单张图下载上限
PREFETCH_WORKERS = 6
CACHE_NAME = ".图片缓存.json"
_BUDGET = {"deadline": None}
UA = {"User-Agent": "newx-ppt/1.0 (+https://github.com/ZJN2018/newx-ppt)"}


def _query(topic):
    """检索词：topic 里写了英文就只用英文部分（推荐写法“储能、battery energy storage”）；否则按中文词表翻译。"""
    import re
    eng = re.findall(r"[A-Za-z][A-Za-z0-9 \-]*", topic or "")
    if eng:
        return " ".join(w.strip() for w in eng if w.strip())
    words = _words(topic)
    en = [EN.get(w) or next((v for k, v in EN.items() if k in w), None) for w in words]
    en = [e for e in en if e]
    return " ".join(dict.fromkeys(" ".join(en).split())) if en else None


def _get_json(url, headers=None, timeout=TIMEOUT):
    import urllib.request
    req = urllib.request.Request(url, headers=dict(UA, **(headers or {})))
    return json.load(urllib.request.urlopen(req, timeout=timeout))


def _search(source, q, timeout=TIMEOUT):
    """返回 [{id, url, width, height, title, creator, license, license_url, source_page}]。"""
    import re
    import urllib.parse as U
    if source == "openverse":
        d = _get_json("https://api.openverse.org/v1/images/?" + U.urlencode(
            {"q": q, "license_type": "commercial", "aspect_ratio": "wide", "size": "large", "page_size": 15}), timeout=timeout)
        return [{"id": "ov-" + it["id"], "url": it["url"], "width": it.get("width") or 0, "height": it.get("height") or 0,
                 "title": it.get("title"), "creator": it.get("creator"),
                 "license": ("%s %s" % (it.get("license", ""), it.get("license_version", ""))).strip().upper(),
                 "license_url": it.get("license_url"), "source_page": it.get("foreign_landing_url")} for it in d.get("results", [])]
    if source == "wikimedia":
        d = _get_json("https://commons.wikimedia.org/w/api.php?" + U.urlencode(
            {"action": "query", "format": "json", "generator": "search", "gsrsearch": "filetype:bitmap " + q,
             "gsrnamespace": 6, "gsrlimit": 15, "prop": "imageinfo", "iiprop": "url|size|mime|extmetadata",
             "iiurlwidth": 1600}), timeout=timeout)
        out = []
        junk = re.compile(r"NARA|scan|document|page|letter|map|poster|diagram|chart|logo|drawing|plan|manuscript|card|stamp|"
                          r"\.(pdf|tif|tiff|djvu|svg|png)$", re.I)
        for pg in (d.get("query", {}).get("pages", {}) or {}).values():
            ii = (pg.get("imageinfo") or [{}])[0]
            if ii.get("mime") != "image/jpeg" or junk.search(pg.get("title", "")):
                continue  # 维基共享资源里档案扫描件、地图、海报很多，只要实拍照片
            md = ii.get("extmetadata", {})
            val = lambda k: (md.get(k) or {}).get("value")
            artist = re.sub(r"<[^>]+>", " ", val("Artist") or "")
            artist = re.sub(r"\s+", " ", artist).strip()
            out.append({"id": "wm-%s" % pg.get("pageid"), "url": ii.get("thumburl") or ii.get("url"),
                        "width": ii.get("thumbwidth") or ii.get("width") or 0, "height": ii.get("thumbheight") or ii.get("height") or 0,
                        "title": pg.get("title"), "creator": artist, "license": val("LicenseShortName"),
                        "license_url": val("LicenseUrl"), "source_page": ii.get("descriptionurl")})
        return out
    return []


# ---------------- 取图提速：缓存、并行预取、总时长上限 ----------------
# 联网取图只是配角，不能拖住整套生成：
# 1. 缓存：检索结果和下载过的图记在 out_dir/.图片缓存.json，重新生成直接复用，不再联网；
# 2. 预取：生成开始时把整稿要的图并行检索、并行下载（PREFETCH_WORKERS 路）；
# 3. 上限：单次请求 TIMEOUT 秒，整套联网取图合计不超过 NIUPPT_PHOTO_BUDGET 秒（默认 90），
#    超时的图退回自带图库，仍没有就不放图，不再等；
# 4. 下小图：下载后长边压到 MAX_SIDE 像素，PPTX 不会被几 MB 的原图撑大。
# （常量 TIMEOUT、MAX_SIDE 等定义在 SOURCES 下方）


def budget_start(seconds=None):
    """开始计时：之后的联网检索和下载合计不超过 seconds 秒。"""
    import time
    if seconds is None:
        seconds = float(os.environ.get("NIUPPT_PHOTO_BUDGET", "90"))
    _BUDGET["deadline"] = time.time() + seconds
    _NET["down"] = False


def _left():
    import time
    d = _BUDGET["deadline"]
    return None if d is None else d - time.time()


def _net_ok():
    if os.environ.get("NIUPPT_OFFLINE") == "1":
        return False
    left = _left()
    return not _NET["down"] and (left is None or left > 0.5)


def _timeout():
    left = _left()
    return TIMEOUT if left is None else max(0.5, min(TIMEOUT, left))


def _cache(out_dir):
    path = os.path.join(out_dir, CACHE_NAME)
    try:
        c = json.load(open(path, encoding="utf-8"))
    except Exception:
        c = {}
    c.setdefault("search", {})
    c.setdefault("files", {})
    return c


def _cache_save(out_dir, c):
    json.dump(c, open(os.path.join(out_dir, CACHE_NAME), "w", encoding="utf-8"), ensure_ascii=False)


def _usable(item, used=None):
    if not item.get("url") or (used is not None and item["id"] in used):
        return False
    w, h = item.get("width") or 0, item.get("height") or 0
    return w >= 1200 and not (h and w / h < 1.2)


def _dst(out_dir, item):
    return os.path.join(out_dir, "img_%s.jpg" % item["id"].replace("/", "_")[:24])


def _sock_of(resp):
    """取 HTTP 响应底层的 socket，用来把每次读取的等待上限设成剩余预算；取不到返回 None。"""
    try:
        return resp.fp.raw._sock
    except AttributeError:
        return None


def _download(item, dst):
    """下载一张图，长边压到 MAX_SIDE；失败返回 False。
    整套时长预算对下载全程有效：每次只读当前到达的一批数据（read1），
    读之前把 socket 等待上限设成剩余预算，服务端滴数据或中途停住都会按时退出。"""
    import io
    import urllib.request
    try:
        req = urllib.request.Request(item["url"], headers=UA)
        chunks, size = [], 0
        with urllib.request.urlopen(req, timeout=_timeout()) as r:
            sock = _sock_of(r)
            while True:
                left = _left()
                if left is not None and left <= 0:
                    raise TimeoutError("photo budget exceeded")
                wait = TIMEOUT if left is None else max(0.05, min(TIMEOUT, left))
                # 正文读完后 urllib 会自动关闭连接，此时再设超时会报 EBADF；
                # 已关闭就跳过，交给 read1 返回空字节正常结束。
                if sock is not None and not r.isclosed():
                    try:
                        sock.settimeout(wait)
                    except OSError:
                        pass
                b = r.read1(64 * 1024)
                if not b:
                    break
                size += len(b)
                if size > MAX_BYTES:
                    raise ValueError("image too large")
                chunks.append(b)
        # 在内存里校验和压缩，不写临时文件：用户连接的文件夹可能不允许删除文件，
        # 先写 .part 再删的做法在那里会把每张图都判成失败
        data = b"".join(chunks)
        Image.open(io.BytesIO(data)).verify()
        im = Image.open(io.BytesIO(data)).convert("RGB")
        im.thumbnail((MAX_SIDE, MAX_SIDE))
        im.save(dst, quality=88)
        return True
    except Exception:
        return False


def _record(out_dir, topic, q, src, item, dst):
    rec_path = os.path.join(out_dir, "图片来源.json")
    records = json.load(open(rec_path, encoding="utf-8")) if os.path.exists(rec_path) else []
    if any(r.get("file") == os.path.basename(dst) for r in records):
        return
    records.append({"file": os.path.basename(dst), "topic": topic, "query": q, "source": src,
                    "title": item.get("title"), "creator": item.get("creator"), "license": item.get("license"),
                    "license_url": item.get("license_url"), "source_page": item.get("source_page")})
    json.dump(records, open(rec_path, "w", encoding="utf-8"), ensure_ascii=False, indent=1)


def _take(out_dir, topic, q, src, item, dst, used):
    _record(out_dir, topic, q, src, item, dst)
    if used is not None:
        used.add(item["id"])
    return dst


def fetch(topic, out_dir, used=None, timeout=None, sources=None, allow_net=True):
    """取一张开源图库的横图，记入 out_dir/图片来源.json。分两遍：
    1. 只读本地缓存：按图库和候选顺序找已经下载过的图，没下载的候选跳过、继续往后找；
    2. 缓存全部没命中，且 allow_net 为真、未离线、还在时长上限内，才联网检索和下载。
    断网或 NIUPPT_OFFLINE=1 重建时，上次用过的图照样从缓存取回，不发任何请求。
    全部失败返回 None（调用方退回自带图库，仍没有就不放图）。"""
    q = _query(topic)
    if not q or not out_dir:
        return None
    srcs = list(sources or SOURCES)
    c = _cache(out_dir)
    for src in srcs:  # 第 1 遍：缓存
        for item in c["search"].get("%s|%s" % (src, q)) or []:
            if not _usable(item, used):
                continue
            dst = _dst(out_dir, item)
            if c["files"].get(item["id"]) and os.path.exists(dst):
                return _take(out_dir, topic, q, src, item, dst, used)
    if not (allow_net and _net_ok()):
        return None
    os.makedirs(out_dir, exist_ok=True)
    fails = tried = 0
    for src in srcs:  # 第 2 遍：联网
        key = "%s|%s" % (src, q)
        items = c["search"].get(key)
        if items is None:
            if not _net_ok():
                continue
            tried += 1
            try:
                items = _search(src, q, timeout or _timeout())
            except Exception:
                fails += 1
                continue
            c["search"][key] = items
            _cache_save(out_dir, c)
        for item in items:
            if not _usable(item, used):
                continue
            if not _net_ok():
                return None
            dst = _dst(out_dir, item)
            if not _download(item, dst):
                continue
            c["files"][item["id"]] = os.path.basename(dst)
            _cache_save(out_dir, c)
            return _take(out_dir, topic, q, src, item, dst, used)
    if tried and fails == tried and fails == len(srcs):  # 所有图库都连不上：本次不再联网
        _NET["down"] = True
        print("提醒：连不上开源图库，本次改用自带图库")
    return None


def prefetch(topics, out_dir, sources=None, workers=PREFETCH_WORKERS):
    """把整稿要用的图并行检索、并行下载进缓存，之后逐页取图直接命中缓存。
    topics 里的本机路径、图库文件名会跳过。返回 (下载张数, 用时秒)。"""
    import time
    from concurrent.futures import ThreadPoolExecutor
    t0 = time.time()
    if not _net_ok():  # 离线或时长已用完：不发任何请求
        return 0, 0.0
    qs = []
    for t in topics:
        f = _local_file(t)
        q = None if f else _query(t)
        if q and q not in qs:
            qs.append(q)
    if not qs:
        return 0, 0.0
    os.makedirs(out_dir, exist_ok=True)
    c = _cache(out_dir)
    srcs = list(sources or SOURCES)
    jobs = [(s, q) for q in qs for s in srcs if "%s|%s" % (s, q) not in c["search"]]

    def search(job):
        s, q = job
        try:
            return job, _search(s, q, _timeout()) if _net_ok() else None
        except Exception:
            return job, None
    if jobs:
        got = 0
        with ThreadPoolExecutor(min(workers, len(jobs))) as ex:
            for (s, q), items in ex.map(search, jobs):
                if items is not None:
                    c["search"]["%s|%s" % (s, q)] = items
                    got += 1
        if not got:  # 一个都没连上：本次生成不再联网，逐页取图不再挨个等超时
            _NET["down"] = True
            print("提醒：连不上开源图库，本次改用自带图库")
    # 按顺序给每个主题挑一张没被挑过的，避免两个主题拿到同一张
    picked, want = set(), []
    for q in qs:
        for s in srcs:
            item = next((it for it in c["search"].get("%s|%s" % (s, q)) or [] if _usable(it, picked)), None)
            if item:
                picked.add(item["id"])
                if not (c["files"].get(item["id"]) and os.path.exists(_dst(out_dir, item))):
                    want.append(item)
                break

    def down(item):
        return item, _net_ok() and _download(item, _dst(out_dir, item))
    n = 0
    if want:
        with ThreadPoolExecutor(min(workers, len(want))) as ex:
            for item, ok in ex.map(down, want):
                if ok:
                    c["files"][item["id"]] = os.path.basename(_dst(out_dir, item))
                    n += 1
    _cache_save(out_dir, c)
    return n, time.time() - t0


def deck_topics(deck):
    """整稿要取图的主题：封面、目录、各章转场、正文图片格。"""
    st = deck.get("settings", {})
    ph = st.get("photos", {}) or {}
    out = [ph.get("cover"), ph.get("agenda")]
    out += [s.get("photo") for s in deck.get("sections", []) if isinstance(s, dict)]

    def walk(o):
        if isinstance(o, dict):
            if o.get("type") == "photo" and o.get("topic"):
                out.append(o["topic"])
            for v in o.values():
                walk(v)
        elif isinstance(o, list):
            for v in o:
                walk(v)
    walk(deck.get("pages", []))
    return [t for t in out if t]


_NET = {"down": False}

COPYRIGHT_NOTE = ("本稿使用了从开源图库联网下载的图片，来源与授权逐张记在 图片来源.json。"
                  "每张图授权不同（公共领域、CC0、CC BY 需署名、CC BY-SA 需署名并相同方式共享），"
                  "且授权不覆盖画面中的人物肖像与商标。对外使用前请逐张核查，需要署名的在页脚或末页注明作者与来源。")


STRONG = 6  # 自带图库“强相关”的分数：至少两个关键词命中


def online_default(settings_photos=None):
    """联网取图默认打开；deck.json 写 photos.online=false 或环境变量 NIUPPT_OFFLINE=1 时关闭。"""
    if os.environ.get("NIUPPT_OFFLINE") == "1":
        return False
    return bool((settings_photos or {}).get("online", True))


LIB_MAX_SHARE = 0.30  # 一套稿子里自带图库的图最多占 30%，其余从开源图库联网取（用户自己给的图不计）


def is_library(path):
    return bool(path) and os.path.dirname(os.path.abspath(path)) == os.path.abspath(DIR)


def get(topic, out_dir=None, used=None, online=False, sources=None, lib_ok=True):
    """取一张与主题相关的图，顺序：
    1. 本机图片路径（用户给的图），直接用；
    2. lib_ok=False（图库份额已用满）且能联网时，先到开源图库检索；
       指定的是图库文件名时，用该图的关键词去检索同类图；
    3. 自带图库里强相关（至少两个关键词命中）的图，或指定的图库文件；
    4. 联网时到开源图库检索（联网图未经人工看过，生成后必须看图核对）；
    5. 联网没取到时，退回自带图库里一般相关（一个关键词命中）的图；
    都没有时返回 None，不放图。"""
    f = _local_file(topic)
    if f and not is_library(f):
        if used is not None:
            used.add(os.path.basename(f))
        return f
    # 断网、离线也要进 fetch：它先读本地缓存，只有要发请求时才看 online 和联网状态
    net = bool(out_dir)
    if net and not lib_ok:
        q = topic
        if f:
            im = next((i for i in lib() if i["file"] == os.path.basename(f)), {})
            q = "、".join(im.get("keywords", [])[:4]) or im.get("desc", "")
        p = fetch(q, out_dir, used, sources=sources, allow_net=online) if q else None
        if p:
            return p
    if f:
        if used is not None:
            used.add(os.path.basename(f))
        return f
    strong = [im for sc, im in candidates(topic, used) if sc >= STRONG]
    if strong:
        if used is not None:
            used.add(strong[0]["file"])
        return os.path.join(DIR, strong[0]["file"])
    if out_dir:
        p = fetch(topic, out_dir, used, sources=sources, allow_net=online)
        if p:
            return p
    return pick(topic, used=used)


def cover_tone(path, main):
    """封面、结束页的调色：照片色调和主色接近时基本保留原色；
    照片明显偏另一种色（如冷色主色配暖黄照片）时加重向主色调色，避免封面和整套色系不一致。"""
    import colorsys as _cs
    try:
        im = Image.open(path).convert("RGB").resize((64, 64))
    except Exception:
        return {"sat": 0.88, "duo": 0.10}
    import math
    sx = sy = ws = 0.0
    for r, g, b in im.getdata():
        h, l, s = _cs.rgb_to_hls(r / 255, g / 255, b / 255)
        w = s * (1 - abs(2 * l - 1))  # 只算有颜色的像素
        sx += math.cos(h * 2 * math.pi) * w; sy += math.sin(h * 2 * math.pi) * w; ws += w
    n = 64 * 64
    if ws / n < 0.06:  # 本身接近黑白灰
        return {"sat": 0.88, "duo": 0.10}
    ph = (math.atan2(sy, sx) / (2 * math.pi)) % 1.0
    mh = _cs.rgb_to_hls(*(int(main[i:i + 2], 16) / 255 for i in (0, 2, 4)))[0]
    d = min(abs(ph - mh), 1 - abs(ph - mh)) * 360
    if d > 60:
        return {"sat": 0.55, "duo": 0.35}
    return {"sat": 0.88, "duo": 0.10}


def tone_for(path):
    """页内与转场照片的调色参数（保留饱和度, 主色调色占比）。
    要"柔和、不抢主色"，但不能压成灰白或整体发棕：原来 0.25/0.55 太狠，看起来像旧照片。
    食品类再轻一些，否则看不出是什么。"""
    name = os.path.basename(path)
    return (0.75, 0.08) if name.startswith("fresh") else (0.60, 0.12)


def main(argv=None):
    """python3 photos.py 搜索 会议、文件   列出匹配的自带图片（文件名、分数、画面）
    python3 photos.py 列表               列出全部自带图片
    python3 photos.py 预取 --out out 主题1 主题2 …   后台先把图下进缓存（故事线写好、等用户确认时就可以跑）
    python3 photos.py 预取 --out out --deck deck.json"""
    import sys
    a = list(argv if argv is not None else sys.argv[1:])
    if a and a[0] in ("预取", "prefetch"):
        import argparse
        ap = argparse.ArgumentParser(prog="photos.py 预取")
        ap.add_argument("--out", default="out")
        ap.add_argument("--deck")
        ap.add_argument("--sources", help="逗号分隔，如 wikimedia；不写时用稿件 settings.photos.sources，再不写用默认顺序")
        ap.add_argument("topics", nargs="*")
        o = ap.parse_args(a[1:])
        topics = list(o.topics)
        ph = {}
        if o.deck:
            deck = json.load(open(o.deck, encoding="utf-8"))
            ph = (deck.get("settings") or {}).get("photos") or {}
            topics += deck_topics(deck)
        if not online_default(ph):  # 与正式生成同一套开关：NIUPPT_OFFLINE=1 或 photos.online=false 时不联网
            print("预取图片：已关闭联网取图（NIUPPT_OFFLINE=1 或 photos.online=false），跳过")
            return 0
        sources = [x.strip() for x in o.sources.split(",") if x.strip()] if o.sources else ph.get("sources")
        budget_start()
        n, sec = prefetch(topics, o.out, sources)
        print("预取图片：新下载 %d 张，用时 %.0f 秒，缓存在 %s" % (n, sec, os.path.join(o.out, CACHE_NAME)))
        return 0
    if not a or a[0] in ("列表", "list"):
        for im in lib():
            print("%-18s %-6s %s%s" % (im["file"], im["category"], im.get("desc", ""), ("　〔不自动用：%s〕" % im["avoid"]) if im.get("avoid") else ""))
        return 0
    topic = "、".join(a[1:] if a[0] in ("搜索", "search") else a)
    res = candidates(topic, include_avoid=True)
    if not res:
        print("自带图库里没有与「%s」相关的图。可开启联网取图，或不放图。" % topic)
        return 1
    for sc, im in res[:12]:
        print("%-18s 分数 %d　%s%s" % (im["file"], sc, im.get("desc", ""), ("　〔不自动用：%s，可按文件名指定〕" % im["avoid"]) if im.get("avoid") else ""))
    return 0


if __name__ == "__main__":
    import sys
    sys.exit(main())

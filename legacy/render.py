"""把 PPTX 渲染成逐页 PNG 与一张总览图，供模型“看图自查”。

python3 render.py deck.pptx [--out 目录] [--dpi 80]
输出：<目录>/<名>.pdf、<名>-p-01.png …、<名>-overview.png（4 列缩略总览）

依赖 LibreOffice（soffice），逐页出图用 poppler（pdftoppm），没有时自动改用 PyMuPDF。若系统没有思源黑体 / 思源宋体，会自动建
fontconfig 别名，映射到同源的 Noto Sans / Serif CJK SC，使预览的行高、
换行接近在 Windows / macOS 上打开的效果。
"""
import os, sys, glob, shutil, subprocess, argparse

from functools import lru_cache
from pathlib import Path
from xml.etree import ElementTree as ET


@lru_cache()
def chinese_font_path():
    """Pillow 与无界面渲染使用同一个本机中文字体。"""
    from PIL import ImageFont
    font_dirs = [Path.home() / "Library/Fonts", Path("/Library/Fonts")]
    if sys.platform == "win32":
        font_dirs = [Path(os.environ.get("WINDIR", "C:/Windows")) / "Fonts"]
        if os.environ.get("LOCALAPPDATA"):
            font_dirs.insert(0, Path(os.environ["LOCALAPPDATA"]) / "Microsoft/Windows/Fonts")
    installed = [path for directory in font_dirs
                 for pattern in ("NotoSansCJKsc-Regular.otf", "SourceHanSansCN-Regular.otf",
                                 "NotoSansCJK*Regular*.otf", "SourceHanSans*Regular.otf")
                 for path in sorted(directory.glob(pattern))]
    if sys.platform == "win32":
        candidates = installed + [font_dirs[-1] / "msyh.ttc"]
    else:
        candidates = [
            "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
            *installed,
            "/System/Library/Fonts/PingFang.ttc",
            "/System/Library/Fonts/STHeiti Light.ttc",
            "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
            "/usr/share/fonts/truetype/droid/DroidSansFallbackFull.ttf",
        ]
    if shutil.which("fc-match"):
        match = subprocess.run(["fc-match", "-f", "%{file}", ":lang=zh-cn"],
                               text=True, capture_output=True)
        candidates.append(match.stdout.strip())
    for path in candidates:
        if Path(path).is_file():
            font = ImageFont.truetype(str(path), 20)
            if bytes(font.getmask("收")) != bytes(font.getmask("益")):
                return str(path)
    raise OSError("未找到可绘制中文的字体。请按 README 安装简体中文思源黑体 CN 或 Noto Sans CJK SC 的常规 OTF 字重。")


@lru_cache(maxsize=64)
def preview_font(size):
    from PIL import ImageFont
    return ImageFont.truetype(chinese_font_path(), size)


def render_env(work):
    """仅覆盖渲染子进程；保留系统和用户的已有字体配置。"""
    env = os.environ.copy()
    root = ET.Element("fontconfig")
    for config in [env.get("FONTCONFIG_FILE"), "/etc/fonts/fonts.conf", "/opt/homebrew/etc/fonts/fonts.conf"]:
        if config and Path(config).is_file():
            ET.SubElement(root, "include").text = config
            break
    for directory in {Path(chinese_font_path()).parent, Path.home() / "Library/Fonts",
                      Path("/Library/Fonts"), Path("/usr/share/fonts"), Path("/usr/local/share/fonts")}:
        if directory.is_dir():
            ET.SubElement(root, "dir").text = str(directory)
    ET.SubElement(root, "cachedir").text = str(Path(work, "cache"))
    for family in ("Source Han Sans CN", "思源黑体 CN", "Microsoft YaHei", "微软雅黑", "sans-serif"):
        alias = ET.SubElement(root, "alias", binding="strong")
        ET.SubElement(alias, "family").text = family
        ET.SubElement(ET.SubElement(alias, "prefer"), "family").text = preview_font(20).getname()[0]
    # 思源宋体与 Noto Serif CJK 同源，未安装时用后者预览
    for family in ("Source Han Serif CN", "思源宋体 CN"):
        alias = ET.SubElement(root, "alias", binding="strong")
        ET.SubElement(alias, "family").text = family
        prefer = ET.SubElement(alias, "prefer")
        for name in ("Source Han Serif CN", "Noto Serif CJK SC"):
            ET.SubElement(prefer, "family").text = name
    config = Path(work, "fonts.conf")
    ET.ElementTree(root).write(config, encoding="utf-8", xml_declaration=True)
    env["FONTCONFIG_FILE"] = str(config)
    return env

def rasterize(pdf, prefix, dpi):
    """PDF 逐页转 PNG：优先 poppler 的 pdftoppm，没有时用 PyMuPDF；两者都没有时说明缺什么。
    文件名与 pdftoppm 一致：<prefix>-1.png，页数 ≥10 时补零为 -01。"""
    if shutil.which("pdftoppm"):
        subprocess.run(["pdftoppm", "-png", "-r", str(dpi), pdf, prefix], check=True)
        return
    try:
        import fitz
    except ImportError:
        raise RuntimeError("逐页 PNG 需要 poppler（pdftoppm）或 PyMuPDF：macOS 执行 brew install poppler，"
                           "或 pip install pymupdf。PPTX 与 PDF 已生成。")
    try:  # LibreOffice 的带标签 PDF 结构树不完整，PyMuPDF 会逐页报错但不影响出图
        fitz.TOOLS.mupdf_display_errors(False)
    except AttributeError:
        pass
    with fitz.open(pdf) as doc:
        width = len(str(len(doc))) if len(doc) >= 10 else 1
        for i, page in enumerate(doc, 1):
            page.get_pixmap(dpi=dpi).save(f"{prefix}-{i:0{width}d}.png")

SOFFICE_HINT = ("未找到 LibreOffice，无法渲染自查。请按 README 的 macOS / Windows 安装说明操作。"
                "自定义安装目录可用 NIUPPT_SOFFICE 指向 soffice（Windows 为 soffice.exe）。")
_SOFFICE_PATHS = (
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    os.path.expanduser("~/Applications/LibreOffice.app/Contents/MacOS/soffice"),
    "C:/Program Files/LibreOffice/program/soffice.exe",
    "C:/Program Files (x86)/LibreOffice/program/soffice.exe",
    "/opt/libreoffice/program/soffice",
    "/snap/bin/libreoffice",
)


def find_soffice():
    """依次找 NIUPPT_SOFFICE、PATH 和各系统的常见安装位置；找不到返回 None。"""
    env = os.environ.get("NIUPPT_SOFFICE")
    if env:
        return env if os.path.isfile(env) else None
    return shutil.which("soffice") or shutil.which("libreoffice") or next(
        (p for p in _SOFFICE_PATHS if os.path.isfile(p)), None)


def render(deck, out=None, dpi=80, cols=4, *, compact_asian=True):
    out = out or os.path.join(os.path.dirname(os.path.abspath(deck)), "_render")
    os.makedirs(out, exist_ok=True)
    name = os.path.splitext(os.path.basename(deck))[0]
    soffice = find_soffice()
    if not soffice:
        print("[渲染] " + SOFFICE_HINT); return None
    import tempfile
    from pathlib import Path
    with tempfile.TemporaryDirectory() as work:
        profile = Path(work, "profile").as_uri()
        # 独立配置避免与已打开的 LibreOffice 共用进程；成功后再替换旧 PDF。
        def convert(source, fmt, target):
            proc = subprocess.run([soffice, "-env:UserInstallation=" + profile, "--headless",
                                   "--convert-to", fmt, "--outdir", work, str(source)],
                                  capture_output=True, timeout=300, env=render_env(work))
            if proc.returncode or not target.is_file():
                raise RuntimeError("LibreOffice 转换失败：" + (proc.stderr + proc.stdout).decode(errors="replace"))

        source = Path(deck).resolve()
        if compact_asian:
            # Impress imports PPTX paragraphs with a quarter-em gap at Latin/CJK
            # boundaries even when the editable source contains no spaces.
            from lxml import etree
            flat = Path(work, name + ".fodp")
            convert(source, "fodp", flat)
            tree = etree.parse(str(flat))
            ns = "urn:oasis:names:tc:opendocument:xmlns:style:1.0"
            for prop in tree.iter("{%s}paragraph-properties" % ns):
                prop.set("{%s}text-autospace" % ns, "none")
            tree.write(str(flat), encoding="UTF-8", xml_declaration=True)
            source = flat
        converted = Path(work, name + ".pdf")
        convert(source, "pdf", converted)
        pdf = os.path.join(out, name + ".pdf")
        shutil.copyfile(converted, pdf)
    for f in glob.glob(os.path.join(out, f"{name}-p-*.png")): os.remove(f)
    rasterize(pdf, os.path.join(out, f"{name}-p"), dpi)
    pngs = sorted(glob.glob(os.path.join(out, f"{name}-p-*.png")))
    try:
        from PIL import Image, ImageDraw
        ims = [Image.open(p) for p in pngs]
        w, h = 480, int(480 * ims[0].height / ims[0].width)
        rows = (len(ims) + cols - 1) // cols
        sheet = Image.new("RGB", (cols * (w + 12) + 12, rows * (h + 34) + 12), "#E9EAEE")
        d = ImageDraw.Draw(sheet)
        for i, im in enumerate(ims):
            x = 12 + (i % cols) * (w + 12); y = 12 + (i // cols) * (h + 34)
            sheet.paste(im.resize((w, h)), (x, y + 22)); d.text((x, y + 4), f"P{i + 1}", fill="#333333")
        ov = os.path.join(out, f"{name}-overview.png"); sheet.save(ov)
        print(f"[渲染] {len(pngs)} 页 → {out}（总览 {os.path.basename(ov)}）")
    except ImportError:
        print(f"[渲染] {len(pngs)} 页 → {out}（未装 Pillow，未生成总览）")
    return pngs

if __name__ == "__main__":
    ap = argparse.ArgumentParser(); ap.add_argument("deck"); ap.add_argument("--out"); ap.add_argument("--dpi", type=int, default=80)
    a = ap.parse_args(); render(a.deck, a.out, a.dpi)

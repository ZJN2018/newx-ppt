"""首次安装后，用一份两页的样稿验证 PPTX、PDF 与逐页 PNG 的生成。"""
import argparse
import subprocess
import sys
import tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--out", type=Path, help="保留样页的新目录，默认使用临时目录")
    args = parser.parse_args()
    from render import SOFFICE_HINT, chinese_font_path, find_soffice
    soffice = find_soffice()
    if not soffice:
        parser.exit(1, SOFFICE_HINT + "\n")
    try:
        import svgelements  # noqa: F401  图标转换依赖
        from PIL import Image
        import build
        print("LibreOffice:", soffice)
        font = Path(chinese_font_path())
        print("中文预览字体:", font)
        if not font.name.lower().startswith(("sourcehansans", "notosanscjk")):
            print("预览采用系统替代字体。要按稿件指定字体排版，请安装 README 所列字体。")
        deck = {"settings": {"title": "中文排版试装", "subtitle": "安装检查", "preset": "藏蓝", "footer": "安装检查",
                             "filename": "install-check"},
                "pages": [{"title": "安装检查：图标、卡片与表格均可生成", "separation": "纯卡片",
                           "partition": {"cut": "v", "ratio": [1, 1], "children": [
                               {"cut": "h", "ratio": [1, 1], "children": [
                                   {"role": "主证据", "title": "图标要点", "element": {"type": "points", "items": [
                                       ["收入增长 62.2%", "检查数字与中文换行", "trending-up"], ["风险可控", "检查图标转换", "protect"]]}},
                                   {"role": "辅证据", "title": "大数字", "element": {"type": "num_pair", "items": [
                                       ["62.2", "%", "收入同比增速"], ["3.0", "亿元", "样例金额"]]}}]},
                               {"cut": "h", "ratio": [2, 3], "children": [
                                   {"role": "辅证据", "title": "表格", "element": {"type": "table", "header": ["项目", "数值"],
                                                                                "rows": [["甲", "1.0"], ["乙", "2.0"], ["丙", "3.0"]]}},
                                   {"role": "结论", "title": "段落", "element": {"type": "para", "lead": "中文段落与换行检查",
                                    "body": "本格检查加粗导语与正文段落能否正常排版：数字 62.2%、单位亿元、中英文混排 PPTX 与标点挤压都应在格内换行，不越界、不重叠。"}}]}]}}]}
        with tempfile.TemporaryDirectory() as work:
            import os
            os.environ.setdefault("PPT_LEDGER", os.path.join(work, "ledger.jsonl"))
            out = Path(work) if args.out is None else args.out
            if args.out is not None:
                out.mkdir(parents=True, exist_ok=False)
            pptx = build.build(deck, str(out))
            if not pptx:
                raise RuntimeError("样稿规格检查未通过")
            pngs = sorted(Path(out).glob("install-check-p-*.png"))
            if len(pngs) != 3 or not Path(pptx).with_suffix(".pdf").is_file():
                raise RuntimeError("PDF 或逐页 PNG 生成失败")
            with Image.open(pngs[1]) as preview:
                preview.verify()
            print("已生成样稿 PPTX、PDF 和 PNG，安装检查通过。")
            if args.out is not None:
                print("样页目录:", out.resolve())
    except (ImportError, OSError, RuntimeError, subprocess.SubprocessError) as error:
        parser.exit(1, f"安装检查未通过：{error}\n请按 README 检查依赖与输出目录。\n")


if __name__ == "__main__":
    main()

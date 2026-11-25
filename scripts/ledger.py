"""交付台账：记录每套成稿的配色与逐页区隔方式，开工前看最近几套，避免新稿与之雷同。

python3 ledger.py show [--last 5]
台账文件默认是当前目录的 ppt-ledger.jsonl，可用环境变量 PPT_LEDGER 指定。build.py 每次生成自动追加一条。
"""
import argparse
import datetime
import json
import os


def ledger_path():
    return os.environ.get("PPT_LEDGER") or os.path.abspath("ppt-ledger.jsonl")


def load():
    path = ledger_path()
    if not os.path.isfile(path):
        return []
    with open(path, encoding="utf-8") as f:
        return [json.loads(line) for line in f if line.strip()]


def append(deck, pal):
    pages = deck.get("pages", [])
    entry = {"date": datetime.date.today().isoformat(), "title": deck.get("settings", {}).get("title", ""),
             "preset": pal.get("preset"), "main": pal.get("main"), "pages": len(pages),
             "separations": [p.get("separation") for p in pages]}
    with open(ledger_path(), "a", encoding="utf-8") as f:
        f.write(json.dumps(entry, ensure_ascii=False) + "\n")


def show(last=5):
    es = load()[-last:]
    if not es:
        print("台账为空（%s）" % ledger_path())
        return
    for e in es:
        print("%s  《%s》  %s %s  %d 页" % (e["date"], e["title"], e.get("preset"), e.get("main"), e["pages"]))
        print("   区隔：" + " → ".join(s or "-" for s in e.get("separations", [])))


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("cmd", choices=["show"])
    ap.add_argument("--last", type=int, default=5)
    a = ap.parse_args()
    show(a.last)

"""写 deck.json 之前检查论据底稿，并由实际缺口生成第二轮材料清单。

底稿是带一个 ``evidence-json`` 围栏的 Markdown，或同结构的 JSON。
每页包含 section、title、midpoints；每个中观点包含 title、evidence。
每条论据包含 text、source={material, locator}；带数字的论据另写
numeric={scope, comparison}，派生数据另写 derived={label, formula}。
本脚本只校验完整性，不能替代对原始材料的忠实核对。
"""
from __future__ import annotations

import argparse
import json
import re
from pathlib import Path

MIN_MIDPOINTS = 2
MIN_EVIDENCE = 4
MIN_BODY_CHARS = 160
REMEDIATION_ORDER = ("合并相邻同主题页", "降为相邻页内区域", "列入补材料清单")


def body_chars(text):
    return len(re.sub(r"\s+", "", str(text or "")))


def load_evidence(path):
    """读取底稿；只接受一个明确的数据块，避免误把普通故事线当成论据。"""
    path = Path(path)
    text = path.read_text(encoding="utf-8")
    if path.suffix.lower() == ".json":
        data = json.loads(text)
    else:
        blocks = re.findall(r"^```evidence-json\s*\n(.*?)^```\s*$", text, re.M | re.S)
        if len(blocks) != 1:
            raise ValueError("evidence.md 必须包含且仅包含一个 evidence-json 数据块")
        data = json.loads(blocks[0])
    if not isinstance(data, dict) or not isinstance(data.get("pages"), list) or not data["pages"]:
        raise ValueError("论据底稿必须包含非空 pages 列表")
    return data


def _normalized(text):
    return re.sub(r"[\s，,。；;：:、（）()%％]", "", str(text or ""))


# 日期、年份、序号不算计量数字：只含这些的论据不要求写 numeric
_DATE_LIKE = re.compile(r"(19|20)\d{2}\s*[年./-]|\d{1,2}\s*[月日号]|第\s*\d+|[QH]\d|\d{1,2}[:：]\d{2}")


def measured_numbers(text):
    """去掉日期、年份、序号后，正文里剩下的计量数字片段。"""
    rest = _DATE_LIKE.sub(" ", str(text or ""))
    rest = re.sub(r"(19|20)\d{2}(?!\d)", " ", rest)          # 单独出现的年份
    return re.findall(r"\d[\d.,]*\s*(?:%|％|亿元?|万元?|元|个百分点|pct|倍|[^\s，。；、]{0,2})", rest)


def _in_body(part, text):
    """口径或比较基准是否引自正文：整段命中，或按逗号、分号拆开的各段依次出现在正文中。"""
    body = _normalized(text)
    whole = _normalized(part)
    if whole and whole in body:
        return True
    pos = 0
    for seg in [x for x in re.split(r"[，,；;、]", str(part or "")) if _normalized(x)]:
        hit = body.find(_normalized(seg), pos)
        if hit < 0:
            return False
        pos = hit + len(_normalized(seg))
    return pos > 0


def _provided(text):
    value = str(text or "").strip()
    return bool(value) and not re.search(r"^(待补|未提供|未披露|未知|todo|n/?a|[-—]+)$", value, re.I)


def _issue(issues, code, detail, midpoint=None, evidence=None):
    item = {"code": code, "detail": detail}
    if midpoint is not None:
        item["midpoint"] = midpoint
    if evidence is not None:
        item["evidence"] = evidence
    issues.append(item)


def validate_evidence(data):
    """返回实际容量输入和缺口；正文计数不含标题、出处、口径等元数据。"""
    pages = data.get("pages", [])
    if not isinstance(pages, list) or not pages:
        raise ValueError("论据底稿必须包含非空 pages 列表")
    report = {"passed": True, "pages": [], "remediation_order": list(REMEDIATION_ORDER)}
    for index, page in enumerate(pages):
        if not isinstance(page, dict):
            raise ValueError("每页必须是对象")
        if page.get("kind", "content") != "content":
            continue
        issues, loads = [], []
        midpoints = page.get("midpoints", [])
        if not isinstance(midpoints, list):
            raise ValueError("midpoints 必须是列表")
        if not str(page.get("title", "")).strip():
            _issue(issues, "page_title", "缺少已确认的页面观点")
        if not str(page.get("section", "")).strip():
            _issue(issues, "section", "缺少部分名称，无法判断相邻同主题页")
        if len(midpoints) < MIN_MIDPOINTS:
            _issue(issues, "midpoint_count", "中观点不足2个")
        count = chars = 0
        for mi, point in enumerate(midpoints, 1):
            if not isinstance(point, dict):
                raise ValueError("每个中观点必须是对象")
            title = str(point.get("title", "")).strip()
            if not title:
                _issue(issues, "midpoint_title", "缺少中观点", mi)
            items = point.get("evidence", [])
            if not isinstance(items, list):
                raise ValueError("evidence 必须是列表")
            actual = [item for item in items if isinstance(item, dict) and str(item.get("text", "")).strip()]
            if len(actual) < 2:
                _issue(issues, "point_evidence_count", "论据不足2条，需补数字、事实或案例", mi)
            numeric_count = 0
            point_chars = 0
            for ei, item in enumerate(items, 1):
                if not isinstance(item, dict):
                    raise ValueError("每条论据必须是对象")
                text = str(item.get("text", "")).strip()
                if not text:
                    _issue(issues, "evidence_text", "论据正文为空", mi, ei)
                    continue
                count += 1
                point_chars += body_chars(text)
                source = item.get("source", {})
                if not isinstance(source, dict) or not all(_provided(source.get(k, "")) for k in ("material", "locator")):
                    _issue(issues, "source", "补原始材料名及页码或章节", mi, ei)
                numeric = item.get("numeric")
                numbers = measured_numbers(text)
                has_number = bool(numbers)
                derived_ok = (isinstance(item.get("derived"), dict) and item["derived"].get("label") == "整理者测算"
                              and str(item["derived"].get("formula", "")).strip())
                # 派生数的比较方式由算式说明，只要求口径写进正文；原始数仍要求口径与比较基准都引自正文
                keys = ("scope",) if derived_ok else ("scope", "comparison")
                missing = [k for k in keys if not (isinstance(numeric, dict) and _provided(numeric.get(k, ""))
                                                  and _in_body(numeric[k], text))]
                valid_numeric = isinstance(numeric, dict) and has_number and not missing
                if valid_numeric:
                    numeric_count += 1
                elif not has_number and numeric is None:
                    pass
                elif not has_number:
                    _issue(issues, "numeric_without_number",
                           "这条论据没有计量数字（日期、年份不算），删去 numeric 字段", mi, ei)
                elif not isinstance(numeric, dict):
                    _issue(issues, "numeric_missing",
                           f"检测到数字“{numbers[0].strip()}”，补 numeric：scope（口径）和 comparison（比较基准）各从正文摘一段原文", mi, ei)
                else:
                    detail = "；".join(
                        f"{k} 为空" if not _provided(numeric.get(k, "")) else
                        f"{k}“{numeric.get(k, '')}”在正文里找不到（各段须按原文顺序出现，可在正文补上这段话或改 {k}）"
                        for k in missing)
                    _issue(issues, "numeric_basis", detail, mi, ei)
                derived = item.get("derived")
                if derived is not None or "整理者测算" in text:
                    if not isinstance(derived, dict) or derived.get("label") != "整理者测算" or not str(derived.get("formula", "")).strip():
                        _issue(issues, "derived_formula", "派生数据需标整理者测算并写明算法", mi, ei)
            if numeric_count < 1:
                _issue(issues, "point_numeric", "至少补1条带口径和比较基准的数字论据", mi)
            for gap in point.get("gaps", []):
                _issue(issues, "material_gap", str(gap), mi)
            chars += point_chars
            loads.append({"midpoint": mi, "title": title, "evidence_count": len(actual),
                          "body_chars": point_chars, "numeric_count": numeric_count})
        if count < MIN_EVIDENCE:
            _issue(issues, "page_evidence_count", "全页论据不足4条")
        if chars < MIN_BODY_CHARS:
            _issue(issues, "body_chars", "全页论据正文不足160字（不计标题和出处）")
        adjacent = [i + 1 for i in (index - 1, index + 1) if 0 <= i < len(pages)
                    and isinstance(pages[i], dict) and pages[i].get("kind", "content") == "content"
                    and page.get("section") and pages[i].get("section") == page["section"]]
        report["pages"].append({"page": page.get("page", index + 1), "title": page.get("title", ""),
                                "section": page.get("section", ""), "midpoint_count": len(midpoints),
                                "evidence_count": count, "body_chars": chars, "region_loads": loads,
                                "adjacent_same_theme": [pages[i - 1].get("page", i) for i in adjacent],
                                "issues": issues, "passed": not issues})
        report["passed"] = report["passed"] and not issues
    if not report["pages"]:
        raise ValueError("论据底稿没有内容页")
    return report


def materials_checklist(report):
    lines = ["# 第二轮材料清单", "", "本清单由论据底稿生成；只列真实缺口，不补写未披露的数据。", ""]
    failed = [page for page in report["pages"] if not page["passed"]]
    if not failed:
        lines.append("当前论据完整性与页面充实度检查无缺口，请用户确认是否补充材料；确认后才能写 deck.json。")
    for page in failed:
        lines += ["## P%s %s" % (page["page"], page["title"]), ""]
        for issue in page["issues"]:
            where = "中观点%s：" % issue["midpoint"] if "midpoint" in issue else "页面："
            if "evidence" in issue:
                where += "论据%s，" % issue["evidence"]
            lines.append("- " + where + issue["detail"])
        lines += ["", "处理顺序：" + " → ".join(REMEDIATION_ORDER) + "。"]
        if page["adjacent_same_theme"]:
            lines.append("相邻同主题候选：" + "、".join("P%s" % p for p in page["adjacent_same_theme"]) + "。先核对观点是否适合合并。")
        else:
            lines.append("未发现相邻同部分内容页；核对主题后再决定是否降为相邻页区域。")
        lines += ["不得用空框、大色块或放大字号补足页面。", ""]
    return "\n".join(lines) + "\n"


def require_evidence(path):
    """供成稿生成器复用；不达标时拒绝进入绘制。"""
    data = load_evidence(path)
    report = validate_evidence(data)
    if not report["passed"]:
        raise ValueError("论据底稿未通过，先合并/降区/补材料：\n" + materials_checklist(report))
    return data, report


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("evidence", help="evidence.md 或同结构 JSON")
    parser.add_argument("--checklist", help="按实际缺口输出第二轮材料清单")
    parser.add_argument("--json", dest="json_out", help="输出充实度与区域实际论据量")
    args = parser.parse_args(argv)
    try:
        report = validate_evidence(load_evidence(args.evidence))
    except (ValueError, OSError) as exc:
        print("论据底稿读取失败：" + str(exc))
        return 2
    if args.checklist:
        Path(args.checklist).write_text(materials_checklist(report), encoding="utf-8")
    if args.json_out:
        Path(args.json_out).write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    for page in report["pages"]:
        print("P%s 中观点%d / 论据%d / 正文%d字：%s" % (page["page"], page["midpoint_count"],
              page["evidence_count"], page["body_chars"], "通过" if page["passed"] else "待处理"))
    return 0 if report["passed"] else 1


if __name__ == "__main__":
    raise SystemExit(main())

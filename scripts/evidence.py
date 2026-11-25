"""故事线的可选网页补充。默认关闭，不访问网络，也不调用付费搜索或付费 API。

只有用户明确要求补充最新公开信息时才打开：

    NIUPPT_WEB_SEARCH=1

打开之后，本模块仍然不会自己去请求任何搜索接口。执行的人如果手头没有不另计费的检索，
就在大纲里写「未检索」，继续用已有材料。不要为了演示去调用付费 API。
"""
from __future__ import annotations

import os


def enabled():
    return os.environ.get("NIUPPT_WEB_SEARCH", "").strip().lower() in {"1", "true", "yes", "on"}


def search(query):
    """不联网。关闭时说明默认关闭；打开时也只返回空结果，避免误调付费接口。"""
    if not enabled():
        return {
            "enabled": False,
            "query": query or "",
            "results": [],
            "note": "网页补充默认关闭，未访问网络，未调用搜索接口。",
        }
    return {
        "enabled": True,
        "query": query or "",
        "results": [],
        "note": "NIUPPT_WEB_SEARCH 已打开，但这里不调用付费搜索。请用手头材料，或由人工检索后写入大纲。未自动联网。",
    }

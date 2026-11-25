"""公开输出的 PPTX：统一字体，移除个人属性和嵌入缩略图。"""
import os
import re
import tempfile
import zipfile

FONT = "Noto Sans CJK SC"


def _blank_marked_elements(xml_bytes, names):
    for name in names:
        pattern = re.compile(
            br"(<(?:[A-Za-z0-9]+:)?" + name.encode("ascii") + br"(?:\s[^>]*)?>)(.*?)(</(?:[A-Za-z0-9]+:)?" + name.encode("ascii") + br">)",
            re.DOTALL,
        )
        xml_bytes = pattern.sub(lambda match: match.group(1) + match.group(3), xml_bytes)
    return xml_bytes


def _replace_typefaces(xml_bytes):
    font = FONT.encode("utf-8")

    def swap(match):
        face = match.group(1)
        # 空的 ea/cs 槽位留空时，+mn-ea 这类主题引用无法解析，出处扫描会告警。
        if face.startswith(b"+") or face == font:
            return match.group(0)
        quote = match.group(0)[9:10]  # typeface= 后面的引号
        return b"typeface=" + quote + font + quote

    xml_bytes = re.sub(br'typeface="([^"]*)"', swap, xml_bytes)
    xml_bytes = re.sub(br"typeface='([^']*)'", swap, xml_bytes)
    return xml_bytes


def _is_thumbnail(name):
    lowered = name.replace("\\", "/").lower()
    return lowered.startswith("docprops/thumbnail.")


def _drop_thumbnail_refs(name, data):
    """去掉指向嵌入缩略图的关系，避免删掉文件后包里还留着断链。"""
    if not (name.endswith(".rels") or name == "[Content_Types].xml"):
        return data
    text = data.decode("utf-8")
    text = re.sub(r"<Relationship\b[^>]*[Tt]humbnail[^>]*/>", "", text)
    text = re.sub(r"<Override\b[^>]*[Tt]humbnail[^>]*/>", "", text)
    return text.encode("utf-8")


def sanitize_pptx(path):
    """清掉作者、说明、嵌入缩略图，以及主题里的商业字体名。

    主题里的各文种字体名一律换成可商用的 Noto Sans CJK SC。
    以 + 开头的主题引用（+mj-lt 这类）保留，由主题表解析到上面的家族名。
    """
    fd, temporary = tempfile.mkstemp(dir=os.path.dirname(os.path.abspath(path)), suffix=".pptx")
    os.close(fd)
    try:
        with zipfile.ZipFile(path, "r") as source, zipfile.ZipFile(temporary, "w") as target:
            for info in source.infolist():
                name = info.filename.replace("\\", "/")
                if _is_thumbnail(name):
                    continue
                data = source.read(info.filename)
                if name.endswith(".xml") or name.endswith(".rels"):
                    if name == "docProps/core.xml":
                        data = _blank_marked_elements(
                            data,
                            ("creator", "lastModifiedBy", "description", "title", "subject", "keywords", "category"),
                        )
                    elif name == "docProps/app.xml":
                        data = _blank_marked_elements(data, ("Company", "Template"))
                    data = _drop_thumbnail_refs(name, data)
                    if name.endswith(".xml"):
                        data = _replace_typefaces(data)
                copied = zipfile.ZipInfo(filename=info.filename, date_time=info.date_time)
                copied.compress_type = info.compress_type or zipfile.ZIP_DEFLATED
                copied.external_attr = info.external_attr
                target.writestr(copied, data)
        os.replace(temporary, path)
    finally:
        if os.path.exists(temporary):
            os.unlink(temporary)

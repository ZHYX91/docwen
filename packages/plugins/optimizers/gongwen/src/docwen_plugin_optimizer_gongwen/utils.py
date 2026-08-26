"""Shared utility functions for gongwen processing."""

from __future__ import annotations

import re

from docwen_plugin_optimizer_gongwen.constants import (
    COPY_TO_LABELS,
    RE_DOC_NUMBER,
)

_COPY_TO_LABEL_RE = re.compile(rf"^(?:{'|'.join(re.escape(label) for label in COPY_TO_LABELS)})\s*[：:]\s*")


def contains_chinese(text: str) -> bool:
    """Check if text contains Chinese characters."""
    return bool(re.search(r"[一-鿿]", text))


def convert_date_format(text: str) -> str:
    """Normalize date formats to standard Chinese format (YYYY年M月D日).

    Handles: 2024年1月15日, 2024-01-15, 2024/01/15, etc.
    """
    if not text:
        return ""

    # Already in Chinese format
    if re.match(r"\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日", text):
        # Normalize whitespace
        return re.sub(r"\s+", "", text)

    # ISO format: 2024-01-15 or 2024/01/15
    m = re.match(r"(\d{4})[/-](\d{1,2})[/-](\d{1,2})", text)
    if m:
        y, mo, d = m.groups()
        return f"{y}年{int(mo)}月{int(d)}日"

    return text


def extract_after_colon(text: str) -> str:
    """Extract text after the last colon (： or :)."""
    for sep in ("：", ":"):
        if sep in text:
            return text.rsplit(sep, 1)[-1].strip()
    return text


def extract_combined_id(text: str) -> tuple[str, str]:
    """Split a compact ``份号 + 发文字号`` line without guessing by spaces."""

    match = re.fullmatch(r"\s*(\d+)\s*(?=[A-Za-z一-鿿])(.+?)\s*", text)
    if match is None:
        return "", ""
    copy_id, doc_number = match.groups()
    if RE_DOC_NUMBER.fullmatch(doc_number) is None:
        return "", ""
    return copy_id, doc_number.strip()


def extract_doc_number_and_signers(text: str) -> tuple[str, list[str]]:
    """Extract both fields from ``发文字号  签发人：姓名``."""

    match = re.fullmatch(r"\s*(.+?)\s+签发人[：:]\s*(.+?)\s*", text)
    if match is None:
        return "", []
    doc_number, signer_text = match.groups()
    if RE_DOC_NUMBER.fullmatch(doc_number.strip()) is None:
        return "", []
    return doc_number.strip(), extract_signers_from_text(signer_text)


def extract_doc_number_and_name(text: str) -> tuple[str, str]:
    """Extract the continuation form ``发文字号  姓名``."""

    match = re.fullmatch(r"\s*(.+?)\s+([^\s]+)\s*", text)
    if match is None:
        return "", ""
    doc_number, name = match.groups()
    if RE_DOC_NUMBER.fullmatch(doc_number.strip()) is None:
        return "", ""
    names = extract_signers_from_text(name)
    return (doc_number.strip(), names[0]) if names == [name] else ("", "")


def extract_printing_line(text: str) -> tuple[str, str]:
    """Split an optional printing authority from a trailing printing date."""

    match = re.fullmatch(
        r"\s*(?P<authority>.*?)(?P<date>\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日)\s*印发\s*",
        text,
    )
    if match is None:
        return "", ""
    authority = re.sub(r"^印发机关[：:]\s*", "", match.group("authority").strip())
    authority = authority.rstrip("：:，,；;").strip()
    return authority, convert_date_format(match.group("date"))


def remove_brackets(text: str) -> str:
    """Remove surrounding brackets (（）、[]、()、《》 etc.) from text."""
    brackets = [
        ("（", "）"),
        ("(", ")"),
        ("【", "】"),
        ("[", "]"),
        ("《", "》"),
        ("「", "」"),
        ("『", "』"),
    ]
    t = text.strip()
    for left, right in brackets:
        if t.startswith(left) and t.endswith(right):
            return t[len(left) : -len(right)].strip()
    return t


def remove_colon(text: str) -> str:
    """Remove trailing colon from text."""
    return text.rstrip("：:").strip()


def convert_to_halfwidth(text: str) -> str:
    """Convert fullwidth characters to halfwidth where applicable."""
    result = []
    for ch in text:
        code = ord(ch)
        if 0xFF01 <= code <= 0xFF5E:
            result.append(chr(code - 0xFEE0))
        elif code == 0x3000:
            result.append(" ")
        else:
            result.append(ch)
    return "".join(result)


def process_attachment_item(text: str, *, cleanup_rules=()) -> str:
    """Clean up an attachment item description."""
    from docwen_core.text.heading_numbering import strip_heading_prefix

    t = re.sub(r"^附件[：:]\s*", "", text.strip())
    _, t = strip_heading_prefix(t, rules=cleanup_rules)
    t = re.sub(r"^\d+[.．]\s*", "", t)
    t = re.sub(r"^[（(]?[一二三四五六七八九十]+[）)、.．、]\s*", "", t)
    t = re.sub(r"[。.]$", "", t)
    return t.strip()


def process_copy_to(text: str) -> list[str]:
    """Split CC text into organization list.

    '抄送：省委组织部、省人社厅。' → ['省委组织部', '省人社厅']
    """
    t = _COPY_TO_LABEL_RE.sub("", text.strip(), count=1)
    # Strip trailing period
    t = re.sub(r"[。.]\s*$", "", t)
    # Split by separators
    parts = re.split(r"[，,、]", t)
    return [p.strip() for p in parts if p.strip()]


def starts_with_copy_to_label(text: str) -> bool:
    """Return whether text starts with a supported copy-distribution label."""

    return bool(_COPY_TO_LABEL_RE.match(text.strip()))


def format_yaml_value(value) -> str:
    """Format a value as a YAML-safe string using PyYAML.

    Handles special characters ([ ] { } : # & * ! | > ' " @ ` % ? -),
    pure numbers, booleans, null literals, and quote mixing.
    """
    import yaml

    if value is None or value == "":
        return ""

    str_value = str(value)
    result = yaml.safe_dump(str_value, default_flow_style=True, allow_unicode=True)
    result = result.strip()
    if result.endswith("..."):
        result = result[:-3].strip()
    return result


def format_display_value(value, separator: str = "") -> str:
    """Format a value for display, flattening nested lists.

    Skips None/empty values and joins with the given separator.
    """
    if value is None:
        return ""

    if isinstance(value, list):

        def flatten(lst):
            result: list[str] = []
            for item in lst:
                if isinstance(item, list):
                    result.extend(flatten(item))
                elif item not in [None, "", "null", "None"]:
                    result.append(str(item))
            return result

        non_empty = flatten(value)
        return separator.join(non_empty)

    return str(value)


# ── 签发人提取 ──────────────────────────────────────────────

_NAME_PART_PATTERN = r"[一-龥]{2,12}"
_NAME_PATTERN = (
    rf"{_NAME_PART_PATTERN}"
    rf"(?:[·•]{_NAME_PART_PATTERN})*"
)
_NAME_RE = re.compile(rf"^{_NAME_PATTERN}$")
_NAME_LIST_SPLIT_RE = re.compile(r"[、\s　]+")


def extract_signers_from_text(text: str) -> list[str]:
    """从纯人名文本中提取多个签发人。

    支持的格式：
    - "李四、王五" — 顿号分隔
    - "李四 王五" — 空格分隔
    - "张三　李四" — 全角空格分隔

    每段按 NAME_PATTERN 验证（2-12 个中文字符，
    支持少数民族间隔号），不符合格式的项被过滤。
    """
    text = text.strip()
    parts = _NAME_LIST_SPLIT_RE.split(text)
    tokens = [p.strip() for p in parts if p.strip()]
    return [t for t in tokens if _NAME_RE.match(t)]

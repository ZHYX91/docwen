"""Conservatively collect a clearly continued recipient block without splitting agencies."""

from __future__ import annotations

import re
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from docwen_plugin_optimizer_gongwen.models import ParagraphFeature, RecognitionResult

_CONTINUED_AUTHORITY = re.compile(
    r"(?:人民政府|委员会|办公室|各单位|各部门|各企业|各学校|各院校|各公司|局|厅|部|院|校|中心|公司|集团|协会)[、，,；;]$"
)
_PROSE_START = re.compile(r"^(?:请|现|为|根据|按照|关于|经|同意|通知|要求|以下|上述|本次|会议)")
_COLLECTIVE_ADDRESS = re.compile(
    r"(?:各|[省市县区]直各)[\u3400-\u9fff、，,；;（）()\s]*"
    r"(?:人民政府|委员会|办公室|单位|部门|企业|学校|院校|公司|局|厅|部|院|校|中心|集团|协会)[、，,；;]$"
)
_RECIPIENT_LABEL = re.compile(r"^主送机关\s*[：:]")
_RECIPIENT_NAME = re.compile(r"[\u3400-\u9fff][\u3400-\u9fffA-Za-z0-9、，,；;（）()\s]*")
_RECIPIENT_SUFFIX = re.compile(
    r"(?:政府|委员会|办公室|单位|部门|企业|学校|公司|集团|协会|学会|研究所|事务所|银行|"
    r"社区|组织|总队|支队|各地|局|厅|部|院|校|中心|委|办|组|室|司|署|处|科|省|市|县|区|州|盟|旗)$"
)
_RECIPIENT_PROSE = re.compile(
    r"^(?:现将|现就|根据|按照|请各)|(?:如下|下列)$|"
    r"(?:单位|部门|机关)(?:应(?!急)|须|将|已|由|现请)"
)
_PROSE_CONTENT = re.compile(
    r"负责|联系|办理|承办|通知|要求|规定|会议|(?:单位|部门|机关|工作)(?:应|须|要|将|已|由|为|现|请)"
)


def is_recipient_address(text: str) -> bool:
    """Accept an explicit label or an agency expression ending in a colon."""
    text = text.strip()
    label = _RECIPIENT_LABEL.match(text)
    if label is not None:
        return bool(text[label.end() :].strip().strip("：:").strip())
    if not text.endswith(("：", ":")):
        return False
    address = text[:-1].rstrip()
    return (
        _RECIPIENT_NAME.fullmatch(address) is not None
        and _RECIPIENT_SUFFIX.search(address.rstrip("）) ")) is not None
        and _RECIPIENT_PROSE.search(address) is None
        and _balanced_brackets(address)
    )


def _clear_continuation(text: str) -> bool:
    return (
        _COLLECTIVE_ADDRESS.fullmatch(text) is not None
        and _PROSE_CONTENT.search(text) is None
        and _balanced_brackets(text)
    )


def _balanced_brackets(text: str) -> bool:
    brackets: list[str] = []
    for character in text:
        if character in "（(":
            brackets.append(character)
        elif character in "）)" and (not brackets or (brackets.pop(), character) not in {("（", "）"), ("(", ")")}):
            return False
    return not brackets


def recipient_block(features: list[ParagraphFeature], result: RecognitionResult, index: int) -> tuple[str, list[int]]:
    """Join only adjacent, same-container agency continuations ending in punctuation.

    The already-recognized final paragraph is the anchor. Uncertain earlier
    prose remains in the body; this does not reclassify arbitrary text or
    interpret punctuation inside parentheses as agency boundaries.
    """
    parts = [features[index].text.strip()]
    consumed: list[int] = []
    current = features[index]
    for previous in reversed(features[:index]):
        candidate = result.candidates.get(previous.index)
        if candidate is not None and candidate.element_type != "body":
            break
        if previous.source != "body" or current.source != "body":
            break
        if (
            previous.source_index is None
            or current.source_index is None
            or previous.source_index + 1 != current.source_index
        ):
            break
        if (
            previous.has_image
            or previous.has_formula
            or previous.has_page_break
            or previous.has_section_break
            or previous.heading_level
            or previous.list_marker
        ):
            break
        text = previous.text.strip()
        if (
            not text
            or len(text) > 200
            or _PROSE_START.match(text)
            or any(character in text for character in "。！？!?")
            or not _CONTINUED_AUTHORITY.search(text)
        ):
            break
        if not _clear_continuation(text):
            # A suffix is only suspicion, not permission to remove body text.
            anchor = result.candidates.get(index)
            if anchor is not None and "recipient_continuation_uncertain" not in anchor.trace:
                anchor.trace.append("recipient_continuation_uncertain")
            break
        parts.insert(0, text)
        consumed.append(previous.index)
        current = previous
    return "".join(parts), consumed

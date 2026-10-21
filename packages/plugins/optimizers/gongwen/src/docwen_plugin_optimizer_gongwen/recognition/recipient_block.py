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
_PROSE_CONTENT = re.compile(
    r"负责|联系|办理|承办|通知|要求|规定|会议|(?:单位|部门|机关|工作)(?:应|须|要|将|已|由|为|现|请)"
)


def _clear_continuation(text: str) -> bool:
    if _COLLECTIVE_ADDRESS.fullmatch(text) is None or _PROSE_CONTENT.search(text):
        return False
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

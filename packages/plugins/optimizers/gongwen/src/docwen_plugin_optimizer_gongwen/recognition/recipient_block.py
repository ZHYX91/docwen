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
_COLLECTIVE_PREFIX = re.compile(r"^(?:各|[省市县区]直各)")
_RECIPIENT_NAME = re.compile(r"\w[\w、，,；;（）()·•＆& \t　-]*")
_CONTINUATION_SEPARATORS = ("、", "，", ",", "；", ";")
_RECIPIENT_PROSE = re.compile(
    r"^(?:现将|现就|根据|按照|请各)|(?:如下|下列)$|"
    r"(?:单位|部门|机关)(?:应(?!急)|须|将|已|由|现请)"
)
_PROSE_CONTENT = re.compile(
    r"负责|联系|办理|承办|通知|要求|规定|会议|(?:单位|部门|机关|工作)(?:应|须|要|将|已|由|为|现|请)"
)


def is_recipient_address(text: str) -> bool:
    """Recognize the text shape of a standalone address ending in a colon."""
    text = text.strip()
    if not text.endswith(("：", ":")):
        return False
    address = text[:-1].rstrip()
    return (
        _RECIPIENT_NAME.fullmatch(address) is not None
        and _RECIPIENT_PROSE.search(address) is None
        and _balanced_brackets(address)
    )


def has_recipient_list_separator(text: str) -> bool:
    """Find a separator between addressees, outside any name qualifiers."""
    depth = 0
    for character in text:
        if character in "（(":
            depth += 1
        elif character in "）)":
            depth -= 1
        elif depth == 0 and character in _CONTINUATION_SEPARATORS:
            return True
    return False


def could_continue_recipient(text: str) -> bool:
    """Recognize a short opening fragment that ends in a separator."""
    return (
        len(text) <= 200
        and text.endswith(_CONTINUATION_SEPARATORS)
        and not any(character in text for character in "。！？!?")
    )


def _clear_continuation(text: str) -> bool:
    return (
        _COLLECTIVE_PREFIX.match(text) is not None
        and _RECIPIENT_NAME.fullmatch(text[:-1].rstrip()) is not None
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
            or previous.has_first_line_indent
            or previous.alignment != current.alignment
        ):
            break
        text = previous.text.strip()
        if not could_continue_recipient(text) or _PROSE_START.match(text):
            break
        if not _clear_continuation(text):
            # A suffix is only suspicion, not permission to remove body text.
            anchor = result.candidates.get(index)
            if (
                _CONTINUED_AUTHORITY.search(text)
                and anchor is not None
                and "recipient_continuation_uncertain" not in anchor.trace
            ):
                anchor.trace.append("recipient_continuation_uncertain")
            break
        parts.insert(0, text)
        consumed.append(previous.index)
        current = previous
    return "".join(parts), consumed

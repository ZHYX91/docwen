"""Tests for image-OCR Markdown presentation helpers."""

from __future__ import annotations

import pytest

from docwen_plugin_image.to_markdown.ocr_text import split_ocr_heading_body

pytestmark = pytest.mark.unit


@pytest.mark.parametrize(
    ("line", "expected"),
    [
        ("一、标题：正文内容", ("一、标题：", "正文内容")),
        ("Title: body text", ("Title:", "body text")),
        ("摘要。详细内容", ("摘要。", "详细内容")),
        ("标题．正文", ("标题．", "正文")),
        ("注意！重要事项", ("注意！", "重要事项")),
        ("Warning! danger ahead", ("Warning!", "danger ahead")),
        ("Plain OCR text", ("Plain OCR text", "")),
        ("：正文", ("：", "正文")),
        ("标题：", ("标题：", "")),
        ("  Title  :   body  ", ("Title  :", "body")),
        ("", ("", "")),
    ],
)
def test_split_ocr_heading_body(line: str, expected: tuple[str, str]) -> None:
    assert split_ocr_heading_body(line) == expected


def test_delimiter_type_priority_is_compatibility_preserving() -> None:
    assert split_ocr_heading_body("first: body：tail") == ("first: body：", "tail")

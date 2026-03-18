"""md2docx md_processor 的单元测试。"""

from __future__ import annotations

import pytest

from docwen.converter.md2docx.processors.md_processor import process_md_body

pytestmark = pytest.mark.unit


def test_md_processor_process_md_body_normalizes_crlf() -> None:
    md = "# 标题\r\n正文\r\n- 列表1\r\n"
    items = process_md_body(md)
    assert items
    assert all("text" in it and "type" in it and "level" in it for it in items)


def test_md_processor_heading_merge_mode_punct_required_merges_on_punct() -> None:
    md = "## 一、工作要求。\n本次会议要求各单位认真落实...\n"
    items = process_md_body(md, heading_merge_mode="punct_required")
    assert items == [
        {
            "text": "一、工作要求。",
            "level": 2,
            "type": "heading_with_content",
            "content": "本次会议要求各单位认真落实...",
        }
    ]


def test_md_processor_heading_merge_mode_punct_required_does_not_merge_without_punct() -> None:
    md = "## 一、工作要求\n本次会议要求各单位认真落实...\n"
    items = process_md_body(md, heading_merge_mode="punct_required")
    assert [it["type"] for it in items] == ["heading", "content"]


def test_md_processor_heading_merge_mode_always_merges_without_punct() -> None:
    md = "## 一、工作要求\n本次会议要求各单位认真落实...\n"
    items = process_md_body(md, heading_merge_mode="always")
    assert items == [
        {
            "text": "一、工作要求",
            "level": 2,
            "type": "heading_with_content",
            "content": "本次会议要求各单位认真落实...",
        }
    ]


def test_md_processor_heading_merge_mode_never_does_not_merge_with_punct() -> None:
    md = "## 一、工作要求。\n本次会议要求各单位认真落实...\n"
    items = process_md_body(md, heading_merge_mode="never")
    assert [it["type"] for it in items[:2]] == ["heading", "content"]

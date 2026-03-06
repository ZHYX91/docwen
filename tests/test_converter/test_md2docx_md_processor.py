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

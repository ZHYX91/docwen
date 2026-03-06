"""converter 单元测试。"""

from __future__ import annotations

from pathlib import Path

import pytest

from docwen.converter.md2docx.core import _read_and_parse_md as read_md_for_docx
from docwen.converter.md2xlsx.core import _read_and_parse_md as read_md_for_xlsx

pytestmark = pytest.mark.unit


def test_md2docx_read_and_parse_md_parses_crlf_yaml(tmp_path: Path) -> None:
    md_path = tmp_path / "in.md"
    md_path.write_text("---\r\ntitle: 测试\r\n---\r\n# 标题\r\n正文", encoding="utf-8")

    yaml_data, md_body = read_md_for_docx(str(md_path), original_md_path=str(md_path))
    assert yaml_data.get("title") == "测试"
    assert md_body.lstrip().startswith("# 标题")
    assert "正文" in md_body


def test_md2xlsx_read_and_parse_md_parses_crlf_yaml(tmp_path: Path) -> None:
    md_path = tmp_path / "in.md"
    md_path.write_text("---\r\ntitle: 测试\r\n---\r\n# 标题\r\n正文", encoding="utf-8")

    yaml_data, md_body = read_md_for_xlsx(str(md_path), original_md_path=str(md_path))
    assert yaml_data.get("title") == "测试"
    assert md_body.lstrip().startswith("# 标题")
    assert "正文" in md_body

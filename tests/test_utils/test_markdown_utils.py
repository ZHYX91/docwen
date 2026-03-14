"""utils 单元测试。"""

from __future__ import annotations

import pytest

from docwen.utils.markdown_utils import clean_heading, extract_yaml

pytestmark = pytest.mark.unit


@pytest.mark.unit
def test_extract_yaml_with_front_matter() -> None:
    content = "---\ntitle: 测试文档\ndate: 2023-01-15\n---\n# 标题内容\n正文内容"
    yaml_part, md_part = extract_yaml(content)
    assert yaml_part == "title: 测试文档\ndate: 2023-01-15"
    assert md_part == "# 标题内容\n正文内容"


@pytest.mark.unit
def test_extract_yaml_without_front_matter() -> None:
    yaml_part, md_part = extract_yaml("# 标题\n正文")
    assert yaml_part == ""
    assert md_part == "# 标题\n正文"


@pytest.mark.unit
def test_extract_yaml_with_front_matter_crlf() -> None:
    content = "---\r\ntitle: 测试文档\r\ndate: 2023-01-15\r\n---\r\n# 标题内容\r\n正文内容"
    yaml_part, md_part = extract_yaml(content)
    assert yaml_part == "title: 测试文档\r\ndate: 2023-01-15"
    assert md_part == "# 标题内容\r\n正文内容"


@pytest.mark.unit
def test_extract_yaml_with_utf8_bom() -> None:
    content = "\ufeff---\ntitle: 测试文档\n---\n# 标题内容\n正文内容"
    yaml_part, md_part = extract_yaml(content)
    assert yaml_part == "title: 测试文档"
    assert md_part == "# 标题内容\n正文内容"


@pytest.mark.unit
def test_extract_yaml_does_not_match_horizontal_rule() -> None:
    content = "# 标题\n\n---\n\n正文"
    yaml_part, md_part = extract_yaml(content)
    assert yaml_part == ""
    assert md_part == "# 标题\n\n---\n\n正文"


@pytest.mark.unit
def test_extract_yaml_allows_closing_marker_at_eof() -> None:
    content = "---\ntitle: 测试文档\n---"
    yaml_part, md_part = extract_yaml(content)
    assert yaml_part == "title: 测试文档"
    assert md_part == ""


@pytest.mark.unit
def test_extract_yaml_requires_front_matter_at_start() -> None:
    content = "# 前言\n\n---\ntitle: 不应匹配\n---\n\n正文"
    yaml_part, md_part = extract_yaml(content)
    assert yaml_part == ""
    assert md_part == content


@pytest.mark.unit
@pytest.mark.parametrize(
    ("raw", "expected"),
    [
        ("1. 测试标题", "测试标题"),
        ("（一）二级标题", "二级标题"),
        ("① 五级标题", "五级标题"),
    ],
)
def test_clean_heading(raw: str, expected: str) -> None:
    assert clean_heading(raw) == expected

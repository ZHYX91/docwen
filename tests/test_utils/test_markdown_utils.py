from __future__ import annotations

import pytest

from docwen.utils.markdown_utils import clean_heading, extract_yaml


@pytest.mark.unit
def test_extract_yaml_with_front_matter() -> None:
    content = "---\n" "title: 测试文档\n" "date: 2023-01-15\n" "---\n" "# 标题内容\n" "正文内容"
    yaml_part, md_part = extract_yaml(content)
    assert yaml_part == "title: 测试文档\ndate: 2023-01-15"
    assert md_part == "# 标题内容\n正文内容"


@pytest.mark.unit
def test_extract_yaml_without_front_matter() -> None:
    yaml_part, md_part = extract_yaml("# 标题\n正文")
    assert yaml_part == ""
    assert md_part == "# 标题\n正文"


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


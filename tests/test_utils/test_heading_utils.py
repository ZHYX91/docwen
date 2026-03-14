"""utils 单元测试。"""

from __future__ import annotations

import pytest

from docwen.utils.heading_utils import convert_to_halfwidth, detect_heading_level, split_content_by_delimiters

pytestmark = pytest.mark.unit


@pytest.mark.unit
def test_detect_heading_level_fullwidth_digits() -> None:
    cleaned, level = detect_heading_level("１．全角数字标题")
    assert cleaned == "全角数字标题"
    assert level == 3


@pytest.mark.unit
def test_detect_heading_level_fullwidth_brackets_digits() -> None:
    cleaned, level = detect_heading_level("（１）全角括号标题")
    assert cleaned == "全角括号标题"
    assert level == 4


@pytest.mark.unit
def test_split_content_by_delimiters() -> None:
    part1, part2 = split_content_by_delimiters("一级标题：正文内容")
    assert part1 == "一级标题："
    assert part2 == "正文内容"


@pytest.mark.unit
def test_convert_to_halfwidth() -> None:
    assert convert_to_halfwidth("１2３") == "123"

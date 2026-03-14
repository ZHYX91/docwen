"""utils 单元测试。"""

from __future__ import annotations

import pytest

from docwen.i18n import t
from docwen.utils.text_utils import format_yaml_value
from docwen.utils.yaml_utils import generate_basic_yaml_frontmatter, parse_yaml, process_list_field

pytestmark = pytest.mark.unit


@pytest.mark.unit
def test_parse_yaml_basic() -> None:
    raw = '标题: 关于[[2023年]]重点工作安排\n负责人:\n  - "[[张三]]"\n  - "[[李四]]"\n金额: 1234.56\n'
    data = parse_yaml(raw)
    assert data["标题"] == "关于[[2023年]]重点工作安排"
    assert data["负责人"] == ["[[张三]]", "[[李四]]"]
    assert data["金额"] == 1234.56


@pytest.mark.unit
@pytest.mark.parametrize(
    ("value", "expected"),
    [
        (None, ""),
        ("x", "x"),
        (123, "123"),
        (["[[张三]]", "[[李四]]"], "[[张三]]、[[李四]]"),
        ("['a', 'b']", "a、b"),
    ],
)
def test_process_list_field(value, expected: str) -> None:
    assert process_list_field(value) == expected


@pytest.mark.unit
def test_generate_basic_yaml_frontmatter_contains_title_key() -> None:
    file_stem = "[2024]年报"
    front = generate_basic_yaml_frontmatter(file_stem)

    safe_value = format_yaml_value(file_stem)
    assert front.startswith("---\n")
    assert "aliases:\n" in front
    assert f"  - {safe_value}\n" in front
    assert f"{t('yaml_keys.title')}: {safe_value}\n" in front
    assert front.endswith("---\n")

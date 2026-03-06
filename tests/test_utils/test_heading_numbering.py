"""
heading_numbering 单元测试

覆盖 HeadingFormatter 核心逻辑和 remove_numbering_from_md。
注意：add_numbering_to_md 和 process_md_numbering 依赖 config_manager，
在此仅测试不涉及 config_manager 的部分。
"""

from __future__ import annotations

import pytest

from docwen.utils.heading_numbering import HeadingFormatter, remove_numbering_from_md

pytestmark = pytest.mark.unit


# ============================================================
# HeadingFormatter 初始化
# ============================================================


class TestHeadingFormatterInit:
    """初始化与参数校验"""

    def test_valid_config(self) -> None:
        config = {"level_1": {"format": "{1.chinese_lower}、"}}
        fmt = HeadingFormatter(config)
        assert fmt.counters == [0] * 9

    def test_invalid_config_type_raises(self) -> None:
        with pytest.raises(ValueError, match="配置格式错误"):
            HeadingFormatter("not a dict")

    def test_empty_config(self) -> None:
        """空配置不报错，所有级别返回空字符串"""
        fmt = HeadingFormatter({})
        fmt.increment_level(1)
        assert fmt.format_heading(1) == ""


# ============================================================
# HeadingFormatter 计数器
# ============================================================


class TestHeadingFormatterCounters:
    """计数器递增与重置"""

    def _make_formatter(self) -> HeadingFormatter:
        return HeadingFormatter(
            {
                "level_1": {"format": "{1.arabic_half}."},
                "level_2": {"format": "{1.arabic_half}.{2.arabic_half}"},
                "level_3": {"format": "{1.arabic_half}.{2.arabic_half}.{3.arabic_half}"},
            }
        )

    def test_increment_and_format(self) -> None:
        fmt = self._make_formatter()
        fmt.increment_level(1)
        assert fmt.format_heading(1) == "1."
        fmt.increment_level(1)
        assert fmt.format_heading(1) == "2."

    def test_sub_level_reset(self) -> None:
        """递增上级时，下级计数器重置"""
        fmt = self._make_formatter()
        fmt.increment_level(1)
        fmt.increment_level(2)
        fmt.increment_level(2)
        assert fmt.format_heading(2) == "1.2"

        # 递增一级 → 二级重置
        fmt.increment_level(1)
        fmt.increment_level(2)
        assert fmt.format_heading(2) == "2.1"

    def test_cross_level_reference(self) -> None:
        """三级标题引用一、二级计数器"""
        fmt = self._make_formatter()
        fmt.increment_level(1)
        fmt.increment_level(2)
        fmt.increment_level(3)
        assert fmt.format_heading(3) == "1.1.1"

    def test_reset_counters(self) -> None:
        fmt = self._make_formatter()
        fmt.increment_level(1)
        fmt.increment_level(2)
        fmt.reset_counters()
        assert fmt.counters == [0] * 9

    def test_invalid_level_ignored(self) -> None:
        fmt = self._make_formatter()
        fmt.increment_level(0)  # 无效
        fmt.increment_level(10)  # 无效
        assert fmt.counters == [0] * 9


# ============================================================
# HeadingFormatter 数字样式
# ============================================================


class TestHeadingFormatterStyles:
    """不同数字样式的模板解析"""

    @pytest.mark.parametrize(
        ("template", "expected"),
        [
            ("{1.chinese_lower}、", "一、"),
            ("{1.chinese_upper}、", "壹、"),
            ("{1.arabic_half}.", "1."),
            ("{1.arabic_full}", "１"),
            ("{1.arabic_circled}", "①"),
            ("{1.letter_upper}.", "A."),
            ("{1.letter_lower})", "a)"),
            ("{1.roman_upper}.", "I."),
            ("{1.roman_lower}.", "i."),
        ],
    )
    def test_style(self, template: str, expected: str) -> None:
        fmt = HeadingFormatter({"level_1": {"format": template}})
        fmt.increment_level(1)
        assert fmt.format_heading(1) == expected

    def test_unknown_style_falls_back_to_arabic(self) -> None:
        fmt = HeadingFormatter({"level_1": {"format": "{1.unknown_style}"}})
        fmt.increment_level(1)
        assert fmt.format_heading(1) == "1"


# ============================================================
# HeadingFormatter 格式化边界
# ============================================================


class TestHeadingFormatterEdgeCases:
    """边界情况"""

    def test_format_heading_invalid_level(self) -> None:
        fmt = HeadingFormatter({"level_1": {"format": "{1.arabic_half}"}})
        assert fmt.format_heading(0) == ""
        assert fmt.format_heading(10) == ""

    def test_format_heading_no_template(self) -> None:
        """级别存在但没有 format 字段"""
        fmt = HeadingFormatter({"level_1": {}})
        fmt.increment_level(1)
        assert fmt.format_heading(1) == ""

    def test_reference_unvisited_level_defaults_to_one(self) -> None:
        """引用计数器为 0 的级别，默认使用 1"""
        fmt = HeadingFormatter({"level_2": {"format": "{1.arabic_half}.{2.arabic_half}"}})
        fmt.increment_level(2)  # 一级未递增，计数器为 0
        # 一级默认为 1，二级为 1
        assert fmt.format_heading(2) == "1.1"

    def test_get_scheme_info(self) -> None:
        config = {
            "name": "测试方案",
            "description": "测试用",
            "level_1": {"format": "{1.arabic_half}"},
            "level_2": {"format": "{2.arabic_half}"},
        }
        fmt = HeadingFormatter(config)
        info = fmt.get_scheme_info()
        assert info["name"] == "测试方案"
        assert info["description"] == "测试用"
        assert info["levels_defined"] == 2

    def test_get_scheme_info_defaults(self) -> None:
        fmt = HeadingFormatter({})
        info = fmt.get_scheme_info()
        assert info["name"] == "未命名方案"
        assert info["description"] == ""
        assert info["levels_defined"] == 0


# ============================================================
# remove_numbering_from_md
# ============================================================


class TestRemoveNumberingFromMd:
    """Markdown 标题序号去除"""

    def test_removes_chinese_numbering(self) -> None:
        content = "# 一、标题内容\n正文"
        result = remove_numbering_from_md(content)
        assert "# " in result
        assert "一、" not in result
        assert "标题内容" in result
        assert "正文" in result

    def test_preserves_non_heading_lines(self) -> None:
        content = "普通文本\n\n## 1.子标题\n\n更多文本"
        result = remove_numbering_from_md(content)
        assert "普通文本" in result
        assert "更多文本" in result

    def test_skips_code_blocks(self) -> None:
        content = "```\n# 一、代码中的标题\n```\n# 一、真正的标题"
        result = remove_numbering_from_md(content)
        # 代码块内保持不变
        assert "# 一、代码中的标题" in result

    def test_empty_content(self) -> None:
        assert remove_numbering_from_md("") == ""

    def test_no_headings(self) -> None:
        content = "只有正文\n没有标题"
        assert remove_numbering_from_md(content) == content

    def test_multiple_levels(self) -> None:
        content = "# 一、一级\n## （一）二级\n### 1.三级"
        result = remove_numbering_from_md(content)
        assert "一、" not in result
        # 确保标题文本保留
        lines = result.split("\n")
        assert lines[0].startswith("# ")
        assert lines[1].startswith("## ")
        assert lines[2].startswith("### ")

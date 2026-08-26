"""Tests for docwen_core.text.heading_numbering.

Covers prefix detection, stripping, ordering of immutable request rules,
edge cases, and HeadingFormatter functionality.
"""

from __future__ import annotations

import re
from types import SimpleNamespace

import pytest

from docwen_core.text.heading_numbering import (
    HeadingFormatter,
    NumberingSchemeResolutionError,
    compile_clean_rules_from_data,
    resolve_heading_numbering_scheme,
)
from docwen_core.text.heading_numbering import (
    detect_heading_prefix as _detect_heading_prefix,
)
from docwen_core.text.heading_numbering import (
    strip_heading_prefix as _strip_heading_prefix,
)

pytestmark = pytest.mark.unit

# Test rules data — mirrors the old built-in fallback rules for detection tests.
# These must be injected explicitly since the module no longer carries any
# built-in fallback rules.
_FALLBACK_RULES_DATA: list[dict] = [
    {
        "id": "legal_unit",
        "enabled": True,
        "pattern": r"^[\s ]*第[\s ]*[一二三四五六七八九十百千万\d]+[\s ]*[编章节条回][\s ]+",
        "level": 1,
    },
    {"id": "chinese_顿号", "enabled": True, "pattern": r"^[一二三四五六七八九十百千万]+、", "level": 1},
    {"id": "chinese_bracket", "enabled": True, "pattern": r"^[（(][一二三四五六七八九十百千万]+[)）]", "level": 2},
    {
        "id": "arabic_separator",
        "enabled": True,
        "pattern": r"^[0-9\uFF10-\uFF19]+[\.．、，,。](?![0-9\uFF10-\uFF19])",
        "level": 3,
    },
    {"id": "arabic_bracket", "enabled": True, "pattern": r"^[（(][0-9\uFF10-\uFF19]+[)）]", "level": 4},
    {
        "id": "circled",
        "enabled": True,
        "pattern": r"^[\u2460-\u32BF\u2776-\u277F\u24EB-\u24F4\u24F5-\u24FE\u3220-\u3229]",
        "level": 5,
    },
    {
        "id": "hierarchical",
        "enabled": True,
        "pattern": r"^[0-9\uFF10-\uFF19]+(?:[\.．][0-9\uFF10-\uFF19]+)+[\.．\s ]*",
        "level": 3,
    },
]


_STANDARD_RULES = compile_clean_rules_from_data(_FALLBACK_RULES_DATA)


def detect_heading_prefix(text: str, *, rules=_STANDARD_RULES):
    """Exercise the production API while keeping rule fixtures concise."""
    return _detect_heading_prefix(text, rules=rules)


def strip_heading_prefix(text: str, *, rules=_STANDARD_RULES):
    """Exercise the production API while keeping rule fixtures concise."""
    return _strip_heading_prefix(text, rules=rules)


class TestRuleCompilation:
    """Verify compilation and explicit request ownership."""

    def test_detection_requires_explicit_rules(self) -> None:
        with pytest.raises(TypeError, match="rules"):
            _detect_heading_prefix("test_heading")  # type: ignore[call-arg]

    def test_strip_requires_explicit_rules(self) -> None:
        with pytest.raises(TypeError, match="rules"):
            _strip_heading_prefix("test_heading")  # type: ignore[call-arg]

    def test_compile_rules_filters_disabled(self) -> None:
        data = [
            {"id": "r1", "enabled": True, "pattern": r"^hello", "level": 1},
            {"id": "r2", "enabled": False, "pattern": r"^world", "level": 2},
            {"id": "r3", "enabled": True, "pattern": r"^foo", "level": 3},
        ]
        rules = compile_clean_rules_from_data(data)
        assert len(rules) == 2
        ids = {r[0] for r in rules}
        assert "r1" in ids
        assert "r2" not in ids
        assert "r3" in ids

    def test_compile_rules_skips_invalid_regex(self) -> None:
        data = [
            {"id": "good", "enabled": True, "pattern": r"^valid", "level": 1},
            {"id": "bad", "enabled": True, "pattern": r"[invalid", "level": 2},
            {"id": "good2", "enabled": True, "pattern": r"^also_valid", "level": 3},
        ]
        rules = compile_clean_rules_from_data(data)
        assert len(rules) == 2
        ids = {r[0] for r in rules}
        assert "good" in ids
        assert "bad" not in ids
        assert "good2" in ids

    def test_compile_rules_preserves_order(self) -> None:
        data = [
            {"id": "third", "enabled": True, "pattern": r"^c", "level": 1},
            {"id": "first", "enabled": True, "pattern": r"^a", "level": 1},
            {"id": "second", "enabled": True, "pattern": r"^b", "level": 1},
        ]
        rules = compile_clean_rules_from_data(data)
        assert [r[0] for r in rules] == ["third", "first", "second"]

    def test_compile_rules_preserves_level(self) -> None:
        data = [
            {"id": "r1", "enabled": True, "pattern": r"^x", "level": 5},
        ]
        rules = compile_clean_rules_from_data(data)
        assert rules[0][2] == 5

    def test_independent_rule_sets_do_not_share_state(self) -> None:
        global_rules = (("global", re.compile(r"^GLOBAL:\s*"), 1),)
        request_rules = compile_clean_rules_from_data(
            [{"id": "request", "enabled": True, "pattern": r"^REQ:\s*", "level": 2}]
        )

        request_info = detect_heading_prefix("REQ: Title", rules=request_rules)
        global_text_in_request = strip_heading_prefix("GLOBAL: Title", rules=request_rules)

        assert request_info is not None
        assert (request_info.rule_id, request_info.clean_text, request_info.numbering_level) == (
            "request",
            "Title",
            2,
        )
        assert global_text_in_request == ("", "GLOBAL: Title")
        assert strip_heading_prefix("GLOBAL: Title", rules=global_rules) == ("GLOBAL: ", "Title")

    def test_explicit_empty_rules_never_strip(self) -> None:
        assert strip_heading_prefix("GLOBAL: Title", rules=()) == ("", "GLOBAL: Title")


class TestSchemeResolution:
    class _Registry:
        def __init__(self, *, enabled: bool = True, levels: dict[str, str] | None = None) -> None:
            self._scheme = SimpleNamespace(
                enabled=enabled,
                levels={"level_1": "{1.arabic_half} "} if levels is None else levels,
            )

        def get_scheme(self, scheme_id: str) -> object:
            if scheme_id != "exact":
                raise LookupError(scheme_id)
            return self._scheme

    def test_resolves_only_the_exact_enabled_scheme(self) -> None:
        config = resolve_heading_numbering_scheme("exact", self._Registry())

        assert config == {"level_1": {"format": "{1.arabic_half} "}}

    @pytest.mark.parametrize(
        ("scheme_id", "registry", "error_type", "diagnostic_code"),
        [
            ("", _Registry(), "invalid_input", "NUMBERING-SCHEME-REQUIRED"),
            ("exact", None, "capability_unavailable", "NUMBERING-REGISTRY-UNAVAILABLE"),
            ("missing", _Registry(), "resource_not_found", "NUMBERING-SCHEME-NOT-FOUND"),
            ("exact", _Registry(enabled=False), "capability_unavailable", "NUMBERING-SCHEME-DISABLED"),
            ("exact", _Registry(levels={}), "invalid_input", "NUMBERING-SCHEME-NO-LEVELS"),
        ],
    )
    def test_rejects_every_unusable_requested_scheme(
        self,
        scheme_id: str,
        registry: object,
        error_type: str,
        diagnostic_code: str,
    ) -> None:
        with pytest.raises(NumberingSchemeResolutionError) as caught:
            resolve_heading_numbering_scheme(scheme_id, registry)

        assert caught.value.error_type == error_type
        assert caught.value.diagnostic_code == diagnostic_code


# Individual rule tests - legal_unit


class TestLegalUnit:
    def test_chapter_with_fullwidth_space(self) -> None:
        info = detect_heading_prefix("第一章　总则")
        assert info is not None
        assert info.rule_id == "legal_unit"
        assert info.prefix == "第一章　"
        assert info.clean_text == "总则"
        assert info.numbering_level == 1

    def test_article_with_ascii_space(self) -> None:
        info = detect_heading_prefix("第一条 标题")
        assert info is not None
        assert info.rule_id == "legal_unit"
        assert info.prefix == "第一条 "
        assert info.clean_text == "标题"

    def test_section(self) -> None:
        info = detect_heading_prefix("第二节 内容")
        assert info is not None
        assert info.rule_id == "legal_unit"
        assert info.prefix == "第二节 "
        assert info.clean_text == "内容"
        assert info.numbering_level == 1

    def test_legal_unit_with_digits(self) -> None:
        info = detect_heading_prefix("第100条　规定")
        assert info is not None
        assert info.rule_id == "legal_unit"
        assert info.prefix == "第100条　"
        assert info.clean_text == "规定"

    def test_section_no_trailing_space_does_not_match(self) -> None:
        assert detect_heading_prefix("第一节") is None

    def test_legal_unit_with_spaces(self) -> None:
        info = detect_heading_prefix("第 一 章 标题")
        assert info is not None
        assert info.rule_id == "legal_unit"
        assert info.prefix == "第 一 章 "
        assert info.clean_text == "标题"


# Individual rule tests - chinese_dunhao


class TestChineseDunhao:
    def test_single_character(self) -> None:
        info = detect_heading_prefix("一、标题")
        assert info is not None
        assert info.rule_id == "chinese_顿号"
        assert info.prefix == "一、"
        assert info.clean_text == "标题"
        assert info.numbering_level == 1

    def test_double_digit(self) -> None:
        info = detect_heading_prefix("十、标题")
        assert info is not None
        assert info.rule_id == "chinese_顿号"
        assert info.prefix == "十、"
        assert info.clean_text == "标题"

    def test_twelve(self) -> None:
        info = detect_heading_prefix("十二、标题")
        assert info is not None
        assert info.rule_id == "chinese_顿号"
        assert info.prefix == "十二、"
        assert info.clean_text == "标题"

    def test_no_dunhao_does_not_match(self) -> None:
        assert detect_heading_prefix("一标题") is None

    def test_hundred(self) -> None:
        info = detect_heading_prefix("百、标题")
        assert info is not None
        assert info.rule_id == "chinese_顿号"
        assert info.prefix == "百、"

    def test_thousand(self) -> None:
        info = detect_heading_prefix("千、标题")
        assert info is not None
        assert info.rule_id == "chinese_顿号"
        assert info.prefix == "千、"


# Individual rule tests - chinese_bracket


class TestChineseBracket:
    def test_fullwidth_brackets(self) -> None:
        info = detect_heading_prefix("（一）标题")
        assert info is not None
        assert info.rule_id == "chinese_bracket"
        assert info.prefix == "（一）"
        assert info.clean_text == "标题"
        assert info.numbering_level == 2

    def test_halfwidth_brackets(self) -> None:
        info = detect_heading_prefix("(二)标题")
        assert info is not None
        assert info.rule_id == "chinese_bracket"
        assert info.prefix == "(二)"
        assert info.clean_text == "标题"

    def test_multi_char(self) -> None:
        info = detect_heading_prefix("（十二）标题")
        assert info is not None
        assert info.rule_id == "chinese_bracket"
        assert info.prefix == "（十二）"
        assert info.clean_text == "标题"


# Individual rule tests - arabic_separator


class TestArabicSeparator:
    def test_halfwidth_dot(self) -> None:
        info = detect_heading_prefix("1.标题")
        assert info is not None
        assert info.rule_id == "arabic_separator"
        assert info.prefix == "1."
        assert info.clean_text == "标题"
        assert info.numbering_level == 3

    def test_fullwidth_dunhao(self) -> None:
        info = detect_heading_prefix("１、标题")
        assert info is not None
        assert info.rule_id == "arabic_separator"
        assert info.prefix == "１、"
        assert info.clean_text == "标题"

    def test_chinese_comma(self) -> None:
        info = detect_heading_prefix("1，标题")
        assert info is not None
        assert info.rule_id == "arabic_separator"
        assert info.prefix == "1，"
        assert info.clean_text == "标题"

    def test_fullwidth_dot(self) -> None:
        info = detect_heading_prefix("１．标题")
        assert info is not None
        assert info.rule_id == "arabic_separator"
        assert info.prefix == "１．"
        assert info.clean_text == "标题"

    def test_fullwidth_comma(self) -> None:
        info = detect_heading_prefix("１２３、标题")
        assert info is not None
        assert info.rule_id == "arabic_separator"
        assert info.prefix == "１２３、"
        assert info.clean_text == "标题"

    def test_trailing_period(self) -> None:
        info = detect_heading_prefix("1.标题内容")
        assert info is not None
        assert info.prefix == "1."


# Individual rule tests - arabic_bracket


class TestArabicBracket:
    def test_fullwidth_brackets(self) -> None:
        info = detect_heading_prefix("（1）标题")
        assert info is not None
        assert info.rule_id == "arabic_bracket"
        assert info.prefix == "（1）"
        assert info.clean_text == "标题"
        assert info.numbering_level == 4

    def test_halfwidth_brackets(self) -> None:
        info = detect_heading_prefix("(2)标题")
        assert info is not None
        assert info.rule_id == "arabic_bracket"
        assert info.prefix == "(2)"
        assert info.clean_text == "标题"

    def test_multi_digit(self) -> None:
        info = detect_heading_prefix("（12）标题")
        assert info is not None
        assert info.rule_id == "arabic_bracket"
        assert info.prefix == "（12）"
        assert info.clean_text == "标题"


# Individual rule tests - circled


class TestCircled:
    def test_circled_1(self) -> None:
        info = detect_heading_prefix("①标题")
        assert info is not None
        assert info.rule_id == "circled"
        assert info.prefix == "①"
        assert info.clean_text == "标题"
        assert info.numbering_level == 5

    def test_circled_50(self) -> None:
        info = detect_heading_prefix("㊿标题")
        assert info is not None
        assert info.rule_id == "circled"
        assert info.prefix == "㊿"
        assert info.clean_text == "标题"

    def test_circled_20(self) -> None:
        info = detect_heading_prefix("⑳标题")
        assert info is not None
        assert info.rule_id == "circled"
        assert info.prefix == "⑳"


# Individual rule tests - hierarchical


class TestHierarchical:
    def test_two_level_halfwidth(self) -> None:
        info = detect_heading_prefix("1.1 标题")
        assert info is not None
        assert info.rule_id == "hierarchical"
        assert info.prefix == "1.1 "
        assert info.clean_text == "标题"
        assert info.numbering_level == 3

    def test_three_level(self) -> None:
        info = detect_heading_prefix("1.1.1.标题")
        assert info is not None
        assert info.rule_id == "hierarchical"
        assert info.prefix == "1.1.1."
        assert info.clean_text == "标题"

    def test_four_level_with_space(self) -> None:
        info = detect_heading_prefix("1.2.3.4 内容")
        assert info is not None
        assert info.rule_id == "hierarchical"
        assert info.prefix == "1.2.3.4 "
        assert info.clean_text == "内容"

    def test_fullwidth_digits(self) -> None:
        info = detect_heading_prefix("１．１ 标题")
        assert info is not None
        assert info.rule_id == "hierarchical"
        assert info.prefix == "１．１ "
        assert info.clean_text == "标题"

    def test_single_level_dot_matches_arabic_not_hierarchical(self) -> None:
        info = detect_heading_prefix("1.标题")
        assert info is not None
        assert info.rule_id == "arabic_separator"

    def test_just_numbers_dot(self) -> None:
        info = detect_heading_prefix("1.1.标题")
        assert info is not None
        assert info.rule_id == "hierarchical"
        assert info.prefix == "1.1."


# Edge cases


class TestEdgeCases:
    def test_no_prefix(self) -> None:
        info = detect_heading_prefix("普通标题文本")
        assert info is None

    def test_strip_no_prefix(self) -> None:
        prefix, clean = strip_heading_prefix("普通标题文本")
        assert prefix == ""
        assert clean == "普通标题文本"

    def test_empty_string(self) -> None:
        assert detect_heading_prefix("") is None

    def test_strip_empty_string(self) -> None:
        prefix, clean = strip_heading_prefix("")
        assert prefix == ""
        assert clean == ""

    def test_hundred_years_does_not_match(self) -> None:
        assert detect_heading_prefix("百年来变迁") is None

    def test_fullwidth_arabic_separator_digits(self) -> None:
        info = detect_heading_prefix("１２３、标题")
        assert info is not None
        assert info.rule_id == "arabic_separator"
        assert info.prefix == "１２３、"
        assert info.clean_text == "标题"

    def test_strip_heading_prefix_with_prefix(self) -> None:
        prefix, clean = strip_heading_prefix("一、简介")
        assert prefix == "一、"
        assert clean == "简介"

    def test_whitespace_prefix(self) -> None:
        info = detect_heading_prefix("  第一章　总则")
        assert info is not None
        assert info.rule_id == "legal_unit"
        assert info.prefix == "  第一章　"
        assert info.clean_text == "总则"


# HeadingFormatter tests


class TestHeadingFormatter:
    @staticmethod
    def _gongwen_config() -> dict:
        return {
            "level_1": {"format": "{1.chinese_lower}、"},
            "level_2": {"format": "（{2.chinese_lower}）"},
            "level_3": {"format": "{3.arabic_half}. "},
            "level_4": {"format": "（{4.arabic_half}）"},
            "level_5": {"format": "{5.arabic_circled}"},
        }

    @staticmethod
    def _hierarchical_config() -> dict:
        return {
            "level_1": {"format": "{1.arabic_half} "},
            "level_2": {"format": "{1.arabic_half}.{2.arabic_half} "},
            "level_3": {"format": "{1.arabic_half}.{2.arabic_half}.{3.arabic_half} "},
        }

    def test_gongwen_level1_sequential(self) -> None:
        fmt = HeadingFormatter(self._gongwen_config())
        assert fmt.format_heading("标题一", 1) == "一、标题一"
        assert fmt.format_heading("标题二", 1) == "二、标题二"
        assert fmt.format_heading("标题三", 1) == "三、标题三"

    def test_gongwen_level1_then_level2(self) -> None:
        fmt = HeadingFormatter(self._gongwen_config())
        assert fmt.format_heading("一级", 1) == "一、一级"
        assert fmt.format_heading("二级", 2) == "（一）二级"
        assert fmt.format_heading("二级二", 2) == "（二）二级二"

    def test_gongwen_level2_resets_when_level1_increments(self) -> None:
        fmt = HeadingFormatter(self._gongwen_config())
        assert fmt.format_heading("A", 1) == "一、A"
        assert fmt.format_heading("A1", 2) == "（一）A1"
        assert fmt.format_heading("A2", 2) == "（二）A2"
        assert fmt.format_heading("B", 1) == "二、B"
        assert fmt.format_heading("B1", 2) == "（一）B1"

    def test_gongwen_all_levels(self) -> None:
        fmt = HeadingFormatter(self._gongwen_config())
        r1 = fmt.format_heading("一级", 1)
        assert r1 == "一、一级"
        r2 = fmt.format_heading("二级", 2)
        assert r2 == "（一）二级"
        r3 = fmt.format_heading("三级", 3)
        assert r3 == "1. 三级"
        r4 = fmt.format_heading("四级", 4)
        assert r4 == "（1）四级"
        r5 = fmt.format_heading("五级", 5)
        assert r5 == "①五级"

    def test_gongwen_level_skips_template_not_defined(self) -> None:
        cfg = {"level_1": {"format": "{1.chinese_lower}、"}}
        fmt = HeadingFormatter(cfg)
        assert fmt.format_heading("标题", 1) == "一、标题"
        assert fmt.format_heading("二级", 2) == "二级"

    def test_hierarchical_level1(self) -> None:
        fmt = HeadingFormatter(self._hierarchical_config())
        assert fmt.format_heading("标题", 1) == "1 标题"
        assert fmt.format_heading("标题", 1) == "2 标题"

    def test_hierarchical_level1_then_level2(self) -> None:
        fmt = HeadingFormatter(self._hierarchical_config())
        r1 = fmt.format_heading("一级", 1)
        assert r1 == "1 一级"
        r2 = fmt.format_heading("二级", 2)
        assert r2 == "1.1 二级"
        r3 = fmt.format_heading("二级二", 2)
        assert r3 == "1.2 二级二"

    def test_hierarchical_level2_resets_after_level1(self) -> None:
        fmt = HeadingFormatter(self._hierarchical_config())
        fmt.format_heading("L1", 1)
        fmt.format_heading("L2a", 2)
        fmt.format_heading("L2b", 2)
        fmt.format_heading("L1b", 1)
        r = fmt.format_heading("L2c", 2)
        assert r == "2.1 L2c"

    def test_reset_counters(self) -> None:
        fmt = HeadingFormatter(self._gongwen_config())
        fmt.format_heading("一", 1)
        fmt.format_heading("二", 1)
        fmt.format_heading("三", 1)
        fmt.reset_counters()
        assert fmt.format_heading("新一", 1) == "一、新一"

    def test_reset_counters_all_levels(self) -> None:
        fmt = HeadingFormatter(self._gongwen_config())
        fmt.format_heading("A", 1)
        fmt.format_heading("B", 2)
        fmt.format_heading("C", 3)
        fmt.reset_counters()
        assert fmt.format_heading("X", 1) == "一、X"
        assert fmt.format_heading("Y", 2) == "（一）Y"
        assert fmt.format_heading("Z", 3) == "1. Z"

    def test_empty_scheme_config(self) -> None:
        fmt = HeadingFormatter({})
        assert fmt.format_heading("标题", 1) == "标题"
        assert fmt.format_heading("标题", 2) == "标题"

    def test_max_level_custom(self) -> None:
        cfg = {"level_5": {"format": "{5.arabic_half}. "}}
        fmt = HeadingFormatter(cfg, max_level=9)
        assert fmt.format_heading("深层次", 5) == "1. 深层次"

    def test_format_heading_with_empty_template(self) -> None:
        cfg = {"level_1": {"format": ""}}
        fmt = HeadingFormatter(cfg)
        assert fmt.format_heading("标题", 1) == "标题"

    def test_non_level_keys_ignored(self) -> None:
        cfg = {
            "level_1": {"format": "{1.chinese_lower}、"},
            "description": "gongwen",
        }
        fmt = HeadingFormatter(cfg)
        assert fmt.format_heading("标题", 1) == "一、标题"

    def test_format_heading_with_unknown_style(self) -> None:
        cfg = {"level_1": {"format": "{1.unknown_style}、"}}
        fmt = HeadingFormatter(cfg)
        result = fmt.format_heading("标题", 1)
        assert result == "1、标题"

    def test_legal_scheme_prefix(self) -> None:
        cfg = {
            "level_1": {"format": "第{1.chinese_lower}编　"},
            "level_2": {"format": "第{2.chinese_lower}章　"},
            "level_3": {"format": "第{3.chinese_lower}节　"},
            "level_4": {"format": "第{4.chinese_lower}条　"},
        }
        fmt = HeadingFormatter(cfg)
        assert fmt.format_heading("总则", 1) == "第一编　总则"
        assert fmt.format_heading("一般规定", 2) == "第一章　一般规定"
        assert fmt.format_heading("范围", 3) == "第一节　范围"
        assert fmt.format_heading("第一条", 4) == "第一条　第一条"


# Integration tests


class TestIntegration:
    def test_detect_then_format_roundtrip(self) -> None:
        raw = "一、引言"
        info = detect_heading_prefix(raw)
        assert info is not None
        assert info.rule_id == "chinese_顿号"
        assert info.clean_text == "引言"
        cfg = {"level_1": {"format": "{1.chinese_lower}、"}}
        fmt = HeadingFormatter(cfg)
        formatted = fmt.format_heading(info.clean_text, 1)
        assert formatted == "一、引言"

    def test_multiple_detections_in_sequence(self) -> None:
        cases = [
            ("第一章　总则", "legal_unit", "总则"),
            ("一、概述", "chinese_顿号", "概述"),
            ("（一）背景", "chinese_bracket", "背景"),
            ("1. 详情", "arabic_separator", " 详情"),
            ("（1）说明", "arabic_bracket", "说明"),
            ("①备注", "circled", "备注"),
            ("1.1 内容", "hierarchical", "内容"),
        ]
        for text, expected_rule, expected_clean in cases:
            info = detect_heading_prefix(text)
            assert info is not None, f"Failed to detect prefix in {text!r}"
            assert info.rule_id == expected_rule, f"For {text!r}: expected rule {expected_rule!r}, got {info.rule_id!r}"
            assert info.clean_text == expected_clean, (
                f"For {text!r}: expected clean_text {expected_clean!r}, got {info.clean_text!r}"
            )

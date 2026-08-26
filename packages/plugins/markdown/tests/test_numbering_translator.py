"""Tests for scheme format string → Word multi-level list translation.

Covers the 4 built-in schemes, edge cases, and compatibility analysis.
"""

from __future__ import annotations

import pytest
from tests.support.numbering import repository_numbering_registry

from docwen_core.text.numbering_word_adapter import translate_scheme

pytestmark = pytest.mark.unit

# ═══════════════════════════════════════════════════════════════════════════════
# Helper
# ═══════════════════════════════════════════════════════════════════════════════


def _make_level(level_num: int, fmt: str) -> dict:
    return {f"level_{level_num}": {"format": fmt}}


def _build_scheme(*level_fmts: tuple[int, str]) -> dict:
    """Build a scheme_config dict from (level_num, format_string) pairs."""
    result: dict = {}
    for num, fmt in level_fmts:
        result[f"level_{num}"] = {"format": fmt}
    return result


# ═══════════════════════════════════════════════════════════════════════════════
# Gongwen Standard (approximate)
# ═══════════════════════════════════════════════════════════════════════════════


class TestGongwenStandard:
    """公文标准: 1-5 compatible, 6-9 conflict with level 1 style."""

    SCHEME = _build_scheme(
        (1, "{1.chinese_lower}、"),
        (2, "（{2.chinese_lower}）"),
        (3, "{3.arabic_half}. "),
        (4, "（{4.arabic_half}）"),
        (5, "{5.arabic_circled}"),
        (6, "{1.arabic_half}.{2.arabic_half}.{3.arabic_half}.{4.arabic_half}.{5.arabic_half}.{6.arabic_half} "),
        (
            7,
            "{1.arabic_half}.{2.arabic_half}.{3.arabic_half}.{4.arabic_half}.{5.arabic_half}.{6.arabic_half}.{7.arabic_half} ",
        ),
        (
            8,
            "{1.arabic_half}.{2.arabic_half}.{3.arabic_half}.{4.arabic_half}.{5.arabic_half}.{6.arabic_half}.{7.arabic_half}.{8.arabic_half} ",
        ),
        (
            9,
            "{1.arabic_half}.{2.arabic_half}.{3.arabic_half}.{4.arabic_half}.{5.arabic_half}.{6.arabic_half}.{7.arabic_half}.{8.arabic_half}.{9.arabic_half} ",
        ),
    )

    def test_overall_verdict_approximate(self):
        result = translate_scheme(self.SCHEME)
        assert result.verdict == "approximate", f"Expected approximate, got {result.verdict}: {result.reason}"

    def test_levels_1_to_5_full(self):
        result = translate_scheme(self.SCHEME)
        for lvl in range(1, 6):
            comp = result.per_level.get(lvl)
            assert comp is not None, f"Level {lvl} missing from per_level"
            assert comp.verdict == "full", f"Level {lvl} expected full, got {comp.verdict}: {comp.reason}"

    def test_levels_6_to_9_unsupported(self):
        result = translate_scheme(self.SCHEME)
        for lvl in range(6, 10):
            comp = result.per_level.get(lvl)
            assert comp is not None, f"Level {lvl} missing from per_level"
            assert comp.verdict == "unsupported", f"Level {lvl} expected unsupported, got {comp.verdict}"

    def test_level_1_values(self):
        result = translate_scheme(self.SCHEME)
        lvl = result.levels[0]
        assert lvl.ilvl == 0
        assert lvl.num_fmt == "chineseCounting"
        assert lvl.lvl_text == "%1、"
        assert lvl.suff == "nothing"
        assert lvl.p_style == "Heading1"
        assert lvl.start == "1"

    def test_level_2_values(self):
        result = translate_scheme(self.SCHEME)
        lvl = result.levels[1]
        assert lvl.ilvl == 1
        assert lvl.num_fmt == "chineseCounting"
        assert lvl.lvl_text == "（%2）"
        assert lvl.suff == "nothing"
        assert lvl.p_style == "Heading2"

    def test_level_3_values(self):
        result = translate_scheme(self.SCHEME)
        lvl = result.levels[2]
        assert lvl.ilvl == 2
        assert lvl.num_fmt == "decimal"
        assert lvl.lvl_text == "%3."
        assert lvl.suff == "space", "Level 3 has trailing space stripped"

    def test_level_4_values(self):
        result = translate_scheme(self.SCHEME)
        lvl = result.levels[3]
        assert lvl.ilvl == 3
        assert lvl.num_fmt == "decimal"
        assert lvl.lvl_text == "（%4）"
        assert lvl.suff == "nothing"

    def test_level_5_values(self):
        result = translate_scheme(self.SCHEME)
        lvl = result.levels[4]
        assert lvl.ilvl == 4
        assert lvl.num_fmt == "decimalEnclosedCircleChinese"
        assert lvl.lvl_text == "%5"
        assert lvl.suff == "space", "Level 5 has no separator and no trailing whitespace → default space"

    def test_reason_mentions_conflict(self):
        result = translate_scheme(self.SCHEME)
        assert result.reason, "Reason should not be empty for approximate verdict"
        assert "conflicting" in result.reason.lower() or "multiple" in result.reason.lower(), (
            f"Reason should mention style conflict: {result.reason}"
        )

    def test_only_compatible_levels_in_output(self):
        result = translate_scheme(self.SCHEME)
        assert len(result.levels) == 5, f"Expected 5 compatible levels, got {len(result.levels)}"


# ═══════════════════════════════════════════════════════════════════════════════
# Hierarchical Standard (full)
# ═══════════════════════════════════════════════════════════════════════════════


class TestHierarchicalStandard:
    """层级数字标准: All levels use arabic_half → no conflicts."""

    SCHEME = _build_scheme(
        *[(n, ".".join(f"{{{i}.arabic_half}}" for i in range(1, n + 1)) + " ") for n in range(1, 10)]
    )

    def test_overall_verdict_full(self):
        result = translate_scheme(self.SCHEME)
        assert result.verdict == "full", f"Expected full, got {result.verdict}: {result.reason}"

    def test_all_levels_full(self):
        result = translate_scheme(self.SCHEME)
        for lvl in range(1, 10):
            comp = result.per_level.get(lvl)
            assert comp is not None, f"Level {lvl} missing"
            assert comp.verdict == "full", f"Level {lvl} expected full, got {comp.verdict}: {comp.reason}"

    def test_all_9_levels_translated(self):
        result = translate_scheme(self.SCHEME)
        assert len(result.levels) == 9

    def test_level_1_values(self):
        result = translate_scheme(self.SCHEME)
        lvl = result.levels[0]
        assert lvl.ilvl == 0
        assert lvl.num_fmt == "decimal"
        assert lvl.lvl_text == "%1"
        assert lvl.suff == "space"
        assert lvl.p_style == "Heading1"

    def test_level_2_values(self):
        result = translate_scheme(self.SCHEME)
        lvl = result.levels[1]
        assert lvl.ilvl == 1
        assert lvl.num_fmt == "decimal"
        assert lvl.lvl_text == "%1.%2"
        assert lvl.suff == "space"

    def test_level_3_values(self):
        result = translate_scheme(self.SCHEME)
        lvl = result.levels[2]
        assert lvl.num_fmt == "decimal"
        assert lvl.lvl_text == "%1.%2.%3"
        assert lvl.suff == "space"

    def test_level_9_values(self):
        result = translate_scheme(self.SCHEME)
        lvl = result.levels[8]
        assert lvl.ilvl == 8
        assert lvl.num_fmt == "decimal"
        assert lvl.lvl_text == "%1.%2.%3.%4.%5.%6.%7.%8.%9"
        assert lvl.suff == "space"
        assert lvl.p_style == "Heading9"


# ═══════════════════════════════════════════════════════════════════════════════
# Hierarchical H2 Start (full)
# ═══════════════════════════════════════════════════════════════════════════════


class TestHierarchicalH2Start:
    """层级数字(H2起): Level 1 is empty, others use arabic_half."""

    SCHEME = _build_scheme(
        (1, ""),
        (2, "{2.arabic_half} "),
        (3, "{2.arabic_half}.{3.arabic_half} "),
        (4, "{2.arabic_half}.{3.arabic_half}.{4.arabic_half} "),
        (5, "{2.arabic_half}.{3.arabic_half}.{4.arabic_half}.{5.arabic_half} "),
        (6, "{2.arabic_half}.{3.arabic_half}.{4.arabic_half}.{5.arabic_half}.{6.arabic_half} "),
        (7, "{2.arabic_half}.{3.arabic_half}.{4.arabic_half}.{5.arabic_half}.{6.arabic_half}.{7.arabic_half} "),
        (
            8,
            "{2.arabic_half}.{3.arabic_half}.{4.arabic_half}.{5.arabic_half}.{6.arabic_half}.{7.arabic_half}.{8.arabic_half} ",
        ),
        (
            9,
            "{2.arabic_half}.{3.arabic_half}.{4.arabic_half}.{5.arabic_half}.{6.arabic_half}.{7.arabic_half}.{8.arabic_half}.{9.arabic_half} ",
        ),
    )

    def test_overall_verdict_full(self):
        result = translate_scheme(self.SCHEME)
        assert result.verdict == "full"

    def test_all_levels_full(self):
        result = translate_scheme(self.SCHEME)
        for lvl in range(1, 10):
            comp = result.per_level.get(lvl)
            assert comp is not None, f"Level {lvl} missing"
            assert comp.verdict == "full", f"Level {lvl} expected full, got {comp.verdict}"

    def test_all_9_levels_translated(self):
        result = translate_scheme(self.SCHEME)
        # Level 1 has empty format but should still produce a level entry
        # with default numFmt
        assert len(result.levels) == 9

    def test_level_1_empty_values(self):
        result = translate_scheme(self.SCHEME)
        lvl = result.levels[0]
        assert lvl.ilvl == 0
        assert lvl.num_fmt == "decimal"  # default fallback
        assert lvl.lvl_text == ""
        assert lvl.suff == "nothing", "Empty lvlText → nothing suff"
        assert lvl.p_style == "Heading1"

    def test_level_2_values(self):
        result = translate_scheme(self.SCHEME)
        lvl = result.levels[1]
        assert lvl.ilvl == 1
        assert lvl.num_fmt == "decimal"
        assert lvl.lvl_text == "%2"
        assert lvl.suff == "space"

    def test_level_3_values(self):
        result = translate_scheme(self.SCHEME)
        lvl = result.levels[2]
        assert lvl.num_fmt == "decimal"
        assert lvl.lvl_text == "%2.%3"
        assert lvl.suff == "space"

    def test_no_conflict_reason(self):
        result = translate_scheme(self.SCHEME)
        assert result.reason == ""


# ═══════════════════════════════════════════════════════════════════════════════
# Legal Standard (approximate)
# ═══════════════════════════════════════════════════════════════════════════════


class TestLegalStandard:
    """法律条文标准: 1-5 compatible, 6-9 conflict with level 1 style."""

    SCHEME = _build_scheme(
        (1, "第{1.chinese_lower}编\u3000"),
        (2, "第{2.chinese_lower}章\u3000"),
        (3, "第{3.chinese_lower}节\u3000"),
        (4, "第{4.chinese_lower}条\u3000"),
        (5, "（{5.chinese_lower}）"),
        (6, "{1.arabic_half}.{2.arabic_half}.{3.arabic_half}.{4.arabic_half}.{5.arabic_half}.{6.arabic_half} "),
        (
            7,
            "{1.arabic_half}.{2.arabic_half}.{3.arabic_half}.{4.arabic_half}.{5.arabic_half}.{6.arabic_half}.{7.arabic_half} ",
        ),
        (
            8,
            "{1.arabic_half}.{2.arabic_half}.{3.arabic_half}.{4.arabic_half}.{5.arabic_half}.{6.arabic_half}.{7.arabic_half}.{8.arabic_half} ",
        ),
        (
            9,
            "{1.arabic_half}.{2.arabic_half}.{3.arabic_half}.{4.arabic_half}.{5.arabic_half}.{6.arabic_half}.{7.arabic_half}.{8.arabic_half}.{9.arabic_half} ",
        ),
    )

    def test_overall_verdict_approximate(self):
        result = translate_scheme(self.SCHEME)
        assert result.verdict == "approximate"

    def test_levels_1_to_5_full(self):
        result = translate_scheme(self.SCHEME)
        for lvl in range(1, 6):
            comp = result.per_level.get(lvl)
            assert comp is not None, f"Level {lvl} missing"
            assert comp.verdict == "full", f"Level {lvl} expected full, got {comp.verdict}: {comp.reason}"

    def test_levels_6_to_9_unsupported(self):
        result = translate_scheme(self.SCHEME)
        for lvl in range(6, 10):
            comp = result.per_level.get(lvl)
            assert comp is not None, f"Level {lvl} missing"
            assert comp.verdict == "unsupported"

    def test_level_1_fullwidth_space_stripped(self):
        result = translate_scheme(self.SCHEME)
        lvl = result.levels[0]
        assert lvl.lvl_text == "第%1编"
        assert lvl.suff == "space", "Full-width space stripped → suff=space"

    def test_level_2_values(self):
        result = translate_scheme(self.SCHEME)
        lvl = result.levels[1]
        assert lvl.num_fmt == "chineseCounting"
        assert lvl.lvl_text == "第%2章"
        assert lvl.suff == "space"

    def test_level_3_values(self):
        result = translate_scheme(self.SCHEME)
        lvl = result.levels[2]
        assert lvl.num_fmt == "chineseCounting"
        assert lvl.lvl_text == "第%3节"
        assert lvl.suff == "space"

    def test_level_4_values(self):
        result = translate_scheme(self.SCHEME)
        lvl = result.levels[3]
        assert lvl.num_fmt == "chineseCounting"
        assert lvl.lvl_text == "第%4条"
        assert lvl.suff == "space"

    def test_level_5_values(self):
        result = translate_scheme(self.SCHEME)
        lvl = result.levels[4]
        assert lvl.num_fmt == "chineseCounting"
        assert lvl.lvl_text == "（%5）"
        assert lvl.suff == "nothing"

    def test_only_5_levels_translated(self):
        result = translate_scheme(self.SCHEME)
        assert len(result.levels) == 5


# ═══════════════════════════════════════════════════════════════════════════════
# Edge cases
# ═══════════════════════════════════════════════════════════════════════════════


class TestEdgeCases:
    """Various edge and boundary cases."""

    def test_empty_scheme_config(self):
        """Empty config → full verdict, no levels."""
        result = translate_scheme({})
        assert result.verdict == "full"
        assert result.levels == []
        assert result.per_level == {}
        assert result.reason == ""

    def test_unknown_style(self):
        """A placeholder with an unknown style → unsupported."""
        scheme = _build_scheme((1, "{1.unknown_style}"))
        result = translate_scheme(scheme)
        assert result.verdict == "unsupported"
        comp = result.per_level.get(1)
        assert comp is not None
        assert comp.verdict == "unsupported"

    def test_single_level_only(self):
        """Only level 1 defined → translate it happily."""
        scheme = _build_scheme((1, "{1.arabic_half}"))
        result = translate_scheme(scheme)
        assert result.verdict == "full"
        assert len(result.levels) == 1
        assert result.levels[0].ilvl == 0
        assert result.levels[0].num_fmt == "decimal"
        assert result.levels[0].lvl_text == "%1"

    def test_fullwidth_space_handling(self):
        """Full-width space ``\\u3000`` stripping → suff=space."""
        scheme = _build_scheme((1, "{1.chinese_lower}　"))
        result = translate_scheme(scheme)
        assert result.verdict == "full"
        lvl = result.levels[0]
        assert lvl.lvl_text == "%1"
        assert lvl.suff == "space"

    def test_no_trailing_whitespace_no_separator(self):
        """Format with neither trailing space nor separator → suff=space."""
        scheme = _build_scheme((1, "{1.arabic_half}"))
        result = translate_scheme(scheme)
        lvl = result.levels[0]
        assert lvl.lvl_text == "%1"
        assert lvl.suff == "space"

    def test_trailing_separator_nothing(self):
        """Format ending with separator character → suff=nothing."""
        scheme = _build_scheme((1, "{1.chinese_lower}、"))
        result = translate_scheme(scheme)
        lvl = result.levels[0]
        assert lvl.lvl_text == "%1、"
        assert lvl.suff == "nothing"

    def test_trailing_right_paren_nothing(self):
        """Format ending with ） → suff=nothing."""
        scheme = _build_scheme((1, "（{1.arabic_half}）"))
        result = translate_scheme(scheme)
        lvl = result.levels[0]
        assert lvl.lvl_text == "（%1）"
        assert lvl.suff == "nothing"

    def test_trailing_regular_paren_nothing(self):
        """Format ending with ) → suff=nothing."""
        scheme = _build_scheme((1, "({1.arabic_half})"))
        result = translate_scheme(scheme)
        lvl = result.levels[0]
        assert lvl.lvl_text == "(%1)"
        assert lvl.suff == "nothing"

    def test_trailing_jp_period_nothing(self):
        """Format ending with ． → suff=nothing."""
        scheme = _build_scheme((1, "{1.arabic_half}．"))
        result = translate_scheme(scheme)
        lvl = result.levels[0]
        assert lvl.lvl_text == "%1．"
        assert lvl.suff == "nothing"


# ═══════════════════════════════════════════════════════════════════════════════
# Failure semantics: unsupported and approximate verdicts (Phase C)
# ═══════════════════════════════════════════════════════════════════════════════


class TestUnsupportedVerdictSemantics:
    """Scheme compatibility failures are diagnosed per-level and
    aggregated to 'unsupported' overall verdict — never silently dropped."""

    def test_unknown_style_all_levels(self):
        """Every level referencing an unknown style -> unsupported overall."""
        scheme = _build_scheme(
            (1, "{1.bogus_style}"),
            (2, "{2.bogus_style}"),
        )
        result = translate_scheme(scheme)
        assert result.verdict == "unsupported"
        assert result.reason != ""
        assert result.levels == []
        assert result.per_level[1].verdict == "unsupported"
        assert result.per_level[2].verdict == "unsupported"

    def test_mixed_full_and_unsupported(self):
        """level_1 ok, level_2 unknown -> overall unsupported."""
        scheme = _build_scheme(
            (1, "{1.chinese_lower}、"),
            (2, "{2.unknown}"),
            (3, "{3.arabic_half}."),
        )
        result = translate_scheme(scheme)
        assert result.verdict == "unsupported", f"Expected unsupported when ANY level fails, got {result.verdict}"
        assert result.per_level[1].verdict == "full"
        assert result.per_level[2].verdict == "unsupported"
        assert result.per_level[3].verdict == "full"

    def test_unsupported_partial_levels_for_mixed(self):
        """When unsupported due to one bad level, compatible levels still translate.

        This documents actual translator behavior: an 'unsupported' overall
        verdict means the scheme cannot be used reliably, but compatible
        levels are still produced so the caller can make their own decision.
        """
        scheme = _build_scheme(
            (1, "{1.chinese_lower}、"),
            (2, "{2.bogus}"),
        )
        result = translate_scheme(scheme)
        assert result.verdict == "unsupported"
        # level 1 is compatible and translated; level 2 is the blocker
        assert result.per_level[2].verdict == "unsupported"
        assert len(result.levels) >= 1


class TestApproximateVerdictSemantics:
    """The built-in gongwen_standard and legal_standard schemes are
    'approximate' because levels 6+ reference the same style as level 1
    with a cross-reference to level 1's counter, which triggers a
    numFmt conflict for the cross-reference.  The verdict is driven by
    the actual scheme configuration, not by the scheme id."""

    def test_gongwen_standard_is_approximate(self):
        """Gongwen from registry must be approximate (6-9 conflict with L1)."""
        registry = repository_numbering_registry()
        info = registry.get_scheme("gongwen_standard")
        # scheme_info.levels is a dict of str (level_1 .. level_9) -> format str
        config: dict = {}
        for key, fmt in info.levels.items():
            config[key] = {"format": fmt}
        result = translate_scheme(config)
        assert result.verdict == "approximate"
        assert len(result.levels) == 5  # only 1-5 translated
        for lvl in [6, 7, 8, 9]:
            assert result.per_level[lvl].verdict == "unsupported"

    def test_legal_standard_is_approximate(self):
        """Legal standard from registry must be approximate."""
        registry = repository_numbering_registry()
        info = registry.get_scheme("legal_standard")
        config: dict = {}
        for key, fmt in info.levels.items():
            config[key] = {"format": fmt}
        result = translate_scheme(config)
        assert result.verdict == "approximate"
        assert len(result.levels) == 5

    def test_approximate_has_warning_reason(self):
        """approximate verdict has a non-empty reason explaining the conflict."""
        registry = repository_numbering_registry()
        info = registry.get_scheme("gongwen_standard")
        config: dict = {}
        for key, fmt in info.levels.items():
            config[key] = {"format": fmt}
        result = translate_scheme(config)
        assert result.reason != ""


class TestFullVerdictAssertsStability:
    """full verdict means ALL levels translate and the result is stable
    (same input -> same output, idempotent)."""

    def test_hierarchical_is_full(self):
        """hierarchical_standard from registry must be full."""
        registry = repository_numbering_registry()
        info = registry.get_scheme("hierarchical_standard")
        # scheme_info.levels is a dict of str (level_1 .. level_9) -> format str
        config: dict = {}
        for key, fmt in info.levels.items():
            config[key] = {"format": fmt}
        result = translate_scheme(config)
        assert result.verdict == "full"
        assert result.reason == ""
        assert len(result.levels) == 9

    def test_hierarchical_h2_start_is_full(self):
        """hierarchical_h2_start from registry must be full."""
        registry = repository_numbering_registry()
        info = registry.get_scheme("hierarchical_h2_start")
        config: dict = {}
        for key, fmt in info.levels.items():
            config[key] = {"format": fmt}
        result = translate_scheme(config)
        assert result.verdict == "full"
        assert len(result.levels) >= 9  # h2_start defines levels 1-9

    def test_arabic_half_single_level_idempotent(self):
        """Single level, single fmt -> always full and idempotent."""
        scheme = _build_scheme((1, "{1.arabic_half}."))
        r1 = translate_scheme(scheme)
        r2 = translate_scheme(scheme)
        assert r1.verdict == "full"
        assert r1 == r2, "translate_scheme must be pure"


# ═══════════════════════════════════════════════════════════════════════════════
# Registry-based tests (4 built-in schemes)
# ═══════════════════════════════════════════════════════════════════════════════


class TestBuiltinSchemesViaRegistry:
    """Load built-in schemes from NumberingSchemeRegistry and verify verdicts."""

    def _scheme_config_from_registry(self, scheme_id: str) -> dict:
        """Load from NumberingSchemeRegistry and convert to plain dict format."""
        registry = repository_numbering_registry()
        info = registry.get_scheme(scheme_id)
        config: dict = {}
        for key, fmt in info.levels.items():
            # key is like "level_1", value is format string
            config[key] = {"format": fmt}
        return config

    def test_gongwen_standard_via_registry(self):
        config = self._scheme_config_from_registry("gongwen_standard")
        result = translate_scheme(config)
        assert result.verdict == "approximate", f"gongwen_standard expected approximate, got {result.verdict}"

    def test_hierarchical_standard_via_registry(self):
        config = self._scheme_config_from_registry("hierarchical_standard")
        result = translate_scheme(config)
        assert result.verdict == "full"

    def test_hierarchical_h2_start_via_registry(self):
        config = self._scheme_config_from_registry("hierarchical_h2_start")
        result = translate_scheme(config)
        assert result.verdict == "full"

    def test_legal_standard_via_registry(self):
        config = self._scheme_config_from_registry("legal_standard")
        result = translate_scheme(config)
        assert result.verdict == "approximate"

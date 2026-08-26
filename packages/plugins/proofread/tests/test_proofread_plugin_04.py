"""Focused tests split from test_proofread_plugin.py."""

from __future__ import annotations

from ._proofread_plugin_support import (
    pytest,
)

pytestmark = pytest.mark.golden


@pytest.mark.unit
class TestTextValidator:
    def test_empty_text(self) -> None:
        """Empty text should return no errors."""
        from docwen_plugin_proofread.text_validator import TextValidator

        v = TextValidator()
        errors = v.validate_text("")
        assert errors == []

    def test_symbol_pairing_matched(self) -> None:
        """Correctly paired brackets should produce no errors."""
        from docwen_plugin_proofread.text_validator import TextValidator

        v = TextValidator()
        errors = v.validate_text("（这是一段文字）[还有方括号]")
        assert errors == []

    def test_symbol_pairing_unmatched_opening(self) -> None:
        """An unclosed opening bracket should be flagged."""
        from docwen_plugin_proofread.text_validator import TextValidator

        v = TextValidator()
        errors = v.validate_text("（没有闭合")
        assert len(errors) == 1
        assert errors[0].error_text == "（"
        assert errors[0].source == "pairing"
        assert errors[0].suggestion == "Unmatched Symbol"

    def test_symbol_pairing_extra_closing(self) -> None:
        """An extra closing bracket with no opener should be flagged."""
        from docwen_plugin_proofread.text_validator import TextValidator

        v = TextValidator()
        errors = v.validate_text("没有开括号）")
        assert len(errors) == 1
        assert errors[0].error_text == "）"
        assert errors[0].source == "pairing"
        assert errors[0].suggestion == "Unmatched Symbol"

    @pytest.mark.parametrize(
        "text",
        (
            'He said "ok".',
            "He said 'ok'.",
            "It's ready and the students' work is complete.",
            "It’s ready and the students’ work is complete.",
            "The '90s are back.",
            "The ’90s are back.",
            "'tis ready.",
            "’tis ready.",
            "’cause it works.",
            "Rock ’n’ roll.",
            "He said 'don't stop'.",
            "He said ‘don’t stop’.",
            'First "one"; then "two".',
        ),
    )
    def test_symbol_pairing_accepts_symmetric_quotes_and_apostrophes(self, text: str) -> None:
        """Balanced quotes and word apostrophes must not create false positives."""
        from docwen_plugin_proofread.text_validator import TextValidator

        pairing_errors = [error for error in TextValidator().validate_text(text) if error.source == "pairing"]

        assert pairing_errors == []

    @pytest.mark.parametrize(
        ("text", "expected_text", "expected_position"),
        (
            ('He said "oops.', '"', 8),
            ("He said 'oops.", "'", 8),
            ("He said ‘oops.", "‘", 8),
            ('First "one"; then "two.', '"', 18),
        ),
    )
    def test_symbol_pairing_reports_only_the_unmatched_quote(
        self,
        text: str,
        expected_text: str,
        expected_position: int,
    ) -> None:
        """An odd quote remains one precise pairing issue after balanced spans."""
        from docwen_plugin_proofread.text_validator import TextValidator

        pairing_errors = [error for error in TextValidator().validate_text(text) if error.source == "pairing"]

        assert [(error.error_text, error.start_pos, error.end_pos) for error in pairing_errors] == [
            (expected_text, expected_position, expected_position + 1)
        ]

    @pytest.mark.parametrize(
        ("text", "expected"),
        (
            ("He said '123.", [("'", 8)]),
            ("He said '9 cats.", [("'", 8)]),
            ("He said ’oops.", [("’", 8)]),
            ("He said ’hello’.", [("’", 8)]),
        ),
    )
    def test_symbol_pairing_does_not_overgeneralize_leading_elisions(
        self,
        text: str,
        expected: list[tuple[str, int]],
    ) -> None:
        """Arbitrary quoted words and numbers remain pairing diagnostics."""
        from docwen_plugin_proofread.text_validator import TextValidator

        pairing_errors = [error for error in TextValidator().validate_text(text) if error.source == "pairing"]

        assert [(error.error_text, error.start_pos) for error in pairing_errors] == expected

    def test_markdown_sanitizer_blanks_escaped_quote_without_offset_drift(self) -> None:
        """A Markdown escape is not a delimiter and later coordinates remain exact."""
        from docwen_plugin_proofread.md_validator import _sanitize_markdown
        from docwen_plugin_proofread.text_validator import TextValidator

        text = 'Escaped \\" literal; unmatched （'
        sanitized = _sanitize_markdown(text)
        escaped_quote = text.index('"')
        unmatched = text.index("（")

        assert len(sanitized.sanitized_text) == len(text)
        assert sanitized.sanitized_text[escaped_quote] == " "
        pairing_errors = [
            error for error in TextValidator().validate_text(sanitized.sanitized_text) if error.source == "pairing"
        ]
        assert [(error.error_text, error.start_pos) for error in pairing_errors] == [("（", unmatched)]

    @pytest.mark.parametrize("text", (r"\[literal]", r"\(literal)", r"\{literal}"))
    def test_markdown_escaped_opening_brackets_remain_balanced(self, text: str) -> None:
        """Escaping Markdown syntax must not manufacture an unmatched closer."""
        from docwen_plugin_proofread.md_validator import _sanitize_markdown
        from docwen_plugin_proofread.text_validator import TextValidator

        sanitized = _sanitize_markdown(text)
        pairing_errors = [
            error for error in TextValidator().validate_text(sanitized.sanitized_text) if error.source == "pairing"
        ]

        assert sanitized.sanitized_text == text
        assert pairing_errors == []

    def test_markdown_escaped_quote_does_not_hide_a_later_unmatched_quote(self) -> None:
        """Only the escaped quote is literal; a later delimiter is still checked."""
        from docwen_plugin_proofread.md_validator import _sanitize_markdown
        from docwen_plugin_proofread.text_validator import TextValidator

        text = r'\"literal"'
        sanitized = _sanitize_markdown(text)
        pairing_errors = [
            error for error in TextValidator().validate_text(sanitized.sanitized_text) if error.source == "pairing"
        ]

        assert sanitized.sanitized_text[text.index('"')] == " "
        assert [(error.error_text, error.start_pos) for error in pairing_errors] == [('"', len(text) - 1)]

    def test_symbol_correction_fullwidth_digits(self) -> None:
        """Fullwidth digits should be flagged."""
        from docwen_plugin_proofread.text_validator import TextValidator

        v = TextValidator()
        errors = v.validate_text("价格１２３")
        assert len(errors) == 3  # three fullwidth digits
        for err in errors:
            assert err.source == "symbol"

    def test_typo_detection(self) -> None:
        """Common Chinese typos should be detected when typos_map is supplied."""
        from docwen_plugin_proofread.text_validator import TextValidator

        v = TextValidator(typos_map={"已": ["己"]})
        errors = v.validate_text("我己经完成")  # 己 → 已
        typo_errors = [e for e in errors if e.source == "typo"]
        assert len(typo_errors) >= 1
        assert typo_errors[0].suggestion == "已"

    def test_typo_detection_multiple_rules(self) -> None:
        """Multiple different typo rules should all be detected."""
        from docwen_plugin_proofread.text_validator import TextValidator

        v = TextValidator(typos_map={"未": ["末"], "作": ["做"]})
        errors = v.validate_text("我末来想做这件事")  # 末→未, 做→作
        typo_errors = [e for e in errors if e.source == "typo"]
        assert len(typo_errors) >= 1

    def test_typo_detection_various(self) -> None:
        """Verify several typo rules produce expected suggestions."""
        from docwen_plugin_proofread.text_validator import TextValidator

        sample_map = {
            "已": ["己"],
            "人": ["入"],
            "侯": ["候"],
            "即": ["既"],
            "象": ["像"],
            "坐": ["座"],
        }
        v = TextValidator(typos_map=sample_map)
        # Each tuple: (wrong_char, expected_correction)
        test_cases = [
            ("己", "已"),  # 己 → 已
            ("入", "人"),  # 入 → 人
            ("候", "侯"),  # 候 → 侯
            ("既", "即"),  # 既 → 即
            ("像", "象"),  # 像 → 象
            ("座", "坐"),  # 座 → 坐
        ]
        for wrong, expected_correction in test_cases:
            errors = v.validate_text(wrong)
            typo = [e for e in errors if e.source == "typo"]
            assert len(typo) >= 1, f"Expected typo detection for '{wrong}'"
            assert typo[0].suggestion == expected_correction, (
                f"'{wrong}' → expected '{expected_correction}', got '{typo[0].suggestion}'"
            )

    def test_sensitive_word_disabled_by_default(self) -> None:
        """Sensitive word check should be off by default."""
        from docwen_plugin_proofread.text_validator import TextValidator

        v = TextValidator()
        # Even if there were sensitive words, they shouldn't be checked
        errors = v.validate_text("test")
        # No errors from sensitive_word since it's disabled
        sensitive = [e for e in errors if e.source == "sensitive"]
        assert len(sensitive) == 0

    def test_disabled_check_no_errors(self) -> None:
        """Disabling a check should prevent errors from that source."""
        from docwen_plugin_proofread.text_validator import TextValidator

        v = TextValidator(
            enabled={"symbol_pairing": False, "symbol_correction": False, "typos_rule": False, "sensitive_word": False},
        )
        errors = v.validate_text("（未闭合１２３我己经")
        assert errors == []  # all checks disabled

    def test_multiple_error_types(self) -> None:
        """A text with multiple error types should report all of them."""
        from docwen_plugin_proofread.text_validator import TextValidator

        v = TextValidator(typos_map={"已": ["己"]})
        errors = v.validate_text("（未闭合 我己经 １２３")
        sources = {e.source for e in errors}
        assert "pairing" in sources  # unmatched bracket
        assert "typo" in sources  # 己→已
        assert "symbol" in sources  # fullwidth digits

    def test_error_fields(self) -> None:
        """Each TextError should have all required fields."""
        from docwen_plugin_proofread.text_validator import TextValidator

        # Verify fields using a text that triggers symbol correction (always enabled)
        v = TextValidator()
        errors = v.validate_text("１２３")  # fullwidth digits
        for err in errors:
            assert err.start_pos >= 0
            assert err.end_pos > err.start_pos
            assert len(err.error_text) > 0
            assert len(err.suggestion) > 0
            assert len(err.error_type) > 0
            assert len(err.source) > 0

    def test_any_enabled(self) -> None:
        """any_enabled should reflect the current state."""
        from docwen_plugin_proofread.text_validator import TextValidator

        v = TextValidator(
            enabled={"symbol_pairing": True, "symbol_correction": False, "typos_rule": False, "sensitive_word": False},
        )
        assert v.any_enabled() is True

        v2 = TextValidator(
            enabled={"symbol_pairing": False, "symbol_correction": False, "typos_rule": False, "sensitive_word": False},
        )
        assert v2.any_enabled() is False

    def test_no_symbol_pairs_no_errors(self) -> None:
        """Empty symbol_pairs should not cause errors."""
        from docwen_plugin_proofread.text_validator import TextValidator

        v = TextValidator(
            symbol_pairs=[],
            symbol_map={},
            typos_map={},
            enabled={"symbol_pairing": True},
        )
        errors = v.validate_text("（test")
        assert errors == []  # no pairs defined, so no errors

    def test_no_typos_map_no_typo_errors(self) -> None:
        """Empty typos_map should not cause typo errors."""
        from docwen_plugin_proofread.text_validator import TextValidator

        v = TextValidator(
            symbol_pairs=[],
            symbol_map={},
            typos_map={},
            enabled={"typos_rule": True},
        )
        errors = v.validate_text("我己经")
        typo = [e for e in errors if e.source == "typo"]
        assert len(typo) == 0  # no typo rules defined

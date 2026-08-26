"""Focused tests split from test_convert_typed_options.py."""

from __future__ import annotations

from ._convert_typed_options_support import (
    _fake_convert_args,
    argparse,
    pytest,
)

pytestmark = pytest.mark.unit


class TestCanonicalRouteOptionProjection:
    def test_explicit_option_not_declared_by_route_is_rejected(self) -> None:
        from docwen_cli.commands.execution_request import project_route_options

        with pytest.raises(ValueError, match=r"route-1.*render_dpi"):
            project_route_options(
                {"render_dpi": 300},
                route_id="route-1",
                route_options=(),
            )

    def test_runtime_defaults_are_only_injected_when_declared(self) -> None:
        from docwen_cli.commands.execution_request import project_route_options

        without_ocr = project_route_options(
            {},
            route_id="route-without-ocr",
            route_options=("locale",),
            configured_ocr_language="japanese",
            ocr_requested=True,
        )
        with_ocr = project_route_options(
            {},
            route_id="route-with-ocr",
            route_options=("to_md_enable_ocr", "ocr_language"),
            configured_ocr_language="japanese",
            ocr_requested=True,
        )
        without_request = project_route_options(
            {},
            route_id="route-with-ocr-default",
            route_options=("to_md_enable_ocr",),
            configured_ocr_language="japanese",
            ocr_requested=False,
        )

        assert "ocr_language" not in without_ocr
        assert with_ocr == {"to_md_enable_ocr": True, "ocr_language": "japanese"}
        assert without_request == {"to_md_enable_ocr": False}

    def test_docx_route_projects_the_active_cli_locale(self) -> None:
        from docwen_cli.commands.execution_request import project_route_options
        from docwen_cli.i18n import init_cli_locale

        init_cli_locale("de_DE")
        try:
            projected = project_route_options(
                {},
                route_id="docwen_plugin_markdown:markdown:docx:convert",
                route_options=("locale", "template_name"),
            )
        finally:
            init_cli_locale("zh_CN")

        assert projected == {"locale": "de_DE"}

    def test_unrequested_ocr_does_not_emit_a_false_default(self) -> None:
        from docwen_cli.commands.execution_options import build_execution_options

        options = build_execution_options(_fake_convert_args({"to": "md", "ocr": False}))

        assert "to_md_enable_ocr" not in options

    def test_ocr_config_read_failure_is_not_silently_defaulted(self) -> None:
        from docwen_application.controller import CapabilityUnavailableError
        from docwen_cli.commands.execution_request import configured_ocr_language

        class BrokenConfig:
            def get(self, _key: str, _default: object = None) -> object:
                raise OSError("corrupt config")

        with pytest.raises(CapabilityUnavailableError, match="could not be read"):
            configured_ocr_language(argparse.Namespace(config_port=BrokenConfig()))


class TestParsePages:
    """Unit tests for ``parse_pages()`` — page range string → list[int]."""

    def test_single_page(self) -> None:
        from docwen_cli.commands.execution_options import parse_pages

        assert parse_pages("5") == [5]

    def test_comma_separated_pages(self) -> None:
        from docwen_cli.commands.execution_options import parse_pages

        assert parse_pages("1,3,5") == [1, 3, 5]

    def test_inclusive_range(self) -> None:
        from docwen_cli.commands.execution_options import parse_pages

        assert parse_pages("1-3") == [1, 2, 3]

    def test_mixed_ranges_and_singles(self) -> None:
        from docwen_cli.commands.execution_options import parse_pages

        assert parse_pages("1-3,7,9-11") == [1, 2, 3, 7, 9, 10, 11]

    def test_descending_range_reversed(self) -> None:
        from docwen_cli.commands.execution_options import parse_pages

        assert parse_pages("5-3") == [3, 4, 5]

    def test_sorted_and_deduplicated(self) -> None:
        from docwen_cli.commands.execution_options import parse_pages

        assert parse_pages("3,1,2") == [1, 2, 3]

    def test_overlapping_ranges_deduplicated(self) -> None:
        from docwen_cli.commands.execution_options import parse_pages

        assert parse_pages("1-5,3-7") == [1, 2, 3, 4, 5, 6, 7]

    def test_empty_string_returns_empty_list(self) -> None:
        from docwen_cli.commands.execution_options import parse_pages

        assert parse_pages("") == []

    def test_non_numeric_parts_ignored(self) -> None:
        from docwen_cli.commands.execution_options import parse_pages

        # Non-numeric parts are silently skipped (contextlib.suppress)
        assert parse_pages("1,abc,3") == [1, 3]

    def test_malformed_ranges_ignored(self) -> None:
        from docwen_cli.commands.execution_options import parse_pages

        assert parse_pages("abc-def") == []

    def test_negative_page_sign_parsed_as_range(self) -> None:
        """A leading minus sign is treated as a range separator (old behavior).

        This is not a real-world PDF page case — real page numbers are positive.
        The old ``parse_pages`` exhibits the same behaviour:
        ``"-1"`` → ``partition("-")`` → ``("", "-", "1")`` → ``[1]``.
        """
        from docwen_cli.commands.execution_options import parse_pages

        assert parse_pages("-1,1") == [1]
        assert parse_pages("0,1,2") == [0, 1, 2]

    def test_whitespace_tolerant(self) -> None:
        from docwen_cli.commands.execution_options import parse_pages

        assert parse_pages(" 1 , 3 , 5 ") == [1, 3, 5]
        assert parse_pages(" 1 - 3 ") == [1, 2, 3]


class TestNormalizeProofreadOptions:
    """Unit tests for ``normalize_proofread_options()``."""

    def test_empty_checks_returns_empty_dict(self) -> None:
        from docwen_cli.commands.execution_options import normalize_proofread_options

        assert normalize_proofread_options([]) == {}

    def test_none_alone_all_false(self) -> None:
        from docwen_cli.commands.execution_options import normalize_proofread_options

        result = normalize_proofread_options(["none"])
        assert result == {
            "enable_symbol_pairing": False,
            "enable_symbol_correction": False,
            "enable_typos_rule": False,
            "enable_sensitive_word": False,
        }

    def test_all_expands_to_four_categories(self) -> None:
        from docwen_cli.commands.execution_options import normalize_proofread_options

        result = normalize_proofread_options(["all"])
        assert result == {
            "enable_symbol_pairing": True,
            "enable_symbol_correction": True,
            "enable_typos_rule": True,
            "enable_sensitive_word": True,
        }

    def test_all_does_not_include_none_semantics(self) -> None:
        from docwen_cli.commands.execution_options import normalize_proofread_options

        result = normalize_proofread_options(["all"])
        # "all" should expand to all four individual categories as True
        assert result["enable_symbol_pairing"] is True
        assert result["enable_sensitive_word"] is True

    def test_single_punct_only(self) -> None:
        from docwen_cli.commands.execution_options import normalize_proofread_options

        result = normalize_proofread_options(["punct"])
        assert result == {
            "enable_symbol_pairing": True,
            "enable_symbol_correction": False,
            "enable_typos_rule": False,
            "enable_sensitive_word": False,
        }

    def test_single_typo_only(self) -> None:
        from docwen_cli.commands.execution_options import normalize_proofread_options

        result = normalize_proofread_options(["typo"])
        assert result == {
            "enable_symbol_pairing": False,
            "enable_symbol_correction": False,
            "enable_typos_rule": True,
            "enable_sensitive_word": False,
        }

    def test_single_sensitive_only(self) -> None:
        from docwen_cli.commands.execution_options import normalize_proofread_options

        result = normalize_proofread_options(["sensitive"])
        assert result == {
            "enable_symbol_pairing": False,
            "enable_symbol_correction": False,
            "enable_typos_rule": False,
            "enable_sensitive_word": True,
        }

    def test_single_symbol_only(self) -> None:
        from docwen_cli.commands.execution_options import normalize_proofread_options

        result = normalize_proofread_options(["symbol"])
        assert result == {
            "enable_symbol_pairing": False,
            "enable_symbol_correction": True,
            "enable_typos_rule": False,
            "enable_sensitive_word": False,
        }

    def test_combination_punct_typo(self) -> None:
        from docwen_cli.commands.execution_options import normalize_proofread_options

        result = normalize_proofread_options(["punct", "typo"])
        assert result == {
            "enable_symbol_pairing": True,
            "enable_symbol_correction": False,
            "enable_typos_rule": True,
            "enable_sensitive_word": False,
        }

    def test_combination_typo_sensitive(self) -> None:
        from docwen_cli.commands.execution_options import normalize_proofread_options

        result = normalize_proofread_options(["typo", "sensitive"])
        assert result == {
            "enable_symbol_pairing": False,
            "enable_symbol_correction": False,
            "enable_typos_rule": True,
            "enable_sensitive_word": True,
        }

    def test_all_with_explicit_individual_duplicates(self) -> None:
        """'all' + explicit 'punct' should not double-enable anything."""
        from docwen_cli.commands.execution_options import normalize_proofread_options

        result = normalize_proofread_options(["all", "punct"])
        # All four should still be True (set semantics deduplicate)
        assert result["enable_symbol_pairing"] is True
        assert result["enable_symbol_correction"] is True
        assert result["enable_typos_rule"] is True
        assert result["enable_sensitive_word"] is True


class TestNormalizeNumberingOptions:
    """Unit tests for ``normalize_numbering_options()``."""

    def test_both_none_returns_empty(self) -> None:
        from docwen_cli.commands.execution_options import normalize_numbering_options

        assert normalize_numbering_options(None, None) == {}

    def test_clean_remove_no_add(self) -> None:
        from docwen_cli.commands.execution_options import normalize_numbering_options

        result = normalize_numbering_options("remove", None)
        assert result == {
            "remove_numbering": True,
            "add_numbering": False,
            "numbering_scheme": "",
        }

    def test_clean_keep_no_add(self) -> None:
        from docwen_cli.commands.execution_options import normalize_numbering_options

        result = normalize_numbering_options("keep", None)
        assert result == {
            "remove_numbering": False,
            "add_numbering": False,
            "numbering_scheme": "",
        }

    def test_clean_default_no_add(self) -> None:
        from docwen_cli.commands.execution_options import normalize_numbering_options

        result = normalize_numbering_options("default", None)
        assert result == {
            "remove_numbering": False,
            "add_numbering": False,
            "numbering_scheme": "",
        }

    def test_add_scheme_no_clean(self) -> None:
        from docwen_cli.commands.execution_options import normalize_numbering_options

        result = normalize_numbering_options(None, "cn_article")
        assert result == {
            "remove_numbering": False,
            "add_numbering": True,
            "numbering_scheme": "cn_article",
        }

    def test_add_default_no_clean(self) -> None:
        from docwen_cli.commands.execution_options import normalize_numbering_options

        result = normalize_numbering_options(None, "default")
        assert result == {
            "remove_numbering": False,
            "add_numbering": False,
            "numbering_scheme": "",
        }

    def test_add_none_no_clean(self) -> None:
        from docwen_cli.commands.execution_options import normalize_numbering_options

        result = normalize_numbering_options(None, "none")
        assert result == {
            "remove_numbering": False,
            "add_numbering": False,
            "numbering_scheme": "",
        }

    def test_clean_remove_with_gongwen_scheme(self) -> None:
        """Full user path: md-numbering --clean-numbering remove --add-numbering gongwen_standard."""
        from docwen_cli.commands.execution_options import normalize_numbering_options

        result = normalize_numbering_options("remove", "gongwen_standard")
        assert result == {
            "remove_numbering": True,
            "add_numbering": True,
            "numbering_scheme": "gongwen_standard",
        }

    def test_clean_keep_with_scheme(self) -> None:
        from docwen_cli.commands.execution_options import normalize_numbering_options

        result = normalize_numbering_options("keep", "cn_article")
        assert result == {
            "remove_numbering": False,
            "add_numbering": True,
            "numbering_scheme": "cn_article",
        }

    def test_invalid_clean_mode_raises(self) -> None:
        from docwen_cli.commands.execution_options import normalize_numbering_options

        with pytest.raises(ValueError, match="清理序号模式不合法"):
            normalize_numbering_options("invalid", None)

    def test_case_insensitive_clean_mode(self) -> None:
        from docwen_cli.commands.execution_options import normalize_numbering_options

        result = normalize_numbering_options("REMOVE", None)
        assert result["remove_numbering"] is True

    def test_case_insensitive_add_mode(self) -> None:
        from docwen_cli.commands.execution_options import normalize_numbering_options

        result = normalize_numbering_options(None, "CN_Article")
        assert result["add_numbering"] is True
        assert result["numbering_scheme"] == "cn_article"

    def test_add_mode_none_case_insensitive(self) -> None:
        from docwen_cli.commands.execution_options import normalize_numbering_options

        result = normalize_numbering_options(None, "NONE")
        assert result["add_numbering"] is False
        assert result["numbering_scheme"] == ""

    def test_clean_none_or_empty_falls_to_default(self) -> None:
        """None clean_mode defaults to 'default' (no-op) per old behaviour."""
        from docwen_cli.commands.execution_options import normalize_numbering_options

        # None → defaults to "default"
        result = normalize_numbering_options(None, "cn_article")
        assert result["remove_numbering"] is False  # "default" = no removal

    # ── render_mode parameter (Phase B-CLI-render) ─────────────────────

    def test_render_mode_explicit_word_native(self) -> None:
        """``--heading-numbering-render-mode word_native`` emits the key."""
        from docwen_cli.commands.execution_options import normalize_numbering_options

        result = normalize_numbering_options(
            None,
            "gongwen_standard",
            render_mode="word_native",
        )
        assert result["add_numbering"] is True
        assert result["numbering_scheme"] == "gongwen_standard"
        assert result["heading_numbering_render_mode"] == "word_native"

    def test_render_mode_explicit_text(self) -> None:
        """``--heading-numbering-render-mode text`` emits the key."""
        from docwen_cli.commands.execution_options import normalize_numbering_options

        result = normalize_numbering_options(
            "remove",
            None,
            render_mode="text",
        )
        assert result["heading_numbering_render_mode"] == "text"

    def test_render_mode_none_not_emitted(self) -> None:
        """When render_mode is not given, the optional key is absent."""
        from docwen_cli.commands.execution_options import normalize_numbering_options

        result = normalize_numbering_options("remove", "gongwen_standard")
        assert "heading_numbering_render_mode" not in result

    def test_render_mode_standalone_no_add(self) -> None:
        """render_mode alone (no add_numbering) is still valid and emitted."""
        from docwen_cli.commands.execution_options import normalize_numbering_options

        result = normalize_numbering_options(
            None,
            None,
            render_mode="word_native",
        )
        assert result == {
            "remove_numbering": False,
            "add_numbering": False,
            "numbering_scheme": "",
            "heading_numbering_render_mode": "word_native",
        }

    def test_render_mode_invalid_raises(self) -> None:
        from docwen_cli.commands.execution_options import normalize_numbering_options

        with pytest.raises(ValueError, match="渲染模式"):
            normalize_numbering_options(None, None, render_mode="bogus")

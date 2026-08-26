"""Tests for immutable export policies and pure resolution helpers."""

from __future__ import annotations

import pytest

from docwen_core.export_semantics import (
    VALID_LINK_STYLES,
    VALID_TABLE_MERGE_EXPORT_STRATEGIES,
    LinkRuntimeConfig,
    MarkdownExportSemantics,
    format_image_link,
    get_markdown_export_modes,
    normalize_markdown_break_separator,
    normalize_table_merge_export_strategy,
)

pytestmark = pytest.mark.unit

# ═══════════════════════════════════════════════════════════════════════════
# Dataclass construction
# ═══════════════════════════════════════════════════════════════════════════


class TestMarkdownExportSemanticsDefaults:
    """Verify the default-constructed ``MarkdownExportSemantics`` matches
    current Export-tab defaults."""

    def test_defaults_match_old_constants(self) -> None:
        s = MarkdownExportSemantics()
        assert s.image_link_style == "wiki_embed"
        assert s.md_file_link_style == "wiki_embed"
        assert s.ocr_blockquote_title_enabled is True
        assert s.ocr_blockquote_title_override_text == ""
        assert s.export_base64_compress_enabled is True
        assert s.export_base64_compress_threshold_kb == 100
        assert s.image_extraction_mode == "file"
        assert s.ocr_placement_mode == "main_md"
        assert s.table_merge_export_strategy == "fill"
        assert s.page_break_separator == "---"
        assert s.section_break_separator == "***"
        assert s.horizontal_rule_separator == "___"
        assert s.yaml_list_separator == "、"
        assert s.save_intermediate_files is False

    def test_frozen(self) -> None:
        s = MarkdownExportSemantics()
        with pytest.raises(Exception):  # noqa: B017 — frozen=True raises FrozenInstanceError; testing general contract
            s.image_link_style = "markdown_embed"  # type: ignore[misc]

    def test_from_config_respects_link_format(self) -> None:
        s = MarkdownExportSemantics.from_config(
            link_format={"image_link_style": "markdown_embed", "md_file_link_style": "wiki_link"},
        )
        assert s.image_link_style == "markdown_embed"
        assert s.md_file_link_style == "wiki_link"

    def test_from_config_ocr_override(self) -> None:
        s = MarkdownExportSemantics.from_config(
            ocr_output={
                "show_blockquote_title": True,
                "blockquote_title_override_by_locale": {"zh_CN": "OCR 结果"},
            },
        )
        assert s.ocr_blockquote_title_enabled is True
        assert s.ocr_blockquote_title_override_text == "OCR 结果"

    def test_from_config_ocr_override_uses_requested_locale(self) -> None:
        s = MarkdownExportSemantics.from_config(
            ocr_output={
                "show_blockquote_title": True,
                "blockquote_title_override_by_locale": {
                    "zh_CN": "中文标题",
                    "en_US": "English title",
                },
            },
            locale="en_US",
        )
        assert s.ocr_blockquote_title_override_text == "English title"

    def test_from_config_ocr_disabled(self) -> None:
        s = MarkdownExportSemantics.from_config(
            ocr_output={"show_blockquote_title": False},
        )
        assert s.ocr_blockquote_title_enabled is False

    def test_from_config_without_mode_config_uses_deterministic_defaults(self) -> None:
        s = MarkdownExportSemantics.from_config()
        assert s.image_extraction_mode == "file"
        assert s.ocr_placement_mode == "main_md"

    def test_from_config_intermediate_files(self) -> None:
        s = MarkdownExportSemantics.from_config(
            intermediate_files_cfg={"save_to_output": True},
        )
        assert s.save_intermediate_files is True

    def test_from_config_ignores_removed_top_level_base64_fields(self) -> None:
        s = MarkdownExportSemantics.from_config(
            export_cfg={"base64_compress_enabled": False, "base64_compress_threshold_kb": 500},
        )
        assert s.export_base64_compress_enabled is True
        assert s.export_base64_compress_threshold_kb == 100

    def test_from_config_uses_conversion_export_for_base64_without_mixing_mode_owner(self) -> None:
        s = MarkdownExportSemantics.from_config(
            export_cfg={
                "to_md_image_extraction_mode": "base64",
                "to_md_ocr_placement_mode": "main_md",
                "base64_compress_enabled": True,
                "base64_compress_threshold_kb": 999,
            },
            conversion_cfg={
                "export": {
                    "base64_compress_enabled": False,
                    "base64_compress_threshold_kb": 37,
                }
            },
        )

        assert s.image_extraction_mode == "base64"
        assert s.ocr_placement_mode == "main_md"
        assert s.export_base64_compress_enabled is False
        assert s.export_base64_compress_threshold_kb == 37

    def test_from_config_respects_export_to_markdown_modes(self) -> None:
        s = MarkdownExportSemantics.from_config(
            export_cfg={
                "to_md_image_extraction_mode": "base64",
                "to_md_ocr_placement_mode": "main_md",
            },
            conversion_cfg={
                "image_extraction_mode": "file",
                "ocr_placement_mode": "image_md",
            },
        )

        assert s.image_extraction_mode == "base64"
        assert s.ocr_placement_mode == "main_md"

    def test_from_config_ignores_removed_conversion_mode_fields(self) -> None:
        s = MarkdownExportSemantics.from_config(
            conversion_cfg={
                "image_extraction_mode": "base64",
                "ocr_placement_mode": "image_md",
            },
        )

        assert s.image_extraction_mode == "file"
        assert s.ocr_placement_mode == "main_md"

    @pytest.mark.parametrize("separator", ["，", ", ", ""])
    def test_from_config_yaml_list_separator(self, separator: str) -> None:
        s = MarkdownExportSemantics.from_config(
            conversion_cfg={"md_to_docx": {"list_separator": separator}},
        )
        assert s.yaml_list_separator == separator

    def test_from_config_ignores_removed_flat_yaml_list_separator(self) -> None:
        s = MarkdownExportSemantics.from_config(conversion_cfg={"list_separator": ";"})

        assert s.yaml_list_separator == "、"

    def test_from_config_table_merge_strategy(self) -> None:
        s = MarkdownExportSemantics.from_config(
            conversion_cfg={"table_merge_export_strategy": "empty"},
        )
        assert s.table_merge_export_strategy == "empty"

    def test_from_config_docx_to_markdown_break_separators(self) -> None:
        s = MarkdownExportSemantics.from_config(
            conversion_cfg={
                "horizontal_rule": {
                    "docx_to_md": {
                        "page_break": "***",
                        "section_break": "___",
                        "horizontal_rule": "---",
                    },
                },
            },
        )

        assert s.page_break_separator == "***"
        assert s.section_break_separator == "___"
        assert s.horizontal_rule_separator == "---"

    def test_from_config_docx_to_markdown_break_ignore(self) -> None:
        s = MarkdownExportSemantics.from_config(
            conversion_cfg={
                "horizontal_rule": {
                    "docx_to_md": {
                        "page_break": "ignore",
                        "section_break": "ignore",
                        "horizontal_rule": "ignore",
                    },
                },
            },
        )

        assert s.page_break_separator == ""
        assert s.section_break_separator == ""
        assert s.horizontal_rule_separator == ""

    def test_from_config_invalid_link_style_falls_back(self) -> None:
        """Unknown link style strings are normalized to ``wiki_embed``."""
        s = MarkdownExportSemantics.from_config(
            link_format={"image_link_style": "garbage"},
        )
        assert s.image_link_style == "wiki_embed"


class TestLinkRuntimeConfig:
    """``LinkRuntimeConfig`` is complete and immutable."""

    def test_defaults_match_old_constants(self) -> None:
        cfg = LinkRuntimeConfig()
        assert cfg.max_depth == 3
        assert cfg.non_embed_wiki_mode == "hyperlink"
        assert cfg.non_embed_markdown_mode == "hyperlink"
        assert cfg.embed_wiki_image_mode == "embed"
        assert cfg.embed_markdown_image_mode == "embed"
        assert cfg.embed_md_file_mode == "embed"
        assert cfg.search_dirs == (".", "assets", "images", "attachments")
        assert cfg.detect_circular is True
        assert cfg.file_not_found_mode == "placeholder"
        assert cfg.circular_reference_mode == "placeholder"
        assert cfg.max_depth_reached_mode == "placeholder"
        assert cfg.auto_link_bare_url is False

    def test_frozen(self) -> None:
        cfg = LinkRuntimeConfig()
        with pytest.raises(Exception):  # noqa: B017 — frozen=True raises FrozenInstanceError; testing general contract
            cfg.max_depth = 99  # type: ignore[misc]

    def test_from_config_respects_all_sections(self) -> None:
        cfg = LinkRuntimeConfig.from_config(
            {
                "non_embed_links": {
                    "wiki_mode": "keep",
                    "markdown_mode": "keep",
                    "auto_link_bare_url": True,
                },
                "embed_links": {
                    "wiki_image_mode": "placeholder",
                    "markdown_image_mode": "placeholder",
                    "md_file_mode": "keep",
                },
                "embedding": {"max_depth": 5},
                "path_resolution": {"search_dirs": ["files", "docs"]},
                "error_handling": {
                    "file_not_found": "keep",
                    "detect_circular": False,
                    "circular_reference": "ignore",
                    "max_depth_reached": "keep",
                },
            }
        )
        assert cfg.max_depth == 5
        assert cfg.non_embed_wiki_mode == "keep"
        assert cfg.auto_link_bare_url is True
        assert cfg.embed_wiki_image_mode == "placeholder"
        assert cfg.embed_md_file_mode == "keep"
        assert cfg.search_dirs == ("files", "docs")
        assert cfg.detect_circular is False
        assert cfg.file_not_found_mode == "keep"
        assert cfg.circular_reference_mode == "ignore"
        assert cfg.max_depth_reached_mode == "keep"

    def test_from_config_missing_keys_use_defaults(self) -> None:
        cfg = LinkRuntimeConfig.from_config({})
        assert cfg.max_depth == 3
        assert cfg.detect_circular is True


# ═══════════════════════════════════════════════════════════════════════════
# Export modes
# ═══════════════════════════════════════════════════════════════════════════


class TestGetMarkdownExportModes:
    def test_default_modes(self) -> None:
        modes = get_markdown_export_modes("docx", semantics=MarkdownExportSemantics())
        assert modes["image_extraction_mode"] == "file"
        assert modes["ocr_placement_mode"] == "main_md"
        assert modes["table_merge_export_strategy"] == "fill"

    def test_kind_is_route_identity_not_configuration_precedence(self) -> None:
        semantics = MarkdownExportSemantics(
            image_extraction_mode="base64",
            ocr_placement_mode="image_md",
        )

        for kind in ("docx", "xlsx", "image", "pdf", "epub", "pptx"):
            modes = get_markdown_export_modes(kind, semantics=semantics)
            assert modes["image_extraction_mode"] == "base64"
            assert modes["ocr_placement_mode"] == "main_md"

    def test_call_overrides_take_precedence(self) -> None:
        modes = get_markdown_export_modes(
            "docx",
            extraction_mode="base64",
            table_merge_export_strategy="marker",
            semantics=MarkdownExportSemantics(),
        )
        assert modes["image_extraction_mode"] == "base64"
        assert modes["table_merge_export_strategy"] == "marker"

    def test_base64_image_mode_forces_main_md_ocr_placement(self) -> None:
        modes = get_markdown_export_modes(
            "docx",
            extraction_mode="base64",
            ocr_placement_mode="image_md",
            semantics=MarkdownExportSemantics(ocr_placement_mode="image_md"),
        )
        assert modes["image_extraction_mode"] == "base64"
        assert modes["ocr_placement_mode"] == "main_md"

    def test_invalid_strategy_is_normalized(self) -> None:
        modes = get_markdown_export_modes(
            "xlsx",
            table_merge_export_strategy="garbage",
            semantics=MarkdownExportSemantics(),
        )
        assert modes["table_merge_export_strategy"] in VALID_TABLE_MERGE_EXPORT_STRATEGIES

    def test_removed_replicate_alias_uses_current_default(self) -> None:
        modes = get_markdown_export_modes(
            "xlsx",
            table_merge_export_strategy="replicate",
            semantics=MarkdownExportSemantics(table_merge_export_strategy="empty"),
        )
        assert modes["table_merge_export_strategy"] == "empty"


# ═══════════════════════════════════════════════════════════════════════════
# Utility functions
# ═══════════════════════════════════════════════════════════════════════════


class TestNormalizeTableMergeExportStrategy:
    def test_valid_values_passthrough(self) -> None:
        for strategy in ("fill", "empty", "marker"):
            assert normalize_table_merge_export_strategy(strategy) == strategy

    def test_removed_replicate_alias_is_invalid(self) -> None:
        assert normalize_table_merge_export_strategy("replicate") == "fill"
        assert normalize_table_merge_export_strategy(" REPLICATE ", default_strategy="empty") == "empty"

    def test_none_returns_default(self) -> None:
        assert normalize_table_merge_export_strategy(None) == "fill"
        assert normalize_table_merge_export_strategy(None, default_strategy="empty") == "empty"

    def test_invalid_returns_default(self) -> None:
        assert normalize_table_merge_export_strategy("invalid") == "fill"
        assert normalize_table_merge_export_strategy("merge") == "fill"

    def test_case_and_whitespace_insensitive(self) -> None:
        assert normalize_table_merge_export_strategy("  FILL  ") == "fill"
        assert normalize_table_merge_export_strategy("Marker") == "marker"

    def test_invalid_default_falls_back_to_fill(self) -> None:
        assert normalize_table_merge_export_strategy(None, default_strategy="bad") == "fill"


class TestNormalizeMarkdownBreakSeparator:
    def test_none_and_empty_return_default(self) -> None:
        assert normalize_markdown_break_separator(None, default="---") == "---"
        assert normalize_markdown_break_separator("  ", default="***") == "***"

    def test_ignore_returns_empty_string(self) -> None:
        assert normalize_markdown_break_separator("ignore", default="---") == ""
        assert normalize_markdown_break_separator(" IGNORE ", default="---") == ""

    def test_separator_passthrough(self) -> None:
        assert normalize_markdown_break_separator("___", default="---") == "___"


class TestFormatImageLink:
    def test_wiki_embed(self) -> None:
        link = format_image_link("alt", "photo.png", style="wiki_embed")
        assert link == "![[photo.png]]"

    def test_wiki_link(self) -> None:
        link = format_image_link("alt", "photo.png", style="wiki_link")
        assert link == "[[photo.png]]"

    def test_markdown_embed(self) -> None:
        link = format_image_link("alt", "photo.png", style="markdown_embed")
        assert link == "![alt](photo.png)"

    def test_markdown_link(self) -> None:
        link = format_image_link("alt", "photo.png", style="markdown_link")
        assert link == "[alt](photo.png)"

    def test_default_style_is_wiki_embed(self) -> None:
        link = format_image_link("alt", "photo.png")
        assert link == "![[photo.png]]"

    def test_invalid_style_falls_back_to_wiki_embed(self) -> None:
        link = format_image_link("alt", "photo.png", style="garbage")
        assert link == "![[photo.png]]"

    def test_all_valid_styles_produce_output(self) -> None:
        for style in VALID_LINK_STYLES:
            link = format_image_link("test", "file.png", style=style)
            assert len(link) > 0
            assert "file.png" in link

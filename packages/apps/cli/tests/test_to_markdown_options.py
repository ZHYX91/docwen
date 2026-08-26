"""Tests for canonical to-markdown option builder / normalizer."""

from __future__ import annotations

import pytest

from docwen_cli.options.to_markdown import (
    CANONICAL_KEYS,
    build_to_markdown_options,
    normalize_to_markdown_options,
)

pytestmark = pytest.mark.unit

# ── build_to_markdown_options ────────────────────────────────────────────


class TestBuildToMarkdownOptions:
    def test_empty_returns_empty_dict(self):
        result = build_to_markdown_options()
        assert result == {}

    def test_keep_images_true(self):
        result = build_to_markdown_options(keep_images=True)
        assert result == {"to_md_keep_images": True}

    def test_keep_images_false(self):
        result = build_to_markdown_options(keep_images=False)
        assert result == {"to_md_keep_images": False}

    def test_enable_ocr(self):
        result = build_to_markdown_options(enable_ocr=True)
        assert result == {"to_md_enable_ocr": True}

    def test_image_mode_valid(self):
        result = build_to_markdown_options(image_mode="file")
        assert result == {"image_mode": "file"}

    def test_image_mode_rejects_different_case(self):
        with pytest.raises(ValueError, match="image_mode"):
            build_to_markdown_options(image_mode="BASE64")

    def test_image_mode_invalid_rejected(self):
        with pytest.raises(ValueError, match="image_mode"):
            build_to_markdown_options(image_mode="bogus")

    def test_ocr_placement_main_md(self):
        result = build_to_markdown_options(ocr_placement="main_md")
        assert result == {"ocr_placement": "main_md"}

    def test_ocr_placement_rejects_different_case(self):
        with pytest.raises(ValueError, match="ocr_placement"):
            build_to_markdown_options(ocr_placement="IMAGE_MD")

    def test_ocr_language_rejects_different_case(self):
        with pytest.raises(ValueError, match="ocr_language"):
            build_to_markdown_options(ocr_language="Japanese")

    def test_ocr_language_invalid_rejected(self):
        with pytest.raises(ValueError, match="ocr_language"):
            build_to_markdown_options(ocr_language="elvish")

    def test_image_link_style(self):
        result = build_to_markdown_options(image_link_style="markdown_link")
        assert result == {"image_link_style": "markdown_link"}

    def test_image_link_style_rejects_different_case(self):
        with pytest.raises(ValueError, match="image_link_style"):
            build_to_markdown_options(image_link_style="WIKI_EMBED")

    def test_table_merge_strategy(self):
        result = build_to_markdown_options(table_merge_strategy="marker")
        assert result == {"table_merge_strategy": "marker"}

    def test_removed_table_merge_alias_is_rejected(self):
        with pytest.raises(ValueError, match="table_merge_strategy"):
            build_to_markdown_options(table_merge_strategy="REPLICATE")

    def test_table_merge_strategy_invalid_rejected(self):
        with pytest.raises(ValueError, match="table_merge_strategy"):
            build_to_markdown_options(table_merge_strategy="preserve")

    def test_all_keys_together(self):
        result = build_to_markdown_options(
            keep_images=True,
            enable_ocr=False,
            image_mode="embed",
            ocr_placement="image_md",
            ocr_language="korean",
            image_link_style="wiki_link",
            table_merge_strategy="empty",
        )
        assert result == {
            "to_md_keep_images": True,
            "to_md_enable_ocr": False,
            "image_mode": "embed",
            "ocr_placement": "image_md",
            "ocr_language": "korean",
            "image_link_style": "wiki_link",
            "table_merge_strategy": "empty",
        }

    def test_none_values_omitted(self):
        """None values are omitted so plugin defaults apply."""
        result = build_to_markdown_options(
            keep_images=True,
            enable_ocr=None,
            image_mode=None,
        )
        assert "to_md_enable_ocr" not in result
        assert "image_mode" not in result
        assert result == {"to_md_keep_images": True}

    def test_no_legacy_keys_in_output(self):
        """Builder output must never contain legacy/alias keys."""
        result = build_to_markdown_options(
            keep_images=True,
            enable_ocr=True,
            image_mode="file",
            ocr_placement="main_md",
            ocr_language="latin",
            image_link_style="wiki_embed",
            table_merge_strategy="fill",
        )
        legacy_keys = {
            "to_md_extract_image",
            "extract_image",
            "extract_ocr",
        }
        found_legacy = legacy_keys & set(result.keys())
        assert not found_legacy, f"Legacy keys found: {found_legacy}"

    def test_output_keys_are_canonical(self):
        """Every output key must be in the canonical key set."""
        result = build_to_markdown_options(
            keep_images=True,
            enable_ocr=True,
            image_mode="file",
            ocr_placement="main_md",
            ocr_language="latin",
            image_link_style="wiki_embed",
            table_merge_strategy="fill",
        )
        for key in result:
            assert key in CANONICAL_KEYS, f"Key {key!r} not in CANONICAL_KEYS"


# ── normalize_to_markdown_options ────────────────────────────────────────


class TestNormalizeToMarkdownOptions:
    def test_passes_through_canonical_keys(self):
        raw = {
            "to_md_keep_images": True,
            "to_md_enable_ocr": False,
            "image_mode": "file",
            "ocr_placement": "main_md",
            "ocr_language": "english",
            "image_link_style": "markdown_link",
            "table_merge_strategy": "marker",
        }
        result = normalize_to_markdown_options(raw)
        assert result == raw

    def test_rejects_unknown_keys(self):
        raw = {
            "to_md_keep_images": True,
            "extract_image": True,  # legacy — should be dropped
            "extract_ocr": True,  # legacy — should be dropped
            "garbage": 123,
        }
        with pytest.raises(ValueError, match="extract_image"):
            normalize_to_markdown_options(raw)

    def test_rejects_different_key_case(self):
        raw = {
            "TO_MD_KEEP_IMAGES": True,
            "IMAGE_MODE": "embed",
        }
        with pytest.raises(ValueError, match="TO_MD_KEEP_IMAGES"):
            normalize_to_markdown_options(raw)

    def test_rejects_different_value_case(self):
        raw: dict[str, object] = {"ocr_placement": "MAIN_MD"}
        with pytest.raises(ValueError, match="ocr_placement"):
            normalize_to_markdown_options(raw)

    def test_rejects_different_ocr_language_case(self):
        raw: dict[str, object] = {"ocr_language": "CHINESE_CHT"}
        with pytest.raises(ValueError, match="ocr_language"):
            normalize_to_markdown_options(raw)

    def test_invalid_value_rejected(self):
        raw: dict[str, object] = {"ocr_placement": "sidecar"}
        with pytest.raises(ValueError, match="ocr_placement"):
            normalize_to_markdown_options(raw)

    def test_removed_table_merge_alias_is_rejected(self):
        raw: dict[str, object] = {"table_merge_strategy": "REPLICATE"}
        with pytest.raises(ValueError, match="table_merge_strategy"):
            normalize_to_markdown_options(raw)

    def test_table_merge_strategy_invalid_rejected(self):
        raw: dict[str, object] = {"table_merge_strategy": "keep"}
        with pytest.raises(ValueError, match="table_merge_strategy"):
            normalize_to_markdown_options(raw)

    def test_bool_coercion_is_rejected(self):
        raw: dict[str, object] = {"to_md_keep_images": 1, "to_md_enable_ocr": 0}
        with pytest.raises(ValueError, match="to_md_keep_images"):
            normalize_to_markdown_options(raw)

    def test_empty_dict(self):
        assert normalize_to_markdown_options({}) == {}

    def test_none_values_for_bool_keys_are_rejected(self):
        raw: dict[str, object] = {"to_md_keep_images": None}
        with pytest.raises(ValueError, match="to_md_keep_images"):
            normalize_to_markdown_options(raw)

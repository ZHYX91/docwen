"""Tests for shared conversion option schemas."""

from __future__ import annotations

import tomllib
from pathlib import Path
from typing import TypeGuard

import pytest

from docwen_core.options import (
    COMMON_OPTIONS_SCHEMA,
    DOCX_OPTIONS_SCHEMA,
    IMAGE_OPTIONS_SCHEMA,
    MARKDOWN_OPTIONS_SCHEMA,
    PDF_OPTIONS_SCHEMA,
    PROOFREAD_OPTIONS_SCHEMA,
)
from docwen_core.text.heading_merge import DEFAULT_HEADING_MERGE_PUNCTUATION

pytestmark = pytest.mark.unit


def _is_object_properties(value: object) -> TypeGuard[dict[str, dict[str, object]]]:
    if not isinstance(value, dict) or not value:
        return False
    return all(
        isinstance(name, str) and isinstance(spec, dict) and all(isinstance(spec_key, str) for spec_key in spec)
        for name, spec in value.items()
    )


def _assert_object_schema(schema: dict[str, object]) -> dict[str, dict[str, object]]:
    assert schema["type"] == "object"
    properties = schema["properties"]
    assert _is_object_properties(properties)
    return properties


def _assert_string_enum(spec: dict[str, object]) -> list[str]:
    raw_values = spec["enum"]
    assert isinstance(raw_values, list)
    values: list[str] = []
    for value in raw_values:
        assert isinstance(value, str)
        values.append(value)
    return values


def test_shared_option_schemas_are_json_schema_objects() -> None:
    for schema in (
        COMMON_OPTIONS_SCHEMA,
        MARKDOWN_OPTIONS_SCHEMA,
        DOCX_OPTIONS_SCHEMA,
        PDF_OPTIONS_SCHEMA,
        IMAGE_OPTIONS_SCHEMA,
        PROOFREAD_OPTIONS_SCHEMA,
    ):
        properties = _assert_object_schema(schema)
        for spec in properties.values():
            assert isinstance(spec, dict)
            assert "type" in spec


def test_markdown_formatting_schema_matches_current_config_surface() -> None:
    properties = _assert_object_schema(MARKDOWN_OPTIONS_SCHEMA)

    formatting_modes = _assert_string_enum(properties["formatting_mode"])
    assert formatting_modes == ["apply", "keep", "remove"]
    assert properties["formatting_mode"]["default"] == "apply"
    assert properties["heading_formatting_mode"]["enum"] == ["apply", "keep", "remove"]
    assert properties["heading_formatting_mode"]["default"] == "remove"
    assert properties["table_header_formatting_mode"]["enum"] == ["apply", "keep", "remove"]
    assert properties["table_header_formatting_mode"]["default"] == "remove"
    assert properties["heading_merge_mode"]["enum"] == ["punct_required", "always", "never"]
    assert properties["heading_merge_punctuation"]["default"] == "。：！？.:!?"
    assert "ignore" not in formatting_modes


def test_heading_merge_punctuation_default_matches_static_conversion_config() -> None:
    project_root = Path(__file__).resolve().parents[3]
    conversion = tomllib.loads((project_root / "configs" / "conversion.toml").read_text(encoding="utf-8"))

    assert MARKDOWN_OPTIONS_SCHEMA["properties"]["heading_merge_punctuation"]["default"] == (
        DEFAULT_HEADING_MERGE_PUNCTUATION
    )
    assert conversion["md_to_docx"]["heading_merge_punctuation"] == DEFAULT_HEADING_MERGE_PUNCTUATION


def test_common_ocr_placement_schema_stays_on_public_values() -> None:
    properties = _assert_object_schema(COMMON_OPTIONS_SCHEMA)

    assert properties["ocr_placement"]["enum"] == ["image_md", "main_md"]
    assert properties["ocr_placement"]["default"] == "main_md"


def test_image_mode_schemas_match_current_to_markdown_surface() -> None:
    expected_modes = ["file", "base64", "embed", "omit"]

    common_properties = _assert_object_schema(COMMON_OPTIONS_SCHEMA)
    docx_properties = _assert_object_schema(DOCX_OPTIONS_SCHEMA)

    assert common_properties["image_mode"]["enum"] == expected_modes
    assert common_properties["image_mode"]["default"] == "file"
    assert docx_properties["image_mode"]["enum"] == expected_modes
    assert docx_properties["image_mode"]["default"] == "file"


def test_common_to_markdown_schema_tracks_current_shared_surface() -> None:
    expected_properties = {
        "to_md_keep_images",
        "to_md_enable_ocr",
        "image_mode",
        "ocr_placement",
        "ocr_language",
        "image_link_style",
        "table_merge_strategy",
    }
    properties = _assert_object_schema(COMMON_OPTIONS_SCHEMA)

    assert expected_properties <= set(properties)
    assert properties["ocr_language"]["enum"] == [
        "auto",
        "chinese",
        "chinese_cht",
        "english",
        "japanese",
        "korean",
        "latin",
        "cyrillic",
    ]
    assert properties["ocr_language"]["default"] == "auto"
    assert properties["image_link_style"]["enum"] == [
        "wiki_embed",
        "wiki_link",
        "markdown_embed",
        "markdown_link",
    ]
    assert properties["image_link_style"]["default"] == "wiki_embed"
    assert properties["table_merge_strategy"]["enum"] == ["fill", "empty", "marker"]
    assert properties["table_merge_strategy"]["default"] == "fill"


def test_docx_schema_tracks_current_to_markdown_shared_surface() -> None:
    properties = _assert_object_schema(DOCX_OPTIONS_SCHEMA)

    assert properties["ocr_placement"]["enum"] == ["image_md", "main_md"]
    assert properties["ocr_placement"]["default"] == "main_md"
    assert properties["ocr_language"]["enum"] == [
        "auto",
        "chinese",
        "chinese_cht",
        "english",
        "japanese",
        "korean",
        "latin",
        "cyrillic",
    ]
    assert properties["ocr_language"]["default"] == "auto"
    assert properties["image_link_style"]["enum"] == [
        "wiki_embed",
        "wiki_link",
        "markdown_embed",
        "markdown_link",
    ]
    assert properties["image_link_style"]["default"] == "wiki_embed"
    assert properties["table_merge_strategy"]["enum"] == ["fill", "empty", "marker"]
    assert properties["table_merge_strategy"]["default"] == "fill"


def test_proofread_schema_tracks_plugin_request_surface() -> None:
    properties = _assert_object_schema(PROOFREAD_OPTIONS_SCHEMA)

    assert set(properties) == {
        "enable_symbol_pairing",
        "enable_symbol_correction",
        "enable_typos_rule",
        "enable_sensitive_word",
        "skip_code_blocks",
        "skip_quote_blocks",
    }
    assert all(spec["type"] == "boolean" for spec in properties.values())
    assert properties["enable_symbol_pairing"]["default"] is True
    assert properties["enable_symbol_correction"]["default"] is True
    assert properties["enable_typos_rule"]["default"] is True
    assert properties["enable_sensitive_word"]["default"] is True
    assert properties["skip_code_blocks"]["default"] is True
    assert properties["skip_quote_blocks"]["default"] is False
    assert "SYMBOL_PAIRING" not in properties
    assert "symbol_pairing" not in properties

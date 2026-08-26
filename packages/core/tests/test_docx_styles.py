"""Contract tests for the stable 0.9 DOCX style registry."""

from __future__ import annotations

from dataclasses import FrozenInstanceError

import pytest

from docwen_core.docx_styles import (
    BUILTIN_DOCUMENT_STYLES,
    CUSTOM_DOCUMENT_STYLE_KEYS,
    CUSTOM_DOCUMENT_STYLES,
    MANAGED_DOCUMENT_STYLES,
    DocumentStyleCatalog,
    DocumentStyleFormat,
)

pytestmark = pytest.mark.contract


def test_registry_has_exact_stable_identity_counts() -> None:
    assert len(BUILTIN_DOCUMENT_STYLES) == 16
    assert len(CUSTOM_DOCUMENT_STYLES) == 27
    assert len(MANAGED_DOCUMENT_STYLES) == 43
    assert len({item.semantic_key for item in MANAGED_DOCUMENT_STYLES}) == 43
    assert len({item.style_id for item in MANAGED_DOCUMENT_STYLES}) == 43
    assert tuple(item.semantic_key for item in CUSTOM_DOCUMENT_STYLES) == CUSTOM_DOCUMENT_STYLE_KEYS


def test_registry_types_and_bases_are_closed() -> None:
    valid_bases = {"Normal", "DefaultParagraphFont", "TableNormal", "Caption"}
    assert {item.kind for item in MANAGED_DOCUMENT_STYLES} == {"paragraph", "character", "table"}
    assert all(item.based_on in valid_bases for item in MANAGED_DOCUMENT_STYLES)
    assert all((item.canonical_name is None) != (item.locale_name_key is None) for item in MANAGED_DOCUMENT_STYLES)


def test_catalog_is_immutable_and_lookups_fail_closed() -> None:
    style_format = DocumentStyleFormat(
        east_asia_font="SimSun",
        ascii_font="Arial",
        font_size_pt=12.0,
        first_line_indent_chars=0,
        first_line_indent_cm=0.0,
        spacing_after_twip=120,
        spacing_before_twip=0,
        bold=False,
        justification="left",
    )
    catalog = DocumentStyleCatalog(
        locale="en_US",
        output_names=(("body_paragraph", "Body Paragraph"),),
        recognition_names=(("body_paragraph", ("Body Paragraph", "正文段落")),),
        formats=(("body_paragraph", style_format),),
    )
    assert catalog.name_for("body_paragraph") == "Body Paragraph"
    assert catalog.recognition_names_for("body_paragraph") == ("Body Paragraph", "正文段落")
    assert catalog.format_for("body_paragraph") is style_format
    assert catalog.format_for("missing") is None
    with pytest.raises(KeyError):
        catalog.name_for("missing")
    with pytest.raises(FrozenInstanceError):
        catalog.locale = "zh_CN"  # type: ignore[misc]

"""Strict request-owned DOCX style catalog projection."""

from __future__ import annotations

import shutil
from pathlib import Path
from typing import Any

import pytest
import tomlkit

from docwen_core.docx_styles import CUSTOM_DOCUMENT_STYLE_KEYS, SHIPPED_STYLE_LOCALES
from docwen_runtime.config import DocumentStyleCatalogError, build_document_style_catalog

pytestmark = pytest.mark.contract

ROOT = Path(__file__).resolve().parents[3]
LOCALES = ROOT / "i18n" / "locales"


@pytest.mark.parametrize("locale", SHIPPED_STYLE_LOCALES)
def test_all_shipped_locales_project_exact_catalog(locale: str) -> None:
    catalog = build_document_style_catalog(
        {"gui": {"language": {"locale": "zh_CN"}}},
        request_options={"locale": locale},
        locales_dir=LOCALES,
    )
    assert catalog.locale == locale
    assert tuple(key for key, _name in catalog.output_names) == CUSTOM_DOCUMENT_STYLE_KEYS
    assert all(catalog.name_for(key).strip() for key in CUSTOM_DOCUMENT_STYLE_KEYS)
    assert catalog.format_for("heading_3") == catalog.format_for("heading_9")
    assert len(catalog.formats) == 10


def test_request_locale_precedes_snapshot_locale() -> None:
    catalog = build_document_style_catalog(
        {"gui": {"language": {"locale": "zh_CN"}}},
        request_options={"locale": "en_US"},
        locales_dir=LOCALES,
    )
    assert catalog.locale == "en_US"
    assert catalog.name_for("figure_caption") == "Figure Caption"


@pytest.mark.parametrize("value", ["", None, "../en_US", r"C:\\en_US"])
def test_explicit_invalid_locale_never_falls_back(value: object) -> None:
    with pytest.raises(DocumentStyleCatalogError) as captured:
        build_document_style_catalog({}, request_options={"locale": value}, locales_dir=LOCALES)
    assert captured.value.diagnostic_code == "DOCX_STYLE_LOCALE_INVALID"


def test_missing_selected_locale_never_falls_back(tmp_path: Path) -> None:
    for locale in SHIPPED_STYLE_LOCALES:
        if locale != "en_US":
            shutil.copy2(LOCALES / f"{locale}.toml", tmp_path / f"{locale}.toml")
    with pytest.raises(DocumentStyleCatalogError) as captured:
        build_document_style_catalog({}, request_options={"locale": "en_US"}, locales_dir=tmp_path)
    assert captured.value.diagnostic_code == "DOCX_STYLE_LOCALE_MISSING"
    assert captured.value.error_type == "conversion_failed"


def test_malformed_toml_has_distinct_resource_failure(tmp_path: Path) -> None:
    for locale in SHIPPED_STYLE_LOCALES:
        shutil.copy2(LOCALES / f"{locale}.toml", tmp_path / f"{locale}.toml")
    (tmp_path / "en_US.toml").write_text("[styles\n", encoding="utf-8")
    with pytest.raises(DocumentStyleCatalogError) as captured:
        build_document_style_catalog({}, request_options={"locale": "en_US"}, locales_dir=tmp_path)
    assert captured.value.diagnostic_code == "DOCX_STYLE_LOCALE_INVALID_TOML"
    assert captured.value.error_type == "conversion_failed"


def test_locale_table_order_is_not_part_of_the_contract(tmp_path: Path) -> None:
    for locale in SHIPPED_STYLE_LOCALES:
        source = LOCALES / f"{locale}.toml"
        if locale != "en_US":
            shutil.copy2(source, tmp_path / source.name)
            continue
        document: Any = tomlkit.parse(source.read_text(encoding="utf-8"))
        styles: Any = document["styles"]
        first = CUSTOM_DOCUMENT_STYLE_KEYS[0]
        value = styles.pop(first)
        styles.add(first, value)
        formats: Any = document["style_formats"]
        format_value = formats.pop("body_paragraph")
        formats.add("body_paragraph", format_value)
        (tmp_path / source.name).write_text(document.as_string(), encoding="utf-8")
    catalog = build_document_style_catalog({}, request_options={"locale": "en_US"}, locales_dir=tmp_path)
    assert tuple(key for key, _name in catalog.output_names) == CUSTOM_DOCUMENT_STYLE_KEYS


def test_blank_name_in_any_recognition_locale_is_rejected(tmp_path: Path) -> None:
    for locale in SHIPPED_STYLE_LOCALES:
        text = (LOCALES / f"{locale}.toml").read_text(encoding="utf-8")
        if locale == "de_DE":
            text = text.replace('figure_caption = "Abbildungsbeschriftung"', 'figure_caption = ""')
        (tmp_path / f"{locale}.toml").write_text(text, encoding="utf-8")
    with pytest.raises(DocumentStyleCatalogError) as captured:
        build_document_style_catalog({}, request_options={"locale": "en_US"}, locales_dir=tmp_path)
    assert captured.value.diagnostic_code == "DOCX_STYLE_NAME_BLANK"


def test_recognition_names_cover_all_locales_without_duplicates() -> None:
    catalog = build_document_style_catalog({}, locales_dir=LOCALES)
    names = catalog.recognition_names_for("body_paragraph")
    assert "正文段落" in names
    assert "Body Paragraph" in names
    assert len(names) == len({name.casefold() for name in names})

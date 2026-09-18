"""Regression coverage for request-owned DOCX foundation-style completion."""

from __future__ import annotations

import pytest
from docx import Document
from docx.oxml.ns import qn

from docwen_plugin_markdown.to_docx.managed_styles import complete_managed_styles
from docwen_runtime.config.document_styles import build_document_style_catalog

from .conftest import PROJECT_ROOT

pytestmark = pytest.mark.unit


def _catalog():
    return build_document_style_catalog(
        {"gui": {"language": {"locale": "zh_CN"}}},
        locales_dir=PROJECT_ROOT / "i18n" / "locales",
    )


def _remove_styles(document, *style_ids: str) -> None:
    root = document.styles.element
    wanted = set(style_ids)
    for element in list(root.findall(qn("w:style"))):
        if element.get(qn("w:styleId"), "") in wanted:
            root.remove(element)


def _style_ids(document) -> set[str]:
    return {style.style_id for style in document.styles}


@pytest.mark.unit
def test_missing_foundations_are_completed_but_title_is_not_required_or_injected() -> None:
    document = Document()
    _remove_styles(document, "Normal", "DefaultParagraphFont", "TableNormal", "Title")

    completed, bindings = complete_managed_styles(document, _catalog())
    style_ids = _style_ids(completed)

    assert {"Normal", "DefaultParagraphFont", "TableNormal"}.issubset(style_ids)
    assert "Title" not in style_ids
    assert len(bindings.styles) == 43


def test_existing_title_is_preserved_without_becoming_a_managed_dependency() -> None:
    document = Document()
    title = document.styles["Title"]
    original_id = title.style_id

    completed, _bindings = complete_managed_styles(document, _catalog())

    assert completed.styles[original_id].style_id == original_id

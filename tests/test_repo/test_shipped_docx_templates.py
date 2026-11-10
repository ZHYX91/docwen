"""Live assets cover every GUI locale without relying on historical snapshots."""

import tomllib
from pathlib import Path

import pytest
from docx import Document
from docx.oxml.ns import qn

from docwen_core.docx_styles import MANAGED_DOCUMENT_STYLES, SHIPPED_STYLE_LOCALES
from docwen_plugin_markdown.to_docx.managed_styles import complete_managed_styles
from docwen_runtime.config.document_styles import build_document_style_catalog

pytestmark = pytest.mark.contract
ROOT = Path(__file__).resolve().parents[2]


@pytest.mark.parametrize("locale", SHIPPED_STYLE_LOCALES)
def test_shipped_template_contains_complete_localized_styles_and_placeholders(locale):
    data = tomllib.loads((ROOT / "i18n/locales" / f"{locale}.toml").read_text(encoding="utf-8"))
    document = Document(ROOT / "templates" / (data["meta"]["template_name"] + ".docx"))
    assert [p.text for p in document.paragraphs if p.text.strip()] == [
        "{{" + data["placeholders"][key] + "}}" for key in ("title", "body")
    ]
    catalog = build_document_style_catalog({}, request_options={"locale": locale}, locales_dir=ROOT / "i18n/locales")
    original = document.styles.element.xml
    document, bindings = complete_managed_styles(document, catalog)
    assert document.styles.element.xml == original, "Shipped styles must already be complete and normalized"
    assert len({s.style_id for s in document.styles}) == len(document.styles)
    for definition in MANAGED_DOCUMENT_STYLES:
        style = bindings.get(definition.semantic_key)
        if not definition.is_builtin:
            assert style.name == catalog.name_for(definition.semantic_key)
    if locale == "ar_SA":
        for key in ("body_paragraph", "heading_1", "list_block", "quote_1"):
            p_pr = bindings.get(key).element.pPr
            assert p_pr.find(qn("w:bidi")).get(qn("w:val")) == "1"
            assert p_pr.find(qn("w:jc")).get(qn("w:val")) == "left"
        assert bindings.get("code_block").element.pPr.find(qn("w:bidi")).get(qn("w:val")) == "0"


def test_every_locale_has_a_dedicated_template():
    names = [
        tomllib.loads(p.read_text(encoding="utf-8"))["meta"]["template_name"]
        for p in (ROOT / "i18n/locales").glob("*.toml")
    ]
    assert len(names) == len(set(names)) == 12

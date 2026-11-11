"""Arabic defaults, custom-format preservation and native list direction."""

import pytest
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.styles.style import ParagraphStyle

from docwen_plugin_markdown.mistune_extensions import parse_markdown_text
from docwen_plugin_markdown.renderer import MdToDocxRenderer
from docwen_plugin_markdown.renderer_utils import apply_list_indent
from docwen_plugin_markdown.to_docx.direction import paragraph_is_rtl, style_is_rtl
from docwen_plugin_markdown.to_docx.managed_styles import complete_managed_styles
from docwen_plugin_markdown.to_docx.numbering import DocxListNumbering
from docwen_runtime.config.document_styles import build_document_style_catalog

from .conftest import PROJECT_ROOT

pytestmark = pytest.mark.contract


def test_arabic_completion_preserves_authored_styles_and_is_idempotent():
    catalog = build_document_style_catalog(
        {}, request_options={"locale": "ar_SA"}, locales_dir=PROJECT_ROOT / "i18n/locales"
    )
    doc, bindings = complete_managed_styles(Document(), catalog)
    body = bindings.get("body_paragraph")
    assert isinstance(body, ParagraphStyle)
    assert style_is_rtl(body)
    assert not style_is_rtl(bindings.get("code_block"))
    body.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    authored = body.element.xml
    again, second = complete_managed_styles(doc, catalog)
    assert second.get("body_paragraph").element.xml == authored
    list_style = second.get("list_block")
    assert isinstance(list_style, ParagraphStyle)
    paragraph = again.add_paragraph("نص عربي", style=list_style)
    assert paragraph_is_rtl(paragraph)
    apply_list_indent(paragraph, 1, left_twips=1080)
    indent = paragraph._p.get_or_add_pPr().find(qn("w:ind"))
    assert indent is not None
    assert indent.get(qn("w:right")) == "1080"
    assert indent.get(qn("w:left")) is None


@pytest.mark.parametrize("rtl", [False, True])
def test_native_list_indent_tracks_paragraph_direction(rtl):
    numbering = DocxListNumbering(Document(), right_to_left=rtl)
    num_id = numbering.create_list_definition({0: "ordered", 1: "unordered"})
    abstracts, _ = numbering.definition_elements()
    levels = abstracts[0].findall(qn("w:lvl"))
    for level, element in enumerate(levels):
        indent = element.find(f"{qn('w:pPr')}/{qn('w:ind')}")
        assert indent is not None
        edge = "right" if rtl else "left"
        assert indent.get(qn(f"w:{edge}")) == str(numbering.continuation_indent(num_id, level))
        alignment = element.find(qn("w:lvlJc"))
        assert alignment is not None
        assert alignment.get(qn("w:val")) == ("right" if rtl else "left")


def test_arabic_table_projects_direction_without_reordering_source_cells():
    catalog = build_document_style_catalog(
        {}, request_options={"locale": "ar_SA"}, locales_dir=PROJECT_ROOT / "i18n/locales"
    )
    doc, bindings = complete_managed_styles(Document(), catalog)
    MdToDocxRenderer(doc, managed_styles=bindings, table_style_key="three_line_table").render(
        parse_markdown_text("| البند | الحالة |\n| --- | --- |\n| مراجعة | مكتمل |\n")
    )
    table = doc.tables[0]
    direction = table._tbl.tblPr.find(qn("w:bidiVisual"))
    assert direction is not None
    assert direction.get(qn("w:val"), "1") == "1"
    assert [cell.text for cell in table.rows[0].cells] == ["البند", "الحالة"]

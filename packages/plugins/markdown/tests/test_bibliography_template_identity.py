"""Saved bibliography output binds real style IDs despite colliding names."""

from __future__ import annotations

from pathlib import Path

import pytest
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from lxml import etree

from docwen_core.docx_semantics import DocxSemanticImporter
from docwen_core.models.semantic_document import (
    SemanticBibliographyEntry,
    SemanticBibliographyFragment,
    SemanticBibliographyRun,
)
from docwen_plugin_markdown.to_docx.bibliography import BibliographyConversionError, prepare_bibliography_anchor
from docwen_plugin_markdown.to_docx.converter import MdToDocxConverter

from .conftest import make_context
from .test_bibliography_resource import _attach_bibliography, _bibliography_payload

pytestmark = pytest.mark.contract


@pytest.mark.parametrize("style_id", ["42", "原始_参考文献"])
def test_synthesized_bibliography_keeps_real_style_id_when_another_style_name_matches(
    tmp_path: Path, style_id: str
) -> None:
    document = Document()
    bibliography = document.styles.add_style("Bibliography", WD_STYLE_TYPE.PARAGRAPH)
    bibliography.element.set(qn("w:styleId"), style_id)
    decoy = document.styles.add_style(style_id, WD_STYLE_TYPE.PARAGRAPH)
    decoy.element.set(qn("w:styleId"), "OtherBibliography")
    hyperlink = document.styles.add_style("Hyperlink", WD_STYLE_TYPE.CHARACTER)
    hyperlink.element.set(qn("w:styleId"), "Link42")
    link_decoy = document.styles.add_style("Link42", WD_STYLE_TYPE.CHARACTER)
    link_decoy.element.set(qn("w:styleId"), "OtherLink")
    before = {
        style.style_id: etree.tostring(style.element, method="c14n")
        for style in (bibliography, decoy, hyperlink, link_decoy)
    }
    document.add_paragraph("{{body}}")
    template = tmp_path / "template.docx"
    document.save(str(template))
    source = tmp_path / "source.md"
    source.write_text("Body.\n", encoding="utf-8")
    resource = tmp_path / "bibliography.json"
    resource.write_bytes(_bibliography_payload())
    context, workspace = make_context(str(source), options={"template_name": str(template)})
    _attach_bibliography(context, workspace, source, resource)

    result = MdToDocxConverter().convert(context)

    assert result.success, result.error
    output = Document(result.artifacts[0].staging_path)
    paragraph = next(item for item in output.paragraphs if "Neutral documents" in item.text)
    assert paragraph._p.pPr is not None
    reference = paragraph._p.pPr.find(qn("w:pStyle"))
    assert reference is not None and reference.get(qn("w:val")) == style_id
    assert [node.get(qn("w:val")) for node in paragraph._p.iter(qn("w:rStyle"))] == ["Link42"]
    for element in output.styles.element:
        identity = element.get(qn("w:styleId"))
        if identity in before:
            assert etree.tostring(element, method="c14n") == before[identity]
    imported = DocxSemanticImporter().import_document(output)
    assert imported.document.bibliography is not None
    assert imported.document.bibliography.entries[0].item_id == "smith2025"


@pytest.mark.parametrize("damage", ["missing", "wrong_type", "duplicate"])
def test_synthesized_bibliography_rejects_invalid_style_identity_before_inserting_anchor(damage: str) -> None:
    document = Document()
    document.add_paragraph("{{body}}")
    style = document.styles.add_style("Template bibliography", WD_STYLE_TYPE.PARAGRAPH)
    style.element.set(qn("w:styleId"), "42")
    if damage == "missing":
        document.styles.element.remove(style.element)
    elif damage == "wrong_type":
        style.element.set(qn("w:type"), "character")
    else:
        other = document.styles.add_style("Other bibliography", WD_STYLE_TYPE.PARAGRAPH)
        other.element.set(qn("w:styleId"), "42")
    before = document.element.xml
    fragment = SemanticBibliographyFragment(
        entries=(SemanticBibliographyEntry("entry", (SemanticBibliographyRun("Reference."),)),)
    )

    with pytest.raises(BibliographyConversionError, match="resolve by ID"):
        prepare_bibliography_anchor(document, fragment, bibliography_style_id="42")

    assert document.element.xml == before

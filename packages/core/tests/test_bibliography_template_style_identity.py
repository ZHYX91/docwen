"""Bibliography links retain authored character-style identities after saving."""

from __future__ import annotations

from copy import deepcopy
from pathlib import Path

import pytest
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml import etree

from docwen_core.docx_semantics import DocxSemanticImporter, DocxSemanticRenderer
from docwen_core.models.semantic_document import (
    SemanticBibliographyEntry,
    SemanticBibliographyFragment,
    SemanticBibliographyRun,
    SemanticDocumentValidationError,
)

pytestmark = pytest.mark.contract


def _fragment() -> SemanticBibliographyFragment:
    return SemanticBibliographyFragment(
        entries=(
            SemanticBibliographyEntry(
                "reference-one",
                (
                    SemanticBibliographyRun("Reference "),
                    SemanticBibliographyRun("linked title", italic=True, href="https://example.org/reference"),
                ),
            ),
        )
    )


def _template(style_id: str):
    document = Document()
    style = document.styles.add_style("Authored link style", WD_STYLE_TYPE.CHARACTER)
    style.element.set(qn("w:styleId"), style_id)
    aliases = OxmlElement("w:aliases")
    aliases.set(qn("w:val"), "超链接,Hyperlink")
    style.element.append(aliases)
    anchor = document.add_paragraph("Bibliography placeholder")
    return document, style, anchor


@pytest.mark.parametrize("style_id", ["42", "原始_链接", 'Link<&"Identity'])
def test_bibliography_saved_round_trip_preserves_template_character_identity(tmp_path: Path, style_id: str) -> None:
    document, style, anchor = _template(style_id)
    style_before = etree.tostring(style.element, method="c14n")
    fragment = _fragment()

    DocxSemanticRenderer(document).render_bibliography_fragment(
        fragment, placeholder_anchor=anchor, hyperlink_style_id=style_id
    )
    output = tmp_path / "references.docx"
    document.save(str(output))
    reopened = Document(str(output))

    assert [node.get(qn("w:val")) for node in reopened.element.iter(qn("w:rStyle"))] == [style_id]
    matching = [node for node in reopened.styles.element if node.get(qn("w:styleId")) == style_id]
    assert len(matching) == 1
    assert etree.tostring(matching[0], method="c14n") == style_before
    imported = DocxSemanticImporter().import_document(reopened)
    assert imported.document.bibliography == fragment
    assert not imported.diagnostics


@pytest.mark.parametrize("damage", ["missing", "wrong_type", "duplicate", "lookalike_name", "unbound", "empty_id"])
def test_bibliography_writer_rejects_unresolved_bound_identity_without_mutation(damage: str) -> None:
    document, style, anchor = _template("TemplateLink")
    bound_id = "TemplateLink"
    if damage == "missing":
        document.styles.element.remove(style.element)
    elif damage == "wrong_type":
        style.element.set(qn("w:type"), "paragraph")
    elif damage == "duplicate":
        document.styles.element.append(deepcopy(style.element))
    elif damage == "lookalike_name":
        style.name = "TemplateLink"
        style.element.set(qn("w:styleId"), "AnotherIdentity")
    elif damage == "empty_id":
        style.element.set(qn("w:styleId"), "")
        bound_id = ""
    before = document.element.xml
    relationships_before = dict(document.part.rels)

    with pytest.raises(SemanticDocumentValidationError):
        DocxSemanticRenderer(document).render_bibliography_fragment(
            _fragment(), placeholder_anchor=anchor, hyperlink_style_id=None if damage == "unbound" else bound_id
        )

    assert document.element.xml == before
    assert dict(document.part.rels) == relationships_before


@pytest.mark.parametrize("damage", ["missing", "wrong_type", "duplicate", "lookalike_name", "no_rpr", "no_rstyle"])
def test_bibliography_reader_requires_exact_unique_character_reference(tmp_path: Path, damage: str) -> None:
    document, style, anchor = _template("TemplateLink")
    DocxSemanticRenderer(document).render_bibliography_fragment(
        _fragment(), placeholder_anchor=anchor, hyperlink_style_id="TemplateLink"
    )
    if damage == "missing":
        document.styles.element.remove(style.element)
    elif damage == "wrong_type":
        style.element.set(qn("w:type"), "paragraph")
    elif damage == "duplicate":
        document.styles.element.append(deepcopy(style.element))
    elif damage == "lookalike_name":
        style.name = "TemplateLink"
        style.element.set(qn("w:styleId"), "AnotherIdentity")
    else:
        hyperlink = next(document.element.iter(qn("w:hyperlink")))
        element = next(hyperlink.iter(qn("w:rPr" if damage == "no_rpr" else "w:rStyle")))
        parent = element.getparent()
        assert parent is not None
        parent.remove(element)
    output = tmp_path / "damaged.docx"
    document.save(str(output))

    imported = DocxSemanticImporter().import_document(Document(str(output)))

    assert imported.document.bibliography is None
    assert "semantic.docx.bibliography.entry_marker_invalid" in {item.code for item in imported.diagnostics}

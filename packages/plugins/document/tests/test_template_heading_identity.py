"""Template style names and decoration cannot override authored Heading levels."""

from __future__ import annotations

import hashlib
from pathlib import Path
from typing import cast

import pytest
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.styles.style import ParagraphStyle

from docwen_core.docx_semantics_v3 import DocxSemanticsV3Session

from ._to_markdown_standard_parity_support import _convert_document_fixture_to_markdown

pytestmark = pytest.mark.contract


@pytest.mark.parametrize("addressable", [False, True])
@pytest.mark.parametrize("template_outline", [None, "8"])
def test_saved_heading_keeps_source_level_with_custom_template_style(
    tmp_path: Path, addressable: bool, template_outline: str | None
) -> None:
    document = Document()
    style = cast(ParagraphStyle, document.styles["Heading 3"])
    style.element.set(qn("w:styleId"), "17_模板标题")
    name = style.element.find(qn("w:name"))
    assert name is not None
    name.set(qn("w:val"), "源模板样式")
    properties = style.element.find(qn("w:pPr"))
    assert properties is not None
    for outline in properties.findall(qn("w:outlineLvl")):
        properties.remove(outline)
    if template_outline is not None:
        outline = OxmlElement("w:outlineLvl")
        outline.set(qn("w:val"), template_outline)
        properties.append(outline)
    paragraph = document.add_paragraph("Generated heading", style=style)
    shading = OxmlElement("w:shd")
    shading.set(qn("w:fill"), "D9D9D9")
    paragraph._p.get_or_add_pPr().append(shading)
    if not addressable:
        # Generated headings without a source ID carry their explicit outline
        # level directly; template decoration remains independent of semantics.
        direct = OxmlElement("w:outlineLvl")
        direct.set(qn("w:val"), "2")
        paragraph._p.get_or_add_pPr().append(direct)
    session = DocxSemanticsV3Session(document, source_sha256=hashlib.sha256(b"### Generated heading").hexdigest())
    if addressable:
        session.bind_heading(paragraph, {"kind": "heading", "id": "generated"}, heading_level=3)
    session.finalize_document()
    path = tmp_path / "template-heading.docx"
    document.save(str(path))
    session.write_package(path)
    session.prove_package(path)
    markdown = _convert_document_fixture_to_markdown(
        tmp_path,
        Document(str(path)),
        request_id="reopened-heading",
        options={"markdown_extensions": {"output": {"captions_references": True, "extended_headings": True}}},
    )
    assert "### Generated heading" in markdown
    assert "######### Generated heading" not in markdown
    assert "```" not in markdown
    if addressable:
        assert "^generated" in markdown

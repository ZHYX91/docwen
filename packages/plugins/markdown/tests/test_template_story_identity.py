"""All template consumers retain relationship-owned header/footer stories."""

from __future__ import annotations

from pathlib import Path
from zipfile import ZipFile

import pytest
from docx import Document
from docx.opc.packuri import PackURI
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml import etree

from docwen_core.models.clipboard_document import (
    ClipboardDocument,
    ClipboardParagraph,
    ClipboardText,
    clipboard_document_to_bytes,
)
from docwen_plugin_markdown.structured_clipboard import convert_clipboard_document_to_docx
from docwen_plugin_markdown.to_docx.converter import MdToDocxConverter

from ._resolved_v4_support import _NEUTRAL, _PLAN, _context, _refs
from .conftest import make_context

pytestmark = pytest.mark.contract

_STORIES = ["header", "first_page_header", "even_page_header", "footer", "first_page_footer", "even_page_footer"]


def _owned_marker(paragraph, kind: str) -> None:
    if kind == "field":
        owner = OxmlElement("w:fldSimple")
        owner.set(qn("w:instr"), " PAGE ")
        content = owner
    else:
        owner = OxmlElement("w:sdt")
        properties = OxmlElement("w:sdtPr")
        binding = OxmlElement("w:dataBinding")
        binding.set(qn("w:xpath"), "/record/name")
        binding.set(qn("w:storeItemID"), "{E5149C46-A9B1-4CB1-AC2D-2D8BFF7D4581}")
        properties.append(binding)
        content = OxmlElement("w:sdtContent")
        owner.extend((properties, content))
    run = OxmlElement("w:r")
    text = OxmlElement("w:t")
    text.text = "{{name}}"
    run.append(text)
    content.append(run)
    paragraph._p.append(owner)


@pytest.mark.parametrize("story_name", _STORIES)
def test_relocated_story_fills_plain_values_and_owns_its_hyperlinks(tmp_path: Path, story_name: str) -> None:
    document = Document()
    document.add_paragraph("{{body}}")
    story = getattr(document.sections[0], story_name)
    story.paragraphs[0].text = "{{name}} / {{link}}"
    story.part._partname = PackURI("/layout/regions/custom.xml")
    template = tmp_path / "template.docx"
    document.save(str(template))
    before = template.read_bytes()
    source = tmp_path / "input.md"
    source.write_text('---\nname: Alice\nlink: "[Website](https://example.org/)"\n---\nBody.\n', encoding="utf-8")
    context, _ = make_context(str(source), options={"template_name": str(template), "heading_merge_mode": "never"})
    result = MdToDocxConverter().convert(context)
    assert result.success, result.error
    reopened = Document(result.artifacts[0].staging_path)
    output_story = getattr(reopened.sections[0], story_name)
    assert str(output_story.part.partname) == "/layout/regions/custom.xml"
    assert "".join(node.text or "" for node in output_story._element.iter(qn("w:t"))) == "Alice / Website"
    links = list(output_story._element.iter(qn("w:hyperlink")))
    assert len(links) == 1
    assert output_story.part.rels[links[0].get(qn("r:id"))].target_ref == "https://example.org/"
    assert template.read_bytes() == before
    with ZipFile(result.artifacts[0].staging_path) as package:
        assert "layout/regions/custom.xml" in package.namelist()


@pytest.mark.parametrize("story_name", ["header", "footer"])
@pytest.mark.parametrize("kind", ["field", "bound"])
@pytest.mark.parametrize("route", ["ordinary", "resolved", "clipboard"])
def test_relocated_story_owned_markers_fail_before_modification(
    tmp_path: Path, story_name: str, kind: str, route: str
) -> None:
    document = Document()
    document.add_paragraph("{{body}}")
    story = getattr(document.sections[0], story_name)
    story.part._partname = PackURI("/layout/region.xml")
    _owned_marker(story.paragraphs[0], kind)
    template = tmp_path / "template.docx"
    document.save(str(template))
    before = template.read_bytes()
    options = {"template_name": str(template), "heading_merge_mode": "never"}
    if route == "resolved":
        context, workspace = _context(tmp_path, _refs(_NEUTRAL, _PLAN), options=options)
    else:
        source = tmp_path / ("input.dwclip" if route == "clipboard" else "input.md")
        if route == "clipboard":
            source.write_bytes(
                clipboard_document_to_bytes(ClipboardDocument((ClipboardParagraph((ClipboardText("Body"),)),)))
            )
        else:
            source.write_text("---\nname: Alice\n---\nBody\n", encoding="utf-8")
        context, workspace = make_context(str(source), options=options)
    result = (
        convert_clipboard_document_to_docx(context) if route == "clipboard" else MdToDocxConverter().convert(context)
    )
    assert not result.success
    assert result.error is not None and result.error.diagnostic_code == "MD2DOCX-TEMPLATE-PLACEHOLDER-OWNERSHIP"
    assert not workspace.registered_artifacts
    assert template.read_bytes() == before


@pytest.mark.parametrize(
    "marker,code",
    [("{{body}}", "MD2DOCX-TEMPLATE-BODY-PLACEMENT"), ("{{附件说明}}", "MD2DOCX-GONGWEN-FIELD-PLACEMENT")],
)
def test_relocated_story_cannot_hide_reserved_layout_markers(tmp_path: Path, marker: str, code: str) -> None:
    document = Document()
    document.add_paragraph("{{body}}")
    header = document.sections[0].header
    header.part._partname = PackURI("/layout/custom.xml")
    header.paragraphs[0].text = marker
    template = tmp_path / "template.docx"
    document.save(str(template))
    source = tmp_path / "input.md"
    source.write_text("Body\n", encoding="utf-8")
    context, workspace = make_context(str(source), options={"template_name": str(template)})
    result = MdToDocxConverter().convert(context)
    assert not result.success
    assert result.error is not None and result.error.diagnostic_code == code
    assert not workspace.registered_artifacts


@pytest.mark.parametrize("member", ["_rels/.rels", "word/_rels/document.xml.rels"])
def test_relationship_annotations_do_not_reject_a_valid_template(tmp_path: Path, member: str) -> None:
    document = Document()
    document.add_paragraph("{{body}}")
    header = document.sections[0].header
    header.part._partname = PackURI("/layout/annotated.xml")
    header.paragraphs[0].text = "{{name}}"
    template = tmp_path / "template.docx"
    document.save(str(template))
    with ZipFile(template) as archive:
        members = {name: archive.read(name) for name in archive.namelist()}
    relationships = etree.fromstring(members[member])
    relationships.insert(0, etree.Comment("Legal producer annotation"))
    relationships.append(etree.ProcessingInstruction("producer", "mode='test'"))
    members[member] = etree.tostring(relationships)
    with ZipFile(template, "w") as archive:
        for name, payload in members.items():
            archive.writestr(name, payload)
    original = template.read_bytes()
    source = tmp_path / "input.md"
    source.write_text("---\nname: Alice\n---\nBody\n", encoding="utf-8")
    context, _ = make_context(str(source), options={"template_name": str(template)})
    result = MdToDocxConverter().convert(context)
    assert result.success, result.error
    output = Document(result.artifacts[0].staging_path)
    assert output.sections[0].header.paragraphs[0].text == "Alice"
    assert str(output.sections[0].header.part.partname) == "/layout/annotated.xml"
    with ZipFile(result.artifacts[0].staging_path) as archive:
        final = {
            item.get("Id"): dict(item.attrib)
            for item in etree.fromstring(archive.read(member))
            if isinstance(item.tag, str)
        }
    assert all(final[item.get("Id")] == dict(item.attrib) for item in relationships if isinstance(item.tag, str))
    assert template.read_bytes() == original

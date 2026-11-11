"""Final DOCX artifacts retain template boundaries across all filling routes."""

from __future__ import annotations

from copy import deepcopy
from pathlib import Path
from typing import TYPE_CHECKING
from zipfile import ZipFile

import pytest
from docx import Document
from docx.enum.section import WD_SECTION_START
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm
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

if TYPE_CHECKING:
    from docx.document import Document as DocumentObject
    from docx.text.paragraph import Paragraph

pytestmark = pytest.mark.contract


def _xml(element: etree._Element) -> bytes:
    return etree.tostring(element, method="c14n", exclusive=True, with_tail=False)


def _section_template() -> DocumentObject:
    document = Document()
    document.add_paragraph("Fixed before")
    anchor = document.add_paragraph("{{body}}")
    document.add_section(WD_SECTION_START.ODD_PAGE)
    boundary_paragraph = document.paragraphs[-1]
    boundary = boundary_paragraph._p.get_or_add_pPr().find(qn("w:sectPr"))
    assert boundary is not None
    anchor._p.get_or_add_pPr().append(boundary)
    document.element.body.remove(boundary_paragraph._p)
    document.add_paragraph("Fixed after")
    for index, section in enumerate(document.sections):
        section.left_margin = Mm(16 if index == 0 else 26)
        section.right_margin = Mm(18 if index == 0 else 28)
        section.header_distance = Mm(9 if index == 0 else 11)
        section.footer_distance = Mm(10 if index == 0 else 12)
        section.different_first_page_header_footer = index == 0
        for name in ("header", "footer", "first_page_header", "first_page_footer"):
            story = getattr(section, name)
            story.is_linked_to_previous = False
            story.paragraphs[0].text = f"Section {index + 1} {name}"
        page_number = OxmlElement("w:pgNumType")
        page_number.set(qn("w:fmt"), "decimal" if index == 0 else "upperRoman")
        page_number.set(qn("w:start"), "12" if index == 0 else "50")
        section._sectPr.insert_element_before(page_number, "w:cols")
        borders = OxmlElement("w:pgBorders")
        borders.set(qn("w:offsetFrom"), "page")
        bottom = OxmlElement("w:bottom")
        bottom.set(qn("w:val"), "single")
        bottom.set(qn("w:sz"), "8")
        bottom.set(qn("w:space"), "12")
        borders.append(bottom)
        section._sectPr.insert_element_before(borders, "w:pgNumType")
    return document


def _request(tmp_path: Path, route: str, template: Path, *, source: str = "Generated body.\n"):
    if route == "resolved":
        return _context(
            tmp_path,
            _refs(_NEUTRAL, _PLAN),
            options={"locale": "zh_CN", "heading_merge_mode": "never", "template_name": str(template)},
        )
    input_path = tmp_path / ("source.dwclip" if route == "clipboard" else "source.md")
    if route == "clipboard":
        input_path.write_bytes(
            clipboard_document_to_bytes(ClipboardDocument((ClipboardParagraph((ClipboardText(source.strip()),)),)))
        )
    else:
        input_path.write_text(source, encoding="utf-8")
    return make_context(str(input_path), options={"template_name": str(template), "heading_merge_mode": "never"})


def _convert(context, route: str):
    if route == "clipboard":
        return convert_clipboard_document_to_docx(context)
    return MdToDocxConverter().convert(context)


@pytest.mark.parametrize("route", ["ordinary", "resolved", "clipboard"])
def test_filled_artifact_retains_anchor_section_and_all_header_footer_stories(tmp_path: Path, route: str) -> None:
    document = _section_template()
    template = tmp_path / "sections.docx"
    document.save(str(template))
    original = template.read_bytes()
    expected_sections = [_xml(section._sectPr) for section in Document(str(template)).sections]
    context, _workspace = _request(tmp_path, route, template)

    result = _convert(context, route)

    assert result.success, result.error
    output = next(item.staging_path for item in result.artifacts if item.is_primary)
    reopened = Document(output)
    assert [_xml(section._sectPr) for section in reopened.sections] == expected_sections
    texts = [paragraph.text for paragraph in reopened.paragraphs]
    assert texts[0] == "Fixed before" and texts[-1] == "Fixed after"
    assert "{{body}}" not in texts
    visible = "".join(node.text or "" for node in reopened.element.iter(qn("w:t")))
    assert ("Architecture" if route == "resolved" else "Generated body.") in visible
    with ZipFile(template) as before, ZipFile(output) as after:
        stories = [name for name in before.namelist() if name.startswith(("word/header", "word/footer"))]
        assert len(stories) == 8
        for name in stories:
            assert after.read(name) == before.read(name), name
    assert template.read_bytes() == original


def test_generated_section_break_inherits_anchor_section_instead_of_last_template_section(tmp_path: Path) -> None:
    document = _section_template()
    template = tmp_path / "section-break.docx"
    document.save(str(template))
    source_sections = [deepcopy(section._sectPr) for section in Document(str(template)).sections]
    context, _workspace = _request(tmp_path, "ordinary", template, source="First block.\n\n***\n\nSecond block.\n")

    result = _convert(context, "ordinary")

    assert result.success, result.error
    reopened = Document(result.artifacts[0].staging_path)
    assert len(reopened.sections) == 3
    assert _xml(reopened.sections[0]._sectPr) == _xml(source_sections[0])
    assert _xml(reopened.sections[2]._sectPr) == _xml(source_sections[1])
    continuation = source_sections[0]
    continuation.get_or_add_type().set(qn("w:val"), "nextPage")
    page_number = continuation.find(qn("w:pgNumType"))
    assert page_number is not None
    page_number.attrib.pop(qn("w:start"))
    assert _xml(reopened.sections[1]._sectPr) == _xml(continuation)
    assert [section.left_margin.twips for section in reopened.sections] == [Mm(16).twips, Mm(16).twips, Mm(26).twips]


def _complex_field(paragraph: Paragraph, result: str) -> None:
    for kind in ("begin", "separate"):
        field = OxmlElement("w:fldChar")
        field.set(qn("w:fldCharType"), kind)
        paragraph.add_run()._r.append(field)
        if kind == "begin":
            instruction = OxmlElement("w:instrText")
            instruction.set(qn("xml:space"), "preserve")
            instruction.text = " DOCPROPERTY Subject \\* MERGEFORMAT "
            paragraph.add_run()._r.append(instruction)
    paragraph.add_run(result)
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    paragraph.add_run()._r.append(end)


def _owned_placeholder(document: DocumentObject, kind: str) -> None:
    if kind == "complex":
        _complex_field(document.add_paragraph(), "{{name}}")
    elif kind == "simple":
        field = OxmlElement("w:fldSimple")
        field.set(qn("w:instr"), " DOCPROPERTY Subject ")
        run = OxmlElement("w:r")
        text = OxmlElement("w:t")
        text.text = "{{name}}"
        run.append(text)
        field.append(run)
        document.sections[0].header.paragraphs[0]._p.append(field)
    else:
        control = OxmlElement("w:sdt")
        properties = OxmlElement("w:sdtPr")
        binding = OxmlElement("w:dataBinding")
        binding.set(qn("w:xpath"), "/record/name")
        binding.set(qn("w:storeItemID"), "{E5149C46-A9B1-4CB1-AC2D-2D8BFF7D4581}")
        properties.append(binding)
        control.append(properties)
        content = OxmlElement("w:sdtContent")
        paragraph = document.add_paragraph("{{name}}")
        content.append(paragraph._p)
        control.append(content)
        document.element.body.insert_element_before(control, "w:sectPr")


@pytest.mark.parametrize("route", ["ordinary", "resolved", "clipboard"])
@pytest.mark.parametrize("kind", ["complex", "simple", "bound_sdt"])
def test_owned_placeholder_fails_before_publishing_any_artifact(tmp_path: Path, route: str, kind: str) -> None:
    document = Document()
    document.add_paragraph("{{body}}")
    _owned_placeholder(document, kind)
    template = tmp_path / "owned-placeholder.docx"
    document.save(str(template))
    original = template.read_bytes()
    context, workspace = _request(tmp_path, route, template, source="---\nname: Filled\n---\n\nBody.\n")

    result = _convert(context, route)

    assert not result.success
    assert result.error is not None
    assert result.error.error_type == "invalid_input"
    assert result.error.diagnostic_code == "MD2DOCX-TEMPLATE-PLACEHOLDER-OWNERSHIP"
    assert result.artifacts == [] and workspace.registered_artifacts == []
    assert list(Path(workspace.staging_dir).iterdir()) == []
    assert template.read_bytes() == original


def test_metadata_beside_field_fills_without_changing_field_instruction_or_cached_result(tmp_path: Path) -> None:
    document = Document()
    paragraph = document.add_paragraph("{{name}} | ")
    _complex_field(paragraph, "Original cached subject")
    document.add_paragraph("{{body}}")
    template = tmp_path / "field-neighbor.docx"
    document.save(str(template))
    expected_runs = [_xml(run._r) for run in paragraph.runs[1:]]
    context, _workspace = _request(tmp_path, "ordinary", template, source="---\nname: Filled\n---\n\nGenerated body.\n")

    result = _convert(context, "ordinary")

    assert result.success, result.error
    reopened = Document(result.artifacts[0].staging_path)
    assert reopened.paragraphs[0].text == "Filled | Original cached subject"
    assert [_xml(run._r) for run in reopened.paragraphs[0].runs[1:]] == expected_runs

"""Final DOCX contracts for section ownership and structural body anchors."""

from __future__ import annotations

from copy import deepcopy
from io import BytesIO
from pathlib import Path
from typing import Any
from zipfile import ZipFile

import pytest
from docx import Document
from docx.enum.section import WD_ORIENT, WD_SECTION_START
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Mm, Pt
from lxml import etree

from docwen_core.docx_comments import append_docx_comment
from docwen_core.models.clipboard_document import (
    ClipboardDocument,
    ClipboardParagraph,
    ClipboardTable,
    ClipboardTableCell,
    ClipboardText,
    clipboard_document_to_bytes,
)
from docwen_plugin_markdown.structured_clipboard import _place_docx_body, convert_clipboard_document_to_docx
from docwen_plugin_markdown.template_filler import fill_template
from docwen_plugin_markdown.template_utils import body_content_extent
from docwen_plugin_markdown.to_docx.breaks import append_section_break_to_paragraph, insert_section_break
from docwen_plugin_markdown.to_docx.converter import MdToDocxConverter
from docwen_plugin_markdown.to_docx.sections import SectionBreakContext, section_properties_at, splice_body_elements

from .conftest import make_context

pytestmark = pytest.mark.contract


def _xml(element: Any) -> bytes:
    return etree.tostring(element, method="c14n", exclusive=True)


def _saved(document: Any) -> bytes:
    stream = BytesIO()
    document.save(stream)
    return stream.getvalue()


def _reopen(document: Any) -> Any:
    return Document(BytesIO(_saved(document)))


def _word(tag: str, **attributes: str) -> Any:
    element = OxmlElement(f"w:{tag}")
    for key, value in attributes.items():
        element.set(qn(f"w:{key}"), value)
    return element


def _template(*, anchor_owns_boundary: bool = False) -> tuple[Any, Any]:
    document = Document()
    document.add_paragraph("Template preface")
    anchor = document.add_paragraph("{{body}}")
    first = document.sections[0]
    first.orientation = WD_ORIENT.LANDSCAPE
    first.page_width = Inches(11)
    first.page_height = Inches(8.5)
    first.left_margin = Inches(0.75)
    first.right_margin = Inches(0.9)
    first.top_margin = Inches(0.65)
    first.bottom_margin = Inches(0.8)
    first.header_distance = Inches(0.25)
    first.footer_distance = Inches(0.3)
    first.different_first_page_header_footer = True
    document.settings.odd_and_even_pages_header_footer = True
    for name in ("header", "first_page_header", "even_page_header", "footer", "first_page_footer", "even_page_footer"):
        getattr(first, name).paragraphs[0].text = f"First section {name}"
    first._sectPr.get_or_add_type().set(qn("w:val"), "evenPage")
    borders = _word("pgBorders", offsetFrom="page")
    borders.append(_word("top", val="double", sz="8", space="18", color="234567"))
    first._sectPr.insert_element_before(borders, "w:cols", "w:docGrid")
    numbering = _word("pgNumType", start="7", fmt="upperRoman", chapStyle="2", chapSep="hyphen")
    first._sectPr.insert_element_before(numbering, "w:cols", "w:docGrid")
    columns = first._sectPr.find(qn("w:cols"))
    assert columns is not None
    columns.set(qn("w:num"), "2")
    columns.set(qn("w:sep"), "1")
    first._sectPr.insert_element_before(
        _word("textDirection", val="tbRl"), "w:bidi", "w:rtlGutter", "w:docGrid", "w:printerSettings", "w:sectPrChange"
    )
    # The default package declares w14 ignorable, so an unrecognized future
    # property is retained without making the fixture require XML repair.
    extension = etree.SubElement(first._sectPr, qn("w14:sectionProbe"))
    extension.set("keep", "authored-data")
    document.add_section(WD_SECTION_START.ODD_PAGE)
    boundary_paragraph = document.paragraphs[-1]
    if anchor_owns_boundary:
        boundary = boundary_paragraph._p.find(f"{qn('w:pPr')}/{qn('w:sectPr')}")
        assert boundary is not None
        anchor._p.get_or_add_pPr().append(boundary)
        document._element.body.remove(boundary_paragraph._p)
    document.add_paragraph("Unrelated last section")
    last = document.sections[-1]
    last.orientation = WD_ORIENT.PORTRAIT
    last.page_width = Inches(6)
    last.page_height = Inches(9)
    last.left_margin = Inches(1.25)
    last.different_first_page_header_footer = False
    last.header.is_linked_to_previous = False
    last.header.paragraphs[0].text = "Last section header"
    last_numbering = last._sectPr.find(qn("w:pgNumType"))
    assert last_numbering is not None
    last_numbering.set(qn("w:start"), "31")
    return document, anchor


def _continuation_xml(original: Any, start_type: str) -> bytes:
    expected = deepcopy(original)
    expected.get_or_add_type().set(qn("w:val"), start_type)
    numbering = expected.find(qn("w:pgNumType"))
    assert numbering is not None
    numbering.attrib.pop(qn("w:start"))
    return _xml(expected)


def _header_footer_parts(payload: bytes) -> dict[str, bytes]:
    with ZipFile(BytesIO(payload)) as archive:
        return {
            name: archive.read(name) for name in archive.namelist() if name.startswith(("word/header", "word/footer"))
        }


@pytest.mark.parametrize("anchor_owns_boundary", [False, True])
def test_context_splits_actual_template_section_and_preserves_all_settings(anchor_owns_boundary: bool) -> None:
    document, anchor = _template(anchor_owns_boundary=anchor_owns_boundary)
    original = deepcopy(section_properties_at(document, anchor))
    unrelated = _xml(document.sections[-1]._sectPr)
    original_parts = _header_footer_parts(_saved(document))
    context = SectionBreakContext(document, anchor)
    one = document.add_paragraph("First generated block")
    append_section_break_to_paragraph(one, document, "continuous", context=context)
    two = document.add_paragraph("Second generated block")
    append_section_break_to_paragraph(two, document, "odd", context=context)
    three = document.add_paragraph("Third generated block")
    fill_template(document, {}, [one, two, three], anchor)

    payload = _saved(document)
    reopened = Document(BytesIO(payload))
    sections = list(reopened.sections)
    assert len(sections) == 4
    assert _xml(sections[0]._sectPr) == _xml(original)
    # ISO 29500 17.6.22: w:type starts the current section relative to its
    # predecessor; it is not the break following the paragraph holding it.
    # https://learn.microsoft.com/dotnet/api/documentformat.openxml.wordprocessing.sectiontype
    assert _xml(sections[1]._sectPr) == _continuation_xml(original, "continuous")
    assert _xml(sections[2]._sectPr) == _continuation_xml(original, "oddPage")
    assert _xml(sections[3]._sectPr) == unrelated
    assert _header_footer_parts(payload) == original_parts
    assert reopened.settings.odd_and_even_pages_header_footer is True
    for section in sections[:3]:
        assert section.different_first_page_header_footer is True
        assert section.orientation == WD_ORIENT.LANDSCAPE
        for name in (
            "header",
            "first_page_header",
            "even_page_header",
            "footer",
            "first_page_footer",
            "even_page_footer",
        ):
            assert getattr(section, name).paragraphs[0].text == f"First section {name}"


def test_direct_helper_uses_next_boundary_instead_of_last_document_section() -> None:
    document, anchor = _template()
    original = deepcopy(section_properties_at(document, anchor))
    unrelated = _xml(document.sections[-1]._sectPr)
    append_section_break_to_paragraph(anchor, document, "even")
    reopened = _reopen(document)
    assert len(reopened.sections) == 3
    assert _xml(reopened.sections[0]._sectPr) == _xml(original)
    assert _xml(reopened.sections[1]._sectPr) == _continuation_xml(original, "evenPage")
    assert _xml(reopened.sections[2]._sectPr) == unrelated


def test_inserted_sections_keep_absent_header_references_as_inheritance() -> None:
    document = Document()
    document.sections[0].header.paragraphs[0].text = "Inherited header"
    document.add_paragraph("Original first section")
    document.add_section()
    anchor = document.add_paragraph("{{body}}")
    document.add_section()
    document.add_paragraph("Original last section")
    context = SectionBreakContext(document, anchor)
    generated = document.add_paragraph("First generated block")
    append_section_break_to_paragraph(generated, document, context=context)
    remaining = document.add_paragraph("Second generated block")
    fill_template(document, {}, [generated, remaining], anchor)
    reopened = _reopen(document)
    assert len(reopened.sections) == 4
    for section in list(reopened.sections)[1:]:
        assert section._sectPr.find(qn("w:headerReference")) is None
        assert section.header.is_linked_to_previous is True
        assert section.header.paragraphs[0].text == "Inherited header"


@pytest.mark.parametrize("placement", ["ordinary", "structured"])
@pytest.mark.parametrize("ending", ["paragraph", "table", "empty", "section_break"])
def test_anchor_section_survives_each_replacement_shape(placement: str, ending: str) -> None:
    document, anchor = _template(anchor_owns_boundary=True)
    original = _xml(section_properties_at(document, anchor))
    unrelated = _xml(document.sections[-1]._sectPr)
    rendered: list[Any] = []
    if ending == "paragraph":
        rendered.append(document.add_paragraph("Generated body"))
    elif ending == "table":
        table = document.add_table(rows=1, cols=1)
        table.cell(0, 0).text = "Generated table"
        rendered.append(table)
    elif ending == "section_break":
        rendered.append(insert_section_break(document, context=SectionBreakContext(document, anchor)))
    if placement == "ordinary":
        fill_template(document, {}, rendered, anchor)
    else:
        _place_docx_body(document, [item._element for item in rendered], anchor)
    reopened = _reopen(document)
    assert len(reopened.sections) == (3 if ending == "section_break" else 2)
    assert _xml(reopened.sections[0]._sectPr) == original
    assert _xml(reopened.sections[-1]._sectPr) == unrelated
    assert "{{body}}" not in reopened.element.xml
    for paragraph in reopened.paragraphs:
        assert len(paragraph._p.findall(f"{qn('w:pPr')}/{qn('w:sectPr')}")) <= 1
    if ending in {"table", "empty", "section_break"}:
        boundary_paragraphs = [
            paragraph
            for paragraph in reopened.paragraphs
            if paragraph._p.find(f"{qn('w:pPr')}/{qn('w:sectPr')}") is not None
        ]
        assert boundary_paragraphs[-1].text == ""
    if ending == "table":
        assert reopened.tables[0].cell(0, 0).text == "Generated table"


@pytest.mark.parametrize("ending", ["paragraph", "table", "empty"])
def test_bookmarks_and_comment_ranges_keep_identity_order_and_comment_part(ending: str) -> None:
    document, anchor = _template(anchor_owns_boundary=True)
    before = document.paragraphs[0]
    before._p.append(_word("bookmarkStart", id="21", name="EndsBeforeBody"))
    anchor._p.insert(1, _word("bookmarkEnd", id="21"))
    anchor._p.insert(2, _word("bookmarkStart", id="22", name="BodyBookmark"))
    anchor._p.append(_word("bookmarkEnd", id="22"))
    append_docx_comment(document, anchor.runs, text="Original review")
    anchor._p.append(_word("bookmarkStart", id="23", name="StartsAfterBody"))
    document.paragraphs[-1]._p.append(_word("bookmarkEnd", id="23"))
    with ZipFile(BytesIO(_saved(document))) as archive:
        original_comments = archive.read("word/comments.xml")
    if ending == "table":
        table = document.add_table(rows=1, cols=1)
        table.cell(0, 0).text = "Generated table"
        elements = [table._element]
    elif ending == "empty":
        elements = []
    else:
        elements = [document.add_paragraph("Generated body")._element]
    _place_docx_body(document, elements, anchor)
    payload = _saved(document)
    reopened = Document(BytesIO(payload))
    with ZipFile(BytesIO(payload)) as archive:
        assert archive.read("word/comments.xml") == original_comments
    significant = []
    for element in reopened._element.body.iter():
        if element.tag in {
            qn("w:bookmarkStart"),
            qn("w:bookmarkEnd"),
            qn("w:commentRangeStart"),
            qn("w:commentRangeEnd"),
            qn("w:commentReference"),
        }:
            significant.append((etree.QName(element).localname, element.get(qn("w:id"))))
        elif element.tag == qn("w:t") and (element.text or "").startswith("Generated"):
            significant.append(("body", None))
    expected = [
        ("bookmarkStart", "21"),
        ("bookmarkEnd", "21"),
        ("bookmarkStart", "22"),
        ("commentRangeStart", "0"),
        *([] if ending == "empty" else [("body", None)]),
        ("commentRangeEnd", "0"),
        ("commentReference", "0"),
        ("bookmarkEnd", "22"),
        ("bookmarkStart", "23"),
        ("bookmarkEnd", "23"),
    ]
    assert significant == expected
    assert len(reopened.sections) == 2
    if ending == "paragraph":
        assert [p.text for p in reopened.paragraphs] == ["Template preface", "Generated body", "Unrelated last section"]
    elif ending == "table":
        assert len(reopened.paragraphs) == 3  # one carrier for section and comment reference


def test_unknown_anchor_subtree_keeps_text_and_original_carrier_properties() -> None:
    document, anchor = _template(anchor_owns_boundary=True)
    anchor.paragraph_format.space_after = Pt(17)
    unknown = etree.SubElement(anchor._p, "{urn:docwen:test:anchor}content")
    text = etree.SubElement(unknown, qn("w:t"))
    text.text = "Preserve this embedded text"
    original = _xml(unknown)
    body = document.add_paragraph("Generated body")
    fill_template(document, {}, [body], anchor)
    reopened = _reopen(document)
    preserved = list(reopened.element.iter("{urn:docwen:test:anchor}content"))
    assert len(preserved) == 1 and _xml(preserved[0]) == original
    carrier = reopened.paragraphs[-2]
    assert carrier.paragraph_format.space_after == Pt(17)
    assert preserved[0].getparent() is carrier._p
    assert "{{body}}" not in reopened.element.xml
    assert len(reopened.sections) == 2


def test_hyperlink_body_marker_is_consumed_without_destroying_wrapper() -> None:
    document = Document()
    anchor = document.add_paragraph()
    hyperlink = _word("hyperlink", anchor="ExistingTarget", history="1")
    run = _word("r")
    text = _word("t")
    text.text = "{{body}}"
    run.append(text)
    hyperlink.append(run)
    anchor._p.append(hyperlink)
    assert anchor.text == "{{body}}"
    fill_template(document, {}, [document.add_paragraph("Generated body")], anchor)
    reopened = _reopen(document)
    links = list(reopened.element.iter(qn("w:hyperlink")))
    assert len(links) == 1
    assert links[0].get(qn("w:anchor")) == "ExistingTarget"
    assert links[0].get(qn("w:history")) == "1"
    assert "{{body}}" not in reopened.element.xml
    assert [paragraph.text for paragraph in reopened.paragraphs] == ["Generated body"]


@pytest.mark.parametrize("wrapped", [False, True])
@pytest.mark.parametrize("before_text", [False, True])
def test_comment_reference_mixed_with_marker_text_keeps_its_side(wrapped: bool, before_text: bool) -> None:
    document = Document()
    anchor = document.add_paragraph("{{body}}")
    reference = _word("commentReference", id="4")
    run = anchor.runs[0]._r
    run.insert(0 if before_text else len(run), reference)
    if wrapped:
        hyperlink = _word("hyperlink", anchor="ExistingTarget")
        hyperlink.append(run)
        anchor._p.append(hyperlink)
    fill_template(document, {}, [document.add_paragraph("Generated body")], anchor)
    reopened = _reopen(document)
    nodes = [node.tag for node in reopened.element.iter() if node.tag in {qn("w:commentReference"), qn("w:t")}]
    assert nodes == ([qn("w:commentReference"), qn("w:t")] if before_text else [qn("w:t"), qn("w:commentReference")])
    assert [paragraph.text for paragraph in reopened.paragraphs] == ["Generated body"]


def test_split_marker_keeps_interleaved_reference_and_range_order() -> None:
    document = Document()
    anchor = document.add_paragraph("{{")
    reference_run = _word("r")
    reference_run.append(_word("commentReference", id="4"))
    anchor._p.append(reference_run)
    anchor._p.append(_word("bookmarkStart", id="22", name="InternalMarker"))
    anchor.add_run("body}}")
    anchor._p.append(_word("bookmarkEnd", id="22"))
    fill_template(document, {}, [document.add_paragraph("Generated body")], anchor)
    reopened = _reopen(document)
    nodes = [
        node.tag
        for node in reopened.element.iter()
        if node.tag in {qn("w:commentReference"), qn("w:bookmarkStart"), qn("w:bookmarkEnd"), qn("w:t")}
    ]
    assert nodes == [qn("w:commentReference"), qn("w:bookmarkStart"), qn("w:t"), qn("w:bookmarkEnd")]


@pytest.mark.parametrize("alias", ["body", "正文"])
@pytest.mark.parametrize("table_ending", [False, True])
@pytest.mark.parametrize("wrapped", [False, True])
def test_clipboard_conversion_preserves_body_ranges_after_template_prefill(
    tmp_path: Path, alias: str, table_ending: bool, wrapped: bool
) -> None:
    document, anchor = _template(anchor_owns_boundary=True)
    anchor.runs[0].text = "{{" + alias + "}}"
    anchor._p.insert(1, _word("bookmarkStart", id="22", name="BodyBookmark"))
    anchor._p.append(_word("bookmarkEnd", id="22"))
    append_docx_comment(document, anchor.runs, text="Original review")
    if wrapped:
        hyperlink = _word("hyperlink", anchor="ExistingTarget", history="1")
        for child in list(anchor._p):
            if child.tag != qn("w:pPr"):
                hyperlink.append(child)
        anchor._p.append(hyperlink)
    document.add_paragraph("{{OtherTemplateField}}")
    template = tmp_path / "template.docx"
    document.save(str(template))
    with ZipFile(template) as archive:
        comments = archive.read("word/comments.xml")
    sections = [_xml(section._sectPr) for section in document.sections]
    paragraph = ClipboardParagraph((ClipboardText("Generated {{literal}} body"),))
    block = ClipboardTable(1, 1, (ClipboardTableCell(0, 0, 1, 1, (paragraph,)),)) if table_ending else paragraph
    source = tmp_path / "source.dwclip"
    source.write_bytes(clipboard_document_to_bytes(ClipboardDocument((block,))))
    context, _workspace = make_context(str(source), options={"template_name": str(template)})

    result = convert_clipboard_document_to_docx(context)

    assert result.success, result.error
    payload = Path(result.artifacts[0].staging_path).read_bytes()
    reopened = Document(BytesIO(payload))
    assert [_xml(section._sectPr) for section in reopened.sections] == sections
    with ZipFile(BytesIO(payload)) as archive:
        assert archive.read("word/comments.xml") == comments
    significant = []
    for node in reopened.element.iter():
        if node.tag in {
            qn("w:bookmarkStart"),
            qn("w:bookmarkEnd"),
            qn("w:commentRangeStart"),
            qn("w:commentRangeEnd"),
            qn("w:commentReference"),
        }:
            significant.append(etree.QName(node).localname)
        elif node.tag == qn("w:t") and node.text == "Generated {{literal}} body":
            significant.append("body")
    assert significant == [
        "bookmarkStart",
        "commentRangeStart",
        "body",
        "commentRangeEnd",
        "commentReference",
        "bookmarkEnd",
    ]
    assert "{{OtherTemplateField}}" not in reopened.element.xml
    assert "{{" + alias + "}}" not in reopened.element.xml


def test_existing_boundary_and_non_body_anchor_fail_before_content_is_moved() -> None:
    document, anchor = _template(anchor_owns_boundary=True)
    original = _xml(document.element)
    with pytest.raises(ValueError, match="already owns"):
        append_section_break_to_paragraph(anchor, document)
    assert _xml(document.element) == original
    cell_anchor = document.add_table(rows=1, cols=1).cell(0, 0).paragraphs[0]
    generated = document.add_paragraph("Not relocated")
    original = _xml(document.element)
    with pytest.raises(ValueError, match="main-document paragraph"):
        splice_body_elements(document, [generated._element], cell_anchor)
    assert _xml(document.element) == original


def test_empty_anchor_carrier_retains_paragraph_change_after_section_properties() -> None:
    document, anchor = _template(anchor_owns_boundary=True)
    change = _word("pPrChange", id="17", author="Template Author")
    change.append(_word("pPr"))
    anchor._p.get_or_add_pPr().append(change)
    original_change = _xml(change)
    fill_template(document, {}, [], anchor)
    reopened = _reopen(document)
    carrier = reopened.paragraphs[1]
    properties = carrier._p.get_or_add_pPr()
    assert [child.tag for child in properties][-2:] == [qn("w:sectPr"), qn("w:pPrChange")]
    assert _xml(properties[-1]) == original_change


def test_converter_saved_output_keeps_template_section_at_body_slot(tmp_path: Path) -> None:
    document, anchor = _template(anchor_owns_boundary=True)
    original = deepcopy(section_properties_at(document, anchor))
    unrelated = _xml(document.sections[-1]._sectPr)
    template = tmp_path / "sections.docx"
    document.save(str(template))
    original_parts = _header_footer_parts(template.read_bytes())
    source = tmp_path / "source.md"
    source.write_text(
        "First generated block\n\n***\n\nSecond generated block\n\n***\n\nThird generated block\n", encoding="utf-8"
    )
    context, _workspace = make_context(str(source), options={"template_name": str(template)})
    result = MdToDocxConverter().convert(context)
    assert result.success, result.error
    payload = Path(result.artifacts[0].staging_path).read_bytes()
    reopened = Document(BytesIO(payload))
    sections = list(reopened.sections)
    assert len(sections) == 4
    assert _xml(sections[0]._sectPr) == _xml(original)
    assert _xml(sections[1]._sectPr) == _continuation_xml(original, "nextPage")
    assert _xml(sections[2]._sectPr) == _continuation_xml(original, "nextPage")
    assert _xml(sections[3]._sectPr) == unrelated
    assert _header_footer_parts(payload) == original_parts


@pytest.mark.parametrize("with_break", [False, True])
def test_implicit_final_section_is_materialized_only_for_requested_split(tmp_path: Path, with_break: bool) -> None:
    document = Document()
    anchor = document.add_paragraph("{{body}}")
    boundary = document._element.body.sectPr
    assert boundary is not None
    document._element.body.remove(boundary)
    original = _xml(document.element)
    assert SectionBreakContext(document, anchor).boundary is None
    assert body_content_extent(document, anchor) == (
        int(Mm(210)) - 2 * int(Mm(25.4)),
        int(Mm(297)) - 2 * int(Mm(25.4)),
    )
    assert _xml(document.element) == original
    template = tmp_path / "implicit-section.docx"
    document.save(str(template))
    source = tmp_path / "source.md"
    source.write_text("First body\n" + ("\n***\n\nSecond body\n" if with_break else ""), encoding="utf-8")
    context, _workspace = make_context(str(source), options={"template_name": str(template)})
    result = MdToDocxConverter().convert(context)
    assert result.success, result.error
    reopened = Document(result.artifacts[0].staging_path)
    assert len(reopened.sections) == (2 if with_break else 0)
    assert "First body" in "\n".join(paragraph.text for paragraph in reopened.paragraphs)
    if with_break:
        first = reopened.sections[0]._sectPr
        assert len(first) == 0  # No guessed page dimensions are written.
        assert reopened.sections[1].start_type == WD_SECTION_START.NEW_PAGE
    else:
        assert reopened._element.body.sectPr is None

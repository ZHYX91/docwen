"""Saved Word templates keep field/SDT ownership and share exact marker spans."""

from __future__ import annotations

from io import BytesIO
from typing import Any

import pytest
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml import etree

from docwen_plugin_markdown.template_errors import TemplatePackageError
from docwen_plugin_markdown.template_ownership import (
    TemplatePlaceholderOwnershipError,
    iter_paragraph_placeholder_spans,
    validate_template_placeholder_ownership,
)

pytestmark = pytest.mark.contract


def _reopen(document: Any) -> Any:
    stream = BytesIO()
    document.save(stream)
    stream.seek(0)
    return Document(stream)


def _snapshot(document: Any) -> dict[str, bytes]:
    return {str(part.partname): part.blob for part in document.part.package.parts}


def _run(parent: Any, text: str, *, tag: str = "w:t") -> Any:
    run = OxmlElement("w:r")
    node = OxmlElement(tag)
    node.text = text
    run.append(node)
    parent.append(run)
    return node


def _field_char(parent: Any, kind: str) -> None:
    run = OxmlElement("w:r")
    character = OxmlElement("w:fldChar")
    character.set(qn("w:fldCharType"), kind)
    run.append(character)
    parent.append(run)


def _simple_field(parent: Any, *, result: str = "cached value", instruction: str = " PAGE ") -> Any:
    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), instruction)
    _run(field, result)
    parent.append(field)
    return field


def _sdt(parent: Any, *, bound: bool) -> Any:
    control = OxmlElement("w:sdt")
    properties = OxmlElement("w:sdtPr")
    if bound:
        binding = OxmlElement("w:dataBinding")
        binding.set(qn("w:xpath"), "/root/name")
        binding.set(qn("w:storeItemID"), "{44444444-4444-4444-4444-444444444444}")
        properties.append(binding)
    content = OxmlElement("w:sdtContent")
    control.extend((properties, content))
    parent.append(control)
    return content


def _textbox(parent: Any) -> Any:
    run = OxmlElement("w:r")
    pict = OxmlElement("w:pict")
    shape = etree.SubElement(pict, "{urn:schemas-microsoft-com:vml}shape")
    textbox = etree.SubElement(shape, "{urn:schemas-microsoft-com:vml}textbox")
    content = OxmlElement("w:txbxContent")
    paragraph = OxmlElement("w:p")
    content.append(paragraph)
    textbox.append(content)
    run.append(pict)
    parent.append(run)
    return paragraph


def _assert_rejected(document: Any, *, key: str = "name", part: str = "/word/document.xml") -> Any:
    document = _reopen(document)
    before = _snapshot(document)
    with pytest.raises(TemplatePlaceholderOwnershipError) as captured:
        validate_template_placeholder_ownership(document)
    error = captured.value
    assert isinstance(error, TemplatePackageError)
    assert error.diagnostic_code == "MD2DOCX-TEMPLATE-PLACEHOLDER-OWNERSHIP"
    assert error.error_type == "invalid_input"
    assert error.key == key
    assert error.part_name == part
    assert error.paragraph_index > 0
    assert error.end > error.start >= 0
    assert part in str(error)
    assert key in str(error)
    assert _snapshot(document) == before
    return error


def _assert_accepted(document: Any) -> Any:
    document = _reopen(document)
    before = _snapshot(document)
    validate_template_placeholder_ownership(document)
    assert _snapshot(document) == before
    return document


@pytest.mark.parametrize("place", ["cache", "instruction"])
def test_simple_field_placeholder_cannot_be_filled_or_silently_ignored(place: str) -> None:
    document = Document()
    _simple_field(
        document.add_paragraph()._p,
        result="{{name}}" if place == "cache" else "cached value",
        instruction=' DOCVARIABLE "{{name}}" ' if place == "instruction" else " PAGE ",
    )
    _assert_rejected(document)


@pytest.mark.parametrize("split_at", [1, 2, 4, 7])
def test_complex_field_cache_split_across_runs_is_rejected(split_at: int) -> None:
    document = Document()
    paragraph = document.add_paragraph()._p
    _field_char(paragraph, "begin")
    _run(paragraph, " MERGEFIELD name ", tag="w:instrText")
    _field_char(paragraph, "separate")
    _run(paragraph, "{{name}}"[:split_at])
    _run(paragraph, "{{name}}"[split_at:])
    _field_char(paragraph, "end")
    _assert_rejected(document)


def test_complex_fields_span_paragraphs_and_nested_end_does_not_close_outer_field() -> None:
    document = Document()
    first = document.add_paragraph()._p
    _field_char(first, "begin")
    _run(first, " IF ", tag="w:instrText")
    _field_char(first, "begin")
    _run(first, " PAGE ", tag="w:instrText")
    _field_char(first, "separate")
    _run(first, "1")
    _field_char(first, "end")
    second = document.add_paragraph()._p
    _field_char(second, "separate")
    _run(second, "{{name}}")
    _field_char(second, "end")
    error = _assert_rejected(document)
    assert error.paragraph_index == 2


def test_complex_field_ownership_continues_into_table_cells() -> None:
    document = Document()
    _field_char(document.add_paragraph()._p, "begin")
    document.add_table(rows=1, cols=1).cell(0, 0).text = "{{name}}"
    _field_char(document.add_paragraph()._p, "end")
    _assert_rejected(document)


@pytest.mark.parametrize("across_paragraphs", [False, True])
def test_complex_field_instruction_markers_are_detected_across_nodes(across_paragraphs: bool) -> None:
    document = Document()
    first = document.add_paragraph()._p
    _field_char(first, "begin")
    _run(first, ' DOCVARIABLE "{{na', tag="w:instrText")
    second = document.add_paragraph()._p if across_paragraphs else first
    _run(second, 'me}}" ', tag="w:instrText")
    _field_char(second, "separate")
    _run(second, "cached value")
    _field_char(second, "end")
    _assert_rejected(document)


def test_unclosed_field_cannot_turn_following_marker_into_editable_text() -> None:
    document = Document()
    _field_char(document.add_paragraph()._p, "begin")
    document.add_paragraph("{{name}}")
    _assert_rejected(document)


def test_malformed_field_without_markers_is_preserved() -> None:
    document = Document()
    document.add_paragraph("{{ordinary}}")
    _field_char(document.add_paragraph()._p, "end")
    _field_char(document.add_paragraph()._p, "begin")
    document.add_paragraph("cached text without a placeholder")
    _assert_accepted(document)


def test_instruction_regions_are_not_joined_across_nested_fields_or_results() -> None:
    document = Document()
    paragraph = document.add_paragraph()._p
    _field_char(paragraph, "begin")
    _run(paragraph, "{{na", tag="w:instrText")
    _field_char(paragraph, "begin")
    _run(paragraph, " PAGE ", tag="w:instrText")
    _field_char(paragraph, "end")
    _run(paragraph, "me}} {{oth", tag="w:instrText")
    _field_char(paragraph, "separate")
    _run(paragraph, "er}}", tag="w:instrText")
    _field_char(paragraph, "end")
    _assert_accepted(document)


@pytest.mark.parametrize("separator", ["text", "simple", "separate", "end"])
def test_unrelated_orphan_instruction_fragments_do_not_form_a_placeholder(separator: str) -> None:
    document = Document()
    paragraph = document.add_paragraph()._p
    _run(paragraph, "{{na", tag="w:instrText")
    if separator == "text":
        _run(paragraph, "live text")
    elif separator == "simple":
        _simple_field(paragraph, result="")
    else:
        _field_char(paragraph, separator)
    _run(paragraph, "me}}", tag="w:instrText")
    _assert_accepted(document)


def test_contiguous_orphan_instruction_marker_is_rejected() -> None:
    document = Document()
    paragraph = document.add_paragraph()._p
    _run(paragraph, "{{na", tag="w:instrText")
    _run(paragraph, "me}}", tag="w:instrText")
    _assert_rejected(document)


def test_orphan_instruction_fragments_are_not_joined_across_control_exit() -> None:
    document = Document()
    paragraph = document.add_paragraph()._p
    content = _sdt(paragraph, bound=False)
    _run(content, "{{na", tag="w:instrText")
    _run(paragraph, "me}}", tag="w:instrText")
    _assert_accepted(document)


@pytest.mark.parametrize("tag", ["w:del", "w:moveFrom"])
def test_inactive_field_end_does_not_close_live_field(tag: str) -> None:
    document = Document()
    paragraph = document.add_paragraph()._p
    _field_char(paragraph, "begin")
    _field_char(paragraph, "separate")
    revision = OxmlElement(tag)
    _field_char(revision, "end")
    paragraph.append(revision)
    _run(paragraph, "{{name}}")
    _field_char(paragraph, "end")
    _assert_rejected(document)


@pytest.mark.parametrize("tag", ["w:del", "w:moveFrom"])
def test_inactive_field_begin_and_text_are_not_part_of_live_template_text(tag: str) -> None:
    document = Document()
    paragraph = document.add_paragraph()._p
    revision = OxmlElement(tag)
    _field_char(revision, "begin")
    _run(revision, "{{historic}}")
    paragraph.append(revision)
    _run(paragraph, "{{name}}")
    document = _assert_accepted(document)
    assert [span.key for span in iter_paragraph_placeholder_spans(document.paragraphs[0])] == ["name"]


@pytest.mark.parametrize("kind", ["simple", "complex", "sdt", "instruction"])
def test_zero_width_owned_region_between_marker_characters_is_rejected(kind: str) -> None:
    document = Document()
    paragraph = document.add_paragraph()._p
    _run(paragraph, "{{na")
    if kind == "simple":
        _simple_field(paragraph, result="")
    elif kind == "complex":
        _field_char(paragraph, "begin")
        _field_char(paragraph, "separate")
        _field_char(paragraph, "end")
    elif kind == "sdt":
        _sdt(paragraph, bound=True)
    else:
        _run(paragraph, " PAGE ", tag="w:instrText")
    _run(paragraph, "me}}")
    _assert_rejected(document)


@pytest.mark.parametrize("inside", ["start", "middle", "end"])
def test_placeholder_crossing_bound_control_boundary_is_rejected(inside: str) -> None:
    document = Document()
    paragraph = document.add_paragraph()._p
    prefix, controlled, suffix = {
        "start": ("", "{{na", "me}}"),
        "middle": ("{{", "name", "}}"),
        "end": ("{{na", "me}}", ""),
    }[inside]
    _run(paragraph, prefix)
    _run(_sdt(paragraph, bound=True), controlled)
    _run(paragraph, suffix)
    _assert_rejected(document)


def test_bound_block_control_owns_descendant_paragraphs() -> None:
    document = Document()
    body = document.element.find(qn("w:body"))
    assert body is not None
    content = _sdt(body, bound=True)
    body.insert(0, content.getparent())
    paragraph = OxmlElement("w:p")
    _run(paragraph, "{{name}}")
    content.append(paragraph)
    _assert_rejected(document)


def test_bound_ancestor_still_owns_nested_unbound_control() -> None:
    document = Document()
    outer = _sdt(document.add_paragraph()._p, bound=True)
    inner = _sdt(outer, bound=False)
    _run(inner, "{{name}}")
    _assert_rejected(document)


def test_binding_attribute_marker_is_not_silently_ignored() -> None:
    document = Document()
    content = _sdt(document.add_paragraph()._p, bound=True)
    _run(content, "cached value")
    binding = content.getparent().find(qn("w:sdtPr") + "/" + qn("w:dataBinding"))
    assert binding is not None
    binding.set(qn("w:xpath"), "/root/{{name}}")
    _assert_rejected(document)


def test_unbound_controls_and_hyperlinks_remain_regular_template_locations() -> None:
    document = Document()
    paragraph = document.add_paragraph()._p
    _run(_sdt(paragraph, bound=False), "{{na")
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("w:anchor"), "bookmark")
    _run(hyperlink, "me}}")
    paragraph.append(hyperlink)
    document = _assert_accepted(document)
    spans = list(iter_paragraph_placeholder_spans(document.paragraphs[0]))
    assert [span.key for span in spans] == ["name"]
    assert "".join(node.text[offset] for node, offset in spans[0].locations) == "{{name}}"


@pytest.mark.parametrize("which", ["header", "footer"])
def test_related_header_footer_caches_receive_part_specific_diagnostics(which: str) -> None:
    document = Document()
    story = getattr(document.sections[0], which)
    _simple_field(story.paragraphs[0]._p, result="{{name}}")
    _assert_rejected(document, part=f"/word/{which}1.xml")


def test_each_header_footer_and_main_story_has_an_independent_field_stack() -> None:
    document = Document()
    _field_char(document.sections[0].header.paragraphs[0]._p, "begin")
    document.sections[0].footer.paragraphs[0].text = "{{name}}"
    document.add_paragraph("{{name}}")
    _assert_accepted(document)


def test_template_without_markers_preserves_all_field_controls_and_binding_information() -> None:
    document = Document()
    paragraph = document.add_paragraph("{{ordinary}}")._p
    _simple_field(paragraph, instruction=" TOC \\h \\z ")
    _field_char(paragraph, "begin")
    _run(paragraph, ' STYLEREF "Original style" ', tag="w:instrText")
    _field_char(paragraph, "separate")
    _run(paragraph, "Original heading")
    _field_char(paragraph, "end")
    _run(_sdt(paragraph, bound=True), "Bound value")
    document.sections[0].footer.paragraphs[0].text = "Footer"
    _assert_accepted(document)


def test_empty_field_adjacent_to_marker_is_allowed() -> None:
    document = Document()
    paragraph = document.add_paragraph()._p
    _simple_field(paragraph, result="")
    _run(paragraph, "{{name}}")
    _field_char(paragraph, "begin")
    _field_char(paragraph, "end")
    _assert_accepted(document)


def test_textbox_field_stack_cannot_escape_into_containing_story() -> None:
    document = Document()
    paragraph = document.add_paragraph()._p
    _field_char(_textbox(paragraph), "begin")
    _run(paragraph, "{{name}}")
    _assert_accepted(document)


def test_textbox_inside_outer_field_is_still_dynamic_content() -> None:
    document = Document()
    paragraph = document.add_paragraph()._p
    _field_char(paragraph, "begin")
    _field_char(paragraph, "separate")
    _run(_textbox(paragraph), "{{name}}")
    _field_char(paragraph, "end")
    _assert_rejected(document)


def test_textbox_unmatched_end_does_not_close_containing_story_field() -> None:
    document = Document()
    paragraph = document.add_paragraph()._p
    _field_char(paragraph, "begin")
    _field_char(_textbox(paragraph), "end")
    _run(paragraph, "{{name}}")
    _field_char(paragraph, "end")
    _assert_rejected(document)


def test_textbox_field_marker_is_rejected_in_its_own_paragraph() -> None:
    document = Document()
    inner = _textbox(document.add_paragraph()._p)
    _field_char(inner, "begin")
    _run(inner, "{{name}}")
    _field_char(inner, "end")
    error = _assert_rejected(document)
    assert error.paragraph_index == 2


def test_bound_control_still_owns_nested_textbox_story() -> None:
    document = Document()
    content = _sdt(document.add_paragraph()._p, bound=True)
    _run(_textbox(content), "{{name}}")
    _assert_rejected(document)


@pytest.mark.parametrize(
    "tag",
    ["w:tab", "w:br", "w:cr", "w:noBreakHyphen", "w:softHyphen", "w:sym", "w:footnoteReference"],
)
def test_shared_span_reader_never_joins_placeholders_through_controls(tag: str) -> None:
    document = Document()
    paragraph = document.add_paragraph()._p
    _run(paragraph, "{{na")
    run = OxmlElement("w:r")
    run.append(OxmlElement(tag))
    paragraph.append(run)
    _run(paragraph, "me}} / {{valid}}")
    document = _assert_accepted(document)
    assert [span.key for span in iter_paragraph_placeholder_spans(document.paragraphs[0])] == ["valid"]


def test_drawing_cannot_join_outer_marker_fragments() -> None:
    document = Document()
    outer = document.add_paragraph("{{na")._p
    _run(_textbox(outer), "{{inner}}")
    _run(outer, "me}}")
    document = _assert_accepted(document)
    paragraphs = list(document.element.iter(qn("w:p")))
    assert list(iter_paragraph_placeholder_spans(paragraphs[0])) == []
    assert [span.key for span in iter_paragraph_placeholder_spans(paragraphs[1])] == ["inner"]


@pytest.mark.parametrize("tag", ["m:oMath", "m:oMathPara"])
def test_equation_cannot_join_outer_marker_fragments(tag: str) -> None:
    document = Document()
    paragraph = document.add_paragraph("{{na")._p
    equation = OxmlElement(tag)
    math = OxmlElement("m:oMath") if tag == "m:oMathPara" else equation
    if math is not equation:
        equation.append(math)
    run = OxmlElement("m:r")
    text = OxmlElement("m:t")
    text.text = "x"
    run.append(text)
    math.append(run)
    paragraph.append(equation)
    _run(paragraph, "me}} / {{valid}}")
    document = _assert_accepted(document)
    assert [span.key for span in iter_paragraph_placeholder_spans(document.paragraphs[0])] == ["valid"]


def test_shared_span_reader_keeps_textbox_paragraph_text_separate() -> None:
    document = Document()
    outer = document.add_paragraph("{{outer}}")._p
    _run(_textbox(outer), "{{inner}}")
    document = _assert_accepted(document)
    paragraphs = list(document.element.iter(qn("w:p")))
    assert [span.key for span in iter_paragraph_placeholder_spans(paragraphs[0])] == ["outer"]
    assert [span.key for span in iter_paragraph_placeholder_spans(paragraphs[1])] == ["inner"]

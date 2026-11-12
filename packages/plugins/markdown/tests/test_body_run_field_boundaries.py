"""Visible body text inherits formatting without rewriting field identity."""

from __future__ import annotations

from copy import deepcopy
from typing import Any

import pytest
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.oxml.text.paragraph import CT_P
from docx.oxml.text.run import CT_R
from docx.shared import Pt, RGBColor
from docx.text.paragraph import Paragraph
from docx.text.run import Run
from lxml import etree

from docwen_core._docx_semantics_v3_ooxml import inline_reference_sdt, inline_sdt, wrap_direct_body_group
from docwen_plugin_markdown.template_utils import BodyRunFormat, extract_body_run_format

pytestmark = pytest.mark.contract


def _body_format(document: Any) -> BodyRunFormat:
    anchor = document.add_paragraph("{{body}}")
    anchor.runs[0].font.name = "Calibri"
    anchor.runs[0].font.size = Pt(11)
    anchor.runs[0].font.color.rgb = RGBColor.from_string("224466")
    result = extract_body_run_format(document)
    assert result is not None
    return result


def _field_run(paragraph: Any, *, marker: str | None = None, text: str | None = None) -> Any:
    run = paragraph.add_run()
    if marker is not None:
        node = OxmlElement("w:fldChar")
        node.set(qn("w:fldCharType"), marker)
        run._r.append(node)
    if text is not None:
        run.add_text(text)
    return run._r


def _assert_body_font(run: Run) -> None:
    assert run.font.name == "Calibri"
    assert run.font.size == Pt(11)
    assert run.font.color.rgb == RGBColor.from_string("224466")


def _without_visual_properties(element: Any) -> bytes:
    snapshot = deepcopy(element)
    for properties in list(snapshot.iter(qn("w:rPr"))):
        properties.getparent().remove(properties)
    return etree.tostring(snapshot)


@pytest.mark.parametrize("carrier", ("reference", "soft-reference", "simple-field"))
def test_visible_carrier_text_inherits_while_field_identity_is_unchanged(carrier: str) -> None:
    document = Document()
    body_format = _body_format(document)
    paragraph = document.add_paragraph("Before ")
    if carrier == "reference":
        wrapper = inline_reference_sdt(
            "opaque-tag", bookmark_name="Target", cached_number="7", heading_number_only=False, alias="Alias"
        )
    elif carrier == "soft-reference":
        wrapper = inline_sdt("opaque-tag", "Title Alias")
    else:
        wrapper = OxmlElement("w:fldSimple")
        wrapper.set(qn("w:instr"), r" REF Target \h ")
        run = OxmlElement("w:r")
        text = OxmlElement("w:t")
        text.text = "7"
        run.append(text)
        wrapper.append(run)
    paragraph._p.append(wrapper)
    paragraph.add_run(" after.")
    original = _without_visual_properties(wrapper)
    visible = []
    for element in wrapper.iter(qn("w:r")):
        assert isinstance(element, CT_R)
        if element.find(qn("w:t")) is not None:
            visible.append(Run(element, paragraph))
    visible[-1].bold = True

    body_format.apply_to(paragraph)

    assert _without_visual_properties(wrapper) == original
    for run in [paragraph.runs[0], *visible, paragraph.runs[-1]]:
        _assert_body_font(run)
    assert visible[-1].bold is True
    for element in wrapper.iter(qn("w:r")):
        if element.find(qn("w:t")) is None:
            assert element.find(qn("w:rPr")) is None


@pytest.mark.parametrize("start", ("beginning", "nested-cache", "after-field", "last-run"))
def test_selected_visible_field_caches_inherit_without_changing_markers(start: str) -> None:
    document = Document()
    body_format = _body_format(document)
    paragraph = document.add_paragraph("Prefix")
    field_runs = [
        _field_run(paragraph, marker="begin"),
        _field_run(paragraph, marker="separate"),
        _field_run(paragraph, text="outer-before"),
        _field_run(paragraph, marker="begin"),
        _field_run(paragraph, marker="separate"),
        _field_run(paragraph, text="inner-cache"),
        _field_run(paragraph, marker="end"),
        _field_run(paragraph, text="outer-after"),
        _field_run(paragraph, marker="end"),
    ]
    paragraph.add_run("Suffix")
    snapshots = [etree.tostring(run) for run in field_runs]
    original = _without_visual_properties(paragraph._p)
    start_index = {"beginning": 0, "nested-cache": 6, "after-field": 10, "last-run": -1}[start]
    first_selected = len(paragraph.runs[:start_index])

    body_format.apply_to(paragraph, start_index=start_index)

    assert _without_visual_properties(paragraph._p) == original
    for index, element in enumerate(field_runs, start=1):
        if element.find(qn("w:fldChar")) is not None or index < first_selected:
            assert etree.tostring(element) == snapshots[index - 1]
        else:
            _assert_body_font(Run(element, paragraph))
    _assert_body_font(paragraph.runs[-1])
    if start == "beginning":
        _assert_body_font(paragraph.runs[0])
    else:
        assert paragraph.runs[0]._r.rPr is None


@pytest.mark.parametrize("identity", ("fldChar", "instrText", "delInstrText"))
def test_identity_run_is_unchanged_even_when_it_also_contains_text(identity: str) -> None:
    document = Document()
    body_format = _body_format(document)
    paragraph = document.add_paragraph()
    run = paragraph.add_run("Mixed payload")
    run._r.append(OxmlElement(f"w:{identity}"))
    original = etree.tostring(run._r)
    paragraph.add_run("Visible suffix")

    body_format.apply_to(paragraph)

    assert etree.tostring(run._r) == original
    _assert_body_font(paragraph.runs[-1])


def test_existing_control_before_selected_range_remains_unchanged() -> None:
    document = Document()
    body_format = _body_format(document)
    paragraph = document.add_paragraph()
    control = inline_reference_sdt(
        "existing-template-control",
        bookmark_name="Earlier",
        cached_number="8",
        heading_number_only=False,
        alias="Earlier",
    )
    paragraph._p.append(control)
    before = etree.tostring(control)
    start_index = len(list(paragraph._p.iter(qn("w:r"))))
    paragraph.add_run("Generated suffix")

    body_format.apply_to(paragraph, start_index=start_index)

    assert etree.tostring(control) == before
    _assert_body_font(paragraph.runs[-1])


def test_enclosing_block_control_does_not_hide_the_current_body_paragraph() -> None:
    document = Document()
    body_format = _body_format(document)
    paragraph = document.add_paragraph("Ordinary body")
    wrap_direct_body_group((paragraph._p,), "body-container")

    body_format.apply_to(paragraph)

    _assert_body_font(paragraph.runs[0])


def test_nested_paragraph_is_only_formatted_when_it_is_the_selected_paragraph() -> None:
    document = Document()
    body_format = _body_format(document)
    outer = document.add_paragraph("Outer body")
    textbox = OxmlElement("w:txbxContent")
    inner_element = OxmlElement("w:p")
    assert isinstance(inner_element, CT_P)
    textbox.append(inner_element)
    outer.add_run()._r.append(textbox)
    inner = Paragraph(inner_element, outer._parent)
    inner.add_run("Nested textbox body")
    original = etree.tostring(inner_element)

    body_format.apply_to(outer)

    _assert_body_font(outer.runs[0])
    assert etree.tostring(inner_element) == original

    body_format.apply_to(inner)

    _assert_body_font(inner.runs[0])

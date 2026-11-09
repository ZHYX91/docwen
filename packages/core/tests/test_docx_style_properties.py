"""Native style inheritance, explicit resets and structural baselines."""

from typing import cast

import pytest
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.styles.style import CharacterStyle, ParagraphStyle

from docwen_core.docx_parsing.break_utils import extract_paragraph_border_info
from docwen_core.docx_parsing.style_properties import resolve_run_formatting

pytestmark = pytest.mark.unit


def _property(owner, name: str, value: str | None = None):
    element = OxmlElement(f"w:{name}")
    if value is not None:
        element.set(qn("w:val"), value)
    owner.append(element)
    return element


def _paragraph_properties(style):
    properties = style.element.find(qn("w:pPr"))
    return properties if properties is not None else _property(style.element, "pPr")


def test_character_style_chain_preserves_formats_and_direct_off_wins():
    doc = Document()
    base = cast(CharacterStyle, doc.styles.add_style("Author Base", WD_STYLE_TYPE.CHARACTER))
    base.font.bold = True
    base.font.underline = True
    child = cast(CharacterStyle, doc.styles.add_style("Author Child", WD_STYLE_TYPE.CHARACTER))
    child.base_style = base
    child.font.italic = True
    child.font.superscript = True
    paragraph = doc.add_paragraph()
    run = paragraph.add_run("format")
    run.style = child

    inherited = resolve_run_formatting(run._r, paragraph)
    assert inherited.bold and inherited.italic and inherited.underline and inherited.superscript

    run.bold = False
    run.italic = False
    run.underline = False
    # python-docx's superscript=False removes the property; an explicit
    # native baseline is what actually cancels an inherited script.
    _property(run._r.get_or_add_rPr(), "vertAlign", "baseline")
    direct = resolve_run_formatting(run._r, paragraph)
    assert not any((direct.bold, direct.italic, direct.underline, direct.superscript, direct.subscript))


def test_style_toggle_is_not_a_direct_property_override():
    doc = Document()
    base = cast(CharacterStyle, doc.styles.add_style("Toggle Base", WD_STYLE_TYPE.CHARACTER))
    base.font.bold = True
    child = cast(CharacterStyle, doc.styles.add_style("Toggle Child", WD_STYLE_TYPE.CHARACTER))
    child.base_style = base
    child.font.bold = True
    paragraph = doc.add_paragraph()
    run = paragraph.add_run("not bold")
    run.style = child

    assert not resolve_run_formatting(run._r, paragraph).bold
    child.font.bold = False  # false in a style leaves the inherited toggle unchanged
    assert resolve_run_formatting(run._r, paragraph).bold
    run.bold = False  # false in direct formatting disables it
    assert not resolve_run_formatting(run._r, paragraph).bold


def test_paragraph_and_structural_character_formats_are_not_authored_emphasis():
    doc = Document()
    paragraph = doc.add_paragraph(style="Heading 1")
    run = paragraph.add_run("heading baseline")
    assert not resolve_run_formatting(run._r, paragraph).bold
    run.bold = True
    assert resolve_run_formatting(run._r, paragraph).bold

    quote = cast(CharacterStyle, doc.styles["Quote Char"])
    run = doc.add_paragraph().add_run("quote baseline")
    run.style = quote
    assert not resolve_run_formatting(run._r, run._parent, structural_character_style=True).italic
    run.italic = True
    assert resolve_run_formatting(run._r, run._parent, structural_character_style=True).italic


def test_noop_character_toggle_does_not_turn_heading_presentation_into_emphasis():
    doc = Document()
    inactive = cast(CharacterStyle, doc.styles.add_style("Inactive Toggle", WD_STYLE_TYPE.CHARACTER))
    inactive.font.bold = False
    paragraph = doc.add_paragraph(style="Heading 1")
    run = paragraph.add_run("heading baseline")
    run.style = inactive
    assert not resolve_run_formatting(run._r, paragraph).bold
    run.bold = True
    assert resolve_run_formatting(run._r, paragraph).bold
    run.bold = False
    assert not resolve_run_formatting(run._r, paragraph).bold


@pytest.mark.parametrize("base_id", [None, "missing-style", "Emphasis"])
def test_independent_or_invalid_base_does_not_inherit_normal(base_id):
    doc = Document()
    normal = cast(ParagraphStyle, doc.styles["Normal"])
    normal.font.hidden = True
    normal_borders = _property(_paragraph_properties(normal), "pBdr")
    _property(normal_borders, "bottom", "single")
    independent = cast(ParagraphStyle, doc.styles.add_style("Independent Root", WD_STYLE_TYPE.PARAGRAPH))
    if base_id is not None:
        _property(independent.element, "basedOn", base_id)
    paragraph = doc.add_paragraph("visible", style=independent)

    assert not resolve_run_formatting(paragraph.runs[0]._r, paragraph).hidden
    assert extract_paragraph_border_info(paragraph) == {}


@pytest.mark.parametrize("style_id", ["unknown-style", "", None])
def test_unknown_paragraph_style_does_not_guess_a_style(style_id):
    doc = Document()
    cast(ParagraphStyle, doc.styles["Normal"]).font.hidden = True
    paragraph = doc.add_paragraph("visible")
    _property(paragraph._p.get_or_add_pPr(), "pStyle", style_id)

    assert not resolve_run_formatting(paragraph.runs[0]._r, paragraph).hidden


def test_style_cycle_stops_at_each_unique_style():
    doc = Document()
    first = cast(CharacterStyle, doc.styles.add_style("Cycle A", WD_STYLE_TYPE.CHARACTER))
    second = cast(CharacterStyle, doc.styles.add_style("Cycle B", WD_STYLE_TYPE.CHARACTER))
    first.base_style = second
    second.base_style = first
    first.font.bold = True
    second.font.underline = True
    paragraph = doc.add_paragraph()
    run = paragraph.add_run("cycle")
    run.style = first

    result = resolve_run_formatting(run._r, paragraph)
    assert result.bold and result.underline


def test_hidden_character_style_can_be_explicitly_disabled():
    doc = Document()
    hidden = cast(CharacterStyle, doc.styles.add_style("Hidden Text", WD_STYLE_TYPE.CHARACTER))
    hidden.font.hidden = True
    paragraph = doc.add_paragraph()
    run = paragraph.add_run("visible after reset")
    run.style = hidden
    assert resolve_run_formatting(run._r, paragraph).hidden
    run.font.hidden = False
    assert not resolve_run_formatting(run._r, paragraph).hidden


def test_style_border_edges_inherit_independently_and_explicit_nil_clears():
    doc = Document()
    base = cast(ParagraphStyle, doc.styles.add_style("Rule Base", WD_STYLE_TYPE.PARAGRAPH))
    borders = _property(_paragraph_properties(base), "pBdr")
    _property(borders, "top", "single")
    _property(borders, "bottom", "double")
    derived = cast(ParagraphStyle, doc.styles.add_style("Rule Child", WD_STYLE_TYPE.PARAGRAPH))
    derived.base_style = base
    paragraph = doc.add_paragraph(style=derived)
    paragraph.paragraph_format.keep_with_next = True
    assert extract_paragraph_border_info(paragraph) == {"top": "single", "bottom": "double"}
    direct_borders = _property(paragraph._p.get_or_add_pPr(), "pBdr")
    _property(direct_borders, "bottom", "nil")
    assert extract_paragraph_border_info(paragraph) == {"top": "single"}

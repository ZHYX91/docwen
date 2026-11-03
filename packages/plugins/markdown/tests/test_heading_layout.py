"""Real DOCX regression coverage for literal heading indentation precedence."""

from __future__ import annotations

import pytest
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.styles.style import ParagraphStyle

from docwen_core.docx_parsing.numbering_index import NumberingIndex
from docwen_plugin_markdown.to_docx.heading_layout import suppress_heading_numbering

pytestmark = pytest.mark.contract


def _set_indent(element, values):
    properties = element.get_or_add_pPr()
    indent = properties.get_or_add_ind()
    for name, value in values.items():
        indent.set(qn(f"w:{name}"), value)


def _install_numbering(document, style, values):
    root = document.part.numbering_part.element
    abstract = OxmlElement("w:abstractNum")
    abstract.set(qn("w:abstractNumId"), "99")
    level = OxmlElement("w:lvl")
    level.set(qn("w:ilvl"), "0")
    for tag, value in (("start", "1"), ("numFmt", "decimal"), ("lvlText", "%1."), ("pStyle", style.style_id)):
        child = OxmlElement(f"w:{tag}")
        child.set(qn("w:val"), value)
        level.append(child)
    properties = OxmlElement("w:pPr")
    indent = OxmlElement("w:ind")
    for name, value in values.items():
        indent.set(qn(f"w:{name}"), value)
    properties.append(indent)
    level.append(properties)
    abstract.append(level)
    root.append(abstract)
    instance = OxmlElement("w:num")
    instance.set(qn("w:numId"), "99")
    reference = OxmlElement("w:abstractNumId")
    reference.set(qn("w:val"), "99")
    instance.append(reference)
    root.append(instance)
    numbering = style.element.get_or_add_pPr().get_or_add_numPr()
    numbering.get_or_add_ilvl().val = 0
    numbering.get_or_add_numId().val = 99


@pytest.mark.parametrize("bidi", [False, True])
@pytest.mark.parametrize(
    ("direct", "style", "base", "numbering", "expected"),
    [
        # A positive explicit first line must not acquire a style's hanging indent.
        ({"firstLine": "360"}, {"hanging": "720"}, {}, None, {"firstLine": "360"}),
        ({"hanging": "360"}, {"firstLine": "720"}, {}, None, {"hanging": "360"}),
        ({"firstLine": "0"}, {"hangingChars": "200"}, {}, None, {"firstLine": "0"}),
        # The nearest style wins the whole dimension, even across units/signs.
        ({}, {"firstLine": "632"}, {"hangingChars": "200"}, None, {"firstLine": "632"}),
        ({}, {"hangingChars": "100"}, {"firstLine": "632"}, None, {"hangingChars": "100"}),
        ({}, {}, {"firstLine": "632"}, {"firstLine": "400"}, {"firstLine": "632"}),
        ({}, {"firstLine": "632"}, {}, {"firstLine": "400"}, {"firstLine": "632"}),
        ({}, {"hanging": "632"}, {}, {"firstLine": "400"}, {"hanging": "632"}),
        ({"firstLineChars": "100"}, {"hanging": "632"}, {}, {"firstLine": "400"}, {"firstLineChars": "100"}),
        # Each dimension resolves independently; level fallback must not skip style values.
        ({}, {"firstLine": "632"}, {}, {"left": "900", "hanging": "400"}, {"firstLine": "632", "left": "900"}),
        (
            {"left": "-360"},
            {"rightChars": "200"},
            {"right": "800", "firstLine": "632"},
            {"leftChars": "300", "right": "900", "hanging": "400"},
            {"left": "-360", "rightChars": "200", "firstLine": "632"},
        ),
        ({"left": "360"}, {"startChars": "200"}, {}, None, {"left": "360"}),
        ({}, {"end": "-360"}, {"rightChars": "200"}, {"right": "900"}, {"end": "-360"}),
        ({}, {"startChars": "200"}, {"left": "720"}, {"start": "900"}, {"startChars": "200"}),
        # Preserve same-layer representations, leaving Word's intra-layer precedence intact.
        ({}, {"firstLine": "632", "firstLineChars": "200"}, {}, None, {"firstLine": "632", "firstLineChars": "200"}),
        (
            {},
            {},
            {},
            {"left": "720", "leftChars": "200", "hanging": "360"},
            {"left": "720", "leftChars": "200", "hanging": "360"},
        ),
        (
            {"firstLine": "632", "hanging": "360"},
            {},
            {},
            {"hangingChars": "100"},
            {"firstLine": "632", "hanging": "360"},
        ),
        ({}, {}, {}, {"firstLine": "400"}, {"firstLine": "400"}),
        ({}, {}, {}, {"hangingChars": "200"}, {"hangingChars": "200"}),
        ({}, {}, {}, None, {}),
    ],
)
def test_heading_indent_dimensions_survive_real_docx_roundtrip(
    tmp_path, direct, style, base, numbering, expected, bidi
):
    document = Document()
    base_style = document.styles.add_style("Heading indent base", WD_STYLE_TYPE.PARAGRAPH)
    heading_style = document.styles.add_style("Heading indent child", WD_STYLE_TYPE.PARAGRAPH)
    assert isinstance(base_style, ParagraphStyle)
    assert isinstance(heading_style, ParagraphStyle)
    heading_style.base_style = base_style
    _set_indent(base_style.element, base)
    _set_indent(heading_style.element, style)
    if numbering is not None:
        # Bind the ancestor, exercising real NumberingIndex style inheritance too.
        _install_numbering(document, base_style, numbering)
    else:
        # Not just an unbound style: the template truly has no numbering part.
        relationship = next(key for key, rel in document.part.rels.items() if rel.reltype == RT.NUMBERING)
        document.part.drop_rel(relationship)
    paragraph = document.add_paragraph("1. Literal heading", style=heading_style)
    _set_indent(paragraph._p, direct)
    direction = OxmlElement("w:bidi")
    direction.set(qn("w:val"), "1" if bidi else "0")
    paragraph._p.get_or_add_pPr().append(direction)
    source = tmp_path / "source.docx"
    document.save(source)
    reopened = Document(source)
    paragraph = reopened.paragraphs[0]
    assert paragraph.style is not None
    resolved = NumberingIndex(reopened).lookup_by_style_id(paragraph.style.style_id)
    if numbering is None:
        assert resolved is None
    else:
        assert resolved is not None and resolved.num_id == "99"
    styles_before = reopened.styles.element.xml
    numbering_before = reopened.part.numbering_part.element.xml if numbering is not None else None
    suppress_heading_numbering(reopened, paragraph)
    # Repeated suppression must not change the effective indent or shared styles.
    once = paragraph._p.xml
    suppress_heading_numbering(reopened, paragraph)
    assert paragraph._p.xml == once
    output = tmp_path / "suppressed.docx"
    reopened.save(output)
    result = Document(output)
    result_properties = result.paragraphs[0]._p.pPr
    assert result_properties is not None
    actual = result_properties.ind
    assert actual is not None
    assert dict(actual.attrib) == {qn(f"w:{name}"): value for name, value in expected.items()}
    for name in ("numId", "ilvl"):
        child = result_properties.find(f"{qn('w:numPr')}/{qn(f'w:{name}')}")
        assert child is not None and child.get(qn("w:val")) == "0"
    assert result.styles.element.xml == styles_before
    result_direction = result_properties.find(qn("w:bidi"))
    assert result_direction is not None
    assert result_direction.get(qn("w:val")) == ("1" if bidi else "0")
    if numbering is not None:
        assert result.part.numbering_part.element.xml == numbering_before

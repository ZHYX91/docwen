"""Heading-number review and layout regression compatibility checks."""

from __future__ import annotations

from zipfile import ZipFile

import pytest
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.styles.style import CharacterStyle

from docwen_plugin_markdown.to_docx.numbering_comments import NUMBERING_REVIEW_TEXT

from .test_numbering_comments import _anchors, _convert, _paragraph_num_id, _style_bytes

pytestmark = pytest.mark.contract


@pytest.mark.parametrize("comments_enabled", [True, False])
@pytest.mark.parametrize("title", ["**标题**：", "[标题](https://example.com)", "`标题`"])
def test_native_reviews_anchor_only_first_visible_heading_character(tmp_path, comments_enabled, title):
    _result, original_path, output_path = _convert(
        tmp_path,
        mode="word_native",
        enabled=comments_enabled,
        source=f"# {title}\n\n正文\n",
    )
    output = Document(str(output_path))
    reviews = [comment for comment in output.comments if comment.text == NUMBERING_REVIEW_TEXT]
    assert len(reviews) == int(comments_enabled)
    if comments_enabled:
        assert _anchors(output)[str(reviews[0].comment_id)] == "标"
    assert _anchors(output)["0"] == "Original comment anchor"
    with ZipFile(original_path) as before, ZipFile(output_path) as after:
        assert before.read("word/_rels/comments.xml.rels") == after.read("word/_rels/comments.xml.rels")


def test_native_empty_heading_does_not_anchor_merged_body_or_existing_template(tmp_path):
    _result, _original, output_path = _convert(tmp_path, mode="word_native", source="#\n\nBody\n")
    output = Document(str(output_path))
    assert [comment.text for comment in output.comments] == ["Original review"]


@pytest.mark.parametrize("comments_enabled", [True, False])
def test_text_number_suppression_preserves_template_numbering_indent(tmp_path, comments_enabled):
    _result, original_path, output_path = _convert(
        tmp_path, mode="text", enabled=comments_enabled, template_native=True, source="# Title\n"
    )
    output = Document(str(output_path))
    heading = next(p for p in output.paragraphs if p.text == "一、Title")
    original = Document(str(original_path))
    root = original.part.numbering_part.element
    instance = next(n for n in root.findall(qn("w:num")) if n.get(qn("w:numId")) == "1")
    reference = instance.find(qn("w:abstractNumId"))
    assert reference is not None
    abstract_id = reference.get(qn("w:val"))
    abstract = next(n for n in root.findall(qn("w:abstractNum")) if n.get(qn("w:abstractNumId")) == abstract_id)
    indent = abstract.find(f"{qn('w:lvl')}/{qn('w:pPr')}/{qn('w:ind')}")
    assert indent is not None and indent.get(qn("w:hanging")) is not None
    actual = heading._p.find(f"{qn('w:pPr')}/{qn('w:ind')}")
    assert actual is not None and dict(actual.attrib) == dict(indent.attrib)
    assert _paragraph_num_id(heading) == "0"
    assert _style_bytes(output, "Heading 1") == _style_bytes(original, "Heading 1")
    with ZipFile(original_path) as before, ZipFile(output_path) as after:
        assert before.read("word/numbering.xml") == after.read("word/numbering.xml")


def test_numbering_indent_override_and_explicit_style_indent_are_preserved():
    from docwen_plugin_markdown.to_docx.heading_layout import suppress_heading_numbering

    document = Document()
    style = document.styles["Heading 1"]
    style_properties = style.element.find(qn("w:pPr"))
    assert style_properties is not None
    num_properties = OxmlElement("w:numPr")
    num_id = OxmlElement("w:numId")
    num_id.set(qn("w:val"), "1")
    num_properties.append(num_id)
    style_properties.append(num_properties)
    style_indent = OxmlElement("w:ind")
    style_indent.set(qn("w:left"), "1440")
    style_properties.append(style_indent)
    numbering = document.part.numbering_part.element
    instance = next(n for n in numbering.findall(qn("w:num")) if n.get(qn("w:numId")) == "1")
    override = OxmlElement("w:lvlOverride")
    override.set(qn("w:ilvl"), "0")
    level = OxmlElement("w:lvl")
    level.set(qn("w:ilvl"), "0")
    properties = OxmlElement("w:pPr")
    indent = OxmlElement("w:ind")
    indent.set(qn("w:left"), "900")
    indent.set(qn("w:hangingChars"), "200")
    properties.append(indent)
    level.append(properties)
    override.append(level)
    instance.append(override)
    paragraph = document.add_heading("一、Title", level=1)
    suppress_heading_numbering(document, paragraph)
    actual = paragraph._p.find(f"{qn('w:pPr')}/{qn('w:ind')}")
    assert actual is not None
    assert actual.get(qn("w:left")) is None  # continues inheriting explicit style value
    assert actual.get(qn("w:hangingChars")) == "200"
    assert _paragraph_num_id(paragraph) == "0"


@pytest.mark.parametrize("prefix", ["\u2060", "\u200b", "\ufeff", " \t\u200e"])
@pytest.mark.parametrize("inherited_hidden", [False, True])
def test_native_comment_skips_invisible_formatting_characters_and_hidden_runs(prefix, inherited_hidden):
    from docwen_plugin_markdown.to_docx.numbering_comments import add_native_numbering_review

    document = Document()
    paragraph = document.add_paragraph()
    hidden = paragraph.add_run("Hidden")
    if inherited_hidden:
        style = document.styles.add_style("Hidden title", WD_STYLE_TYPE.CHARACTER)
        assert isinstance(style, CharacterStyle)
        style.font.hidden = True
        hidden.style = style
    else:
        hidden.font.hidden = True
    title = paragraph.add_run(prefix + "标题")
    title.bold = True
    before = paragraph.text
    assert add_native_numbering_review(document, paragraph)
    assert paragraph.text == before
    assert _anchors(document)["0"] == "标"
    visible = [run for run in paragraph.runs if run.text == "标"]
    assert len(visible) == 1 and visible[0].bold


@pytest.mark.parametrize("comments_enabled", [True, False])
@pytest.mark.parametrize("tail", ["", "\n# Next\n"])
def test_preserved_marker_suppresses_template_number_without_new_review(tmp_path, comments_enabled, tail):
    title = "\u20609.\u2060 Title"
    _result, original_path, output_path = _convert(
        tmp_path,
        mode="text",
        enabled=comments_enabled,
        template_native=True,
        source=f"# {title}\n{tail}",
    )
    output = Document(str(output_path))
    heading = next(p for p in output.paragraphs if p.text == title)
    assert _paragraph_num_id(heading) == "0"
    assert not list(heading._p.iter(qn("w:commentRangeStart")))
    reviews = [c for c in output.comments if c.text == NUMBERING_REVIEW_TEXT]
    assert len(reviews) == int(bool(tail) and comments_enabled)
    if tail:
        following = next(p for p in output.paragraphs if p.text == "二、Next")
        assert _paragraph_num_id(following) == "0"
        if comments_enabled:
            assert _anchors(output)[str(reviews[0].comment_id)] == "二、"
    assert _style_bytes(output, "Heading 1") == _style_bytes(Document(str(original_path)), "Heading 1")

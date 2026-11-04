"""Untrusted heading hints require physical and visible-content agreement."""

from __future__ import annotations

from copy import deepcopy
from io import BytesIO
from typing import Any, cast

import pytest
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.styles.style import CharacterStyle, ParagraphStyle

from docwen_core.docx_parsing.format_features import DocxMarkdownSyntaxConfig
from docwen_plugin_document.shared.markdown_runs import render_paragraph_runs, resolve_hyperlink_target

pytestmark = pytest.mark.unit

_NAME = "DWNav_" + "a" * 32


def _document(*, title: str = "Intro", hint: str = "#intro"):
    document = Document()
    paragraph = document.add_paragraph()
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("w:anchor"), _NAME)
    hyperlink.set(qn("w:tooltip"), hint)
    run = OxmlElement("w:r")
    text = OxmlElement("w:t")
    text.text = "Jump"
    run.append(text)
    hyperlink.append(run)
    paragraph._p.append(hyperlink)
    heading = document.add_heading(title, level=1)
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:name"), _NAME)
    start.set(qn("w:id"), "5")
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), "5")
    heading._p.insert(1, start)
    heading._p.append(end)
    return document, paragraph, hyperlink, heading, start, end


@pytest.mark.parametrize(
    ("title", "hint", "expected"),
    [
        ("Intro", "#intro", "#intro"),
        ("Intro", "#Intro", "#Intro"),
        ("Project Plan", "#project-plan", "#project-plan"),
        ("Project Plan", "#Project Plan", "#Project%20Plan"),
        ("项目计划", "#%E9%A1%B9%E7%9B%AE%E8%AE%A1%E5%88%92", "#%E9%A1%B9%E7%9B%AE%E8%AE%A1%E5%88%92"),
        ("Plan (one)", "#Plan (one)", "#Plan%20%28one%29"),
        (r"Plan\one", r"#Plan\one", "#Plan%5Cone"),
        (
            "x) [other](https://evil.example)",
            "#x) [other](https://evil.example)",
            "#x%29%20%5Bother%5D%28https%3A//evil.example%29",
        ),
    ],
)
def test_verified_heading_target_is_serialized_safely(title: str, hint: str, expected: str) -> None:
    _doc, paragraph, link, _heading, _start, _end = _document(title=title, hint=hint)
    assert resolve_hyperlink_target(paragraph, link) == expected
    assert render_paragraph_runs(paragraph, syntax_config=DocxMarkdownSyntaxConfig()) == f"[Jump]({expected})"


def test_heading_projection_survives_save_reload() -> None:
    document, *_ = _document()
    stream = BytesIO()
    document.save(stream)
    stream.seek(0)
    reopened = Document(stream)
    paragraph = reopened.paragraphs[0]
    link = paragraph._p.find(qn("w:hyperlink"))
    assert resolve_hyperlink_target(paragraph, link) == "#intro"


def test_inherited_outline_and_formatted_heading_text_are_supported() -> None:
    document, paragraph, link, heading, _start, _end = _document()
    inherited = cast(ParagraphStyle, document.styles.add_style("Custom heading", WD_STYLE_TYPE.PARAGRAPH))
    inherited.base_style = cast(ParagraphStyle, document.styles["Heading 1"])
    heading.style = inherited
    heading.runs[0].bold = True
    assert resolve_hyperlink_target(paragraph, link) == "#intro"


@pytest.mark.parametrize(
    "damage",
    [
        "foreign_name",
        "unknown_name",
        "missing_hint",
        "changed_hint",
        "block_hint",
        "encoded_block_hint",
        "control_hint",
        "encoded_control_hint",
        "invalid_utf8_hint",
        "missing_start",
        "missing_end",
        "duplicate_name",
        "duplicate_id",
        "duplicate_end",
        "mismatched_id",
        "cross_paragraph_end",
        "nested_start",
        "nested_navigation",
        "malformed_marker",
        "malformed_text",
        "reversed_range",
        "partial_range",
        "body_style",
        "body_outline_override",
        "changed_title",
        "duplicate_heading",
        "slug_collision",
        "header_duplicate",
        "unsupported_field",
        "misleading_style_name",
        "style_cycle",
    ],
)
def test_ambiguous_or_modified_target_keeps_plain_readable_label(damage: str) -> None:
    document, paragraph, link, heading, start, end = _document()
    if damage == "foreign_name":
        foreign_name = "DW_T_" + "a" * 32
        link.set(qn("w:anchor"), foreign_name)
        start.set(qn("w:name"), foreign_name)
    elif damage == "unknown_name":
        link.set(qn("w:anchor"), "DWNav_" + "b" * 32)
    elif damage == "missing_hint":
        del link.attrib[qn("w:tooltip")]
    elif damage in {
        "changed_hint",
        "block_hint",
        "encoded_block_hint",
        "control_hint",
        "encoded_control_hint",
        "invalid_utf8_hint",
    }:
        hints = {
            "changed_hint": "#Other",
            "block_hint": "#^intro",
            "encoded_block_hint": "#%5Eintro",
            "control_hint": "#intro\n",
            "encoded_control_hint": "#intro%0A",
            "invalid_utf8_hint": "#%FF",
        }
        link.set(qn("w:tooltip"), hints[damage])
    elif damage == "missing_start":
        heading._p.remove(start)
    elif damage == "missing_end":
        heading._p.remove(end)
    elif damage == "duplicate_name":
        duplicate = deepcopy(start)
        duplicate.set(qn("w:id"), "6")
        document.add_paragraph()._p.append(duplicate)
    elif damage == "duplicate_id":
        duplicate = deepcopy(start)
        duplicate.set(qn("w:name"), "other")
        document.add_paragraph()._p.append(duplicate)
    elif damage == "duplicate_end":
        heading._p.append(deepcopy(end))
    elif damage == "mismatched_id":
        end.set(qn("w:id"), "6")
    elif damage == "cross_paragraph_end":
        document.add_paragraph()._p.append(end)
    elif damage == "nested_start":
        heading.runs[0]._r.insert(0, start)
    elif damage == "nested_navigation":
        inner_start = deepcopy(start)
        inner_start.set(qn("w:name"), "DWNav_" + "b" * 32)
        inner_start.set(qn("w:id"), "6")
        inner_end = deepcopy(end)
        inner_end.set(qn("w:id"), "6")
        heading._p.insert(2, inner_start)
        heading._p.insert(len(heading._p) - 1, inner_end)
    elif damage == "malformed_marker":
        start.text = "not a bookmark child"
    elif damage == "malformed_text":
        heading._p[0].tail = "not a paragraph child"
    elif damage == "reversed_range":
        heading._p.remove(end)
        heading._p.insert(1, end)
    elif damage == "partial_range":
        heading._p.remove(start)
        heading._p.insert(2, start)
    elif damage == "body_style":
        heading.style = cast(ParagraphStyle, document.styles["Normal"])
    elif damage == "body_outline_override":
        outline = OxmlElement("w:outlineLvl")
        outline.set(qn("w:val"), "9")
        heading._p.get_or_add_pPr().append(outline)
    elif damage == "changed_title":
        heading.runs[0].text = "Different"
    elif damage == "duplicate_heading":
        document.add_heading("Intro", level=2)
    elif damage == "slug_collision":
        document.add_heading("Intro!", level=2)
    elif damage == "header_duplicate":
        document.sections[0].header.paragraphs[0]._p.append(deepcopy(start))
    elif damage == "unsupported_field":
        heading._p.insert(2, OxmlElement("w:fldSimple"))
    elif damage == "misleading_style_name":
        style = cast(ParagraphStyle, document.styles.add_style("Heading 8 Fake", WD_STYLE_TYPE.PARAGRAPH))
        heading.style = style
    elif damage == "style_cycle":
        style = cast(ParagraphStyle, document.styles.add_style("Circular heading", WD_STYLE_TYPE.PARAGRAPH))
        style.base_style = style
        heading.style = style
    assert resolve_hyperlink_target(paragraph, link) is None
    assert render_paragraph_runs(paragraph, syntax_config=DocxMarkdownSyntaxConfig()) == "Jump"


def test_external_relationship_still_has_precedence_over_untrusted_heading_hint() -> None:
    _document_object, paragraph, link, _heading, _start, _end = _document()
    relation = paragraph.part.relate_to(
        "https://example.com",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )
    link.set(qn("r:id"), relation)
    assert resolve_hyperlink_target(paragraph, link) == "https://example.com"


@pytest.mark.parametrize(
    "hidden_source", ["direct_run", "paragraph_style", "character_style", "base_character_style", "document_default"]
)
def test_hidden_or_uncertain_title_visibility_never_proves_a_target(hidden_source: str) -> None:
    document, paragraph, link, heading, _start, _end = _document()
    hidden = OxmlElement("w:vanish")
    if hidden_source == "direct_run":
        properties = heading.runs[0]._r.get_or_add_rPr()
    elif hidden_source == "paragraph_style":
        properties = cast(Any, document.styles["Heading 1"].element).get_or_add_rPr()
    elif hidden_source in {"character_style", "base_character_style"}:
        style = cast(CharacterStyle, document.styles.add_style("Hidden title", WD_STYLE_TYPE.CHARACTER))
        heading.runs[0].style = style
        if hidden_source == "base_character_style":
            base = cast(CharacterStyle, document.styles.add_style("Hidden base", WD_STYLE_TYPE.CHARACTER))
            style.base_style = base
            style = base
        properties = cast(Any, style.element).get_or_add_rPr()
    else:
        defaults = document.styles.element.find(qn("w:docDefaults"))
        assert defaults is not None
        properties = defaults.find(f"{qn('w:rPrDefault')}/{qn('w:rPr')}")
    assert properties is not None
    properties.append(hidden)
    assert resolve_hyperlink_target(paragraph, link) is None

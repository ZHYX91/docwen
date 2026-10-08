"""Source ownership must preserve visible semantics and typed DOCX notes."""

from __future__ import annotations

from pathlib import Path
from zipfile import ZipFile

import pytest
from lxml import etree

from docwen_plugin_markdown.number_suite_direct_semantics import analyze_markdown_semantics_v3
from docwen_plugin_markdown.to_docx.converter import MdToDocxConverter
from docwen_plugin_markdown.to_docx.notes import normalize_note_syntax, process_md_body_with_notes

from .conftest import make_context

pytestmark = pytest.mark.contract

_DESTINATIONS = [
    "[示例](local(foo)%%draft)",
    r"[示例](local\(foo\)%%draft)",
    '[示例](local(foo) "标题 ) %%draft")',
    "[示例](<local(foo)%%draft>)",
]
_WIKILINKS = [
    "[[Chapter%%Draft]]",
    "[[Chapter|Alias%%Draft]]",
    "[[Chapter<!--Draft]]",
    "[[Chapter|Alias<!--Draft]]",
]


def _source(literal: str) -> str:
    return (
        f"Figure: 中文😀 ^target\n\n{literal}\n\n"
        "See @[[#^target]] @visible. Foot[^f], end[^endnote:e].\n\n"
        "[^f]: Foot.\n[^endnote:e]: End.\n"
    ).replace("\n", "\r\n")


def _analyze(source: str):
    return analyze_markdown_semantics_v3(source, input_id="fixture.md", consumer_profile="number_suite_direct")


def _assert_visible(analysis, source: str, count: int = 1) -> None:
    assert not analysis.has_errors
    references = analysis.projection["references"]
    citations = analysis.projection["citations"]
    assert len(references) == count
    assert len(citations) == count
    assert all(item["resolution_status"] == "resolved" for item in references)
    for item in references + citations:
        span = item["range"]
        assert source[span["start"] : span["end"]] == item["raw"]


@pytest.mark.parametrize("literal", _DESTINATIONS)
def test_balanced_link_metadata_preserves_later_semantic_tokens(literal: str) -> None:
    source = _source(literal)
    _assert_visible(_analyze(source), source)


@pytest.mark.parametrize("literal", _DESTINATIONS)
def test_balanced_link_metadata_preserves_typed_note_domains(literal: str) -> None:
    source = _source(literal)
    assert literal in normalize_note_syntax(source)
    _ast, notes = process_md_body_with_notes(source)
    assert len(notes._footnote_children) == 1
    assert len(notes._endnote_children) == 1


@pytest.mark.parametrize("literal", _WIKILINKS)
def test_wiki_metadata_owns_delimiters_but_link_stays_visible(literal: str) -> None:
    source = _source(literal)
    analysis = _analyze(source)
    _assert_visible(analysis, source)
    [link] = analysis.projection["links"]
    span = link["range"]
    assert source[span["start"] : span["end"]] == literal
    assert link["raw"] == literal


@pytest.mark.parametrize("literal", _WIKILINKS)
def test_wiki_metadata_keeps_later_typed_note_domains(literal: str) -> None:
    source = _source(literal)
    assert literal in normalize_note_syntax(source)
    _ast, notes = process_md_body_with_notes(source)
    assert len(notes._footnote_children) == 1
    assert len(notes._endnote_children) == 1


def test_ordinary_angle_comparison_keeps_visible_reference_and_citation() -> None:
    source = _source("当 x < 2 @[[#^target]] @visible > 0 时成立。")
    _assert_visible(_analyze(source), source, 2)


def test_ordinary_angle_comparison_keeps_typed_note_references() -> None:
    source = "当 x < 2 Foot[^f], end[^endnote:e] > 0 时成立。\n\n[^f]: Foot.\n[^endnote:e]: End.\n"
    _ast, notes = process_md_body_with_notes(source)
    assert len(notes._footnote_children) == 1
    assert len(notes._endnote_children) == 1


@pytest.mark.parametrize("slashes", [1, 2, 3, 4])
@pytest.mark.parametrize("closed", [False, True])
def test_html_comment_escape_parity_controls_visible_semantics(slashes: int, closed: bool) -> None:
    literal = "Text " + "\\" * slashes + "<!-- @[[#^target]] @visible"
    if closed:
        literal += " -->"
    source = _source(literal)
    expected = 2 if slashes % 2 else (1 if closed else 0)
    _assert_visible(_analyze(source), source, expected)


@pytest.mark.parametrize("slashes", [1, 2, 3, 4])
@pytest.mark.parametrize("closed", [False, True])
def test_html_comment_escape_parity_controls_typed_notes(slashes: int, closed: bool) -> None:
    literal = "Text " + "\\" * slashes + "<!-- literal"
    if closed:
        literal += " -->"
    _ast, notes = process_md_body_with_notes(_source(literal))
    expected = 1 if slashes % 2 or closed else 0
    assert len(notes._endnote_children) == expected


@pytest.mark.parametrize(
    "literal", ["%% [示例](local(foo)%%draft)", "%% [[Chapter%%Draft]]", "Text <!-- [[Chapter-->Draft]]"]
)
def test_earlier_comment_still_owns_metadata_closer(literal: str) -> None:
    source = _source(literal)
    _assert_visible(_analyze(source), source)
    _ast, notes = process_md_body_with_notes(source)
    assert len(notes._footnote_children) == 1
    assert len(notes._endnote_children) == 1


def test_orphan_destination_does_not_own_comment_delimiters() -> None:
    source = _source("](local%%) hidden @[[#^target]] @hidden %%")
    _assert_visible(_analyze(source), source)
    assert [item["raw"] for item in _analyze(source).projection["citations"]] == ["@visible"]


def test_valid_html_attributes_and_autolinks_are_literal_metadata() -> None:
    source = _source('<span data-x="@[[#^target]] @hidden %%">Text</span> <https://example.test/%%draft>')
    _assert_visible(_analyze(source), source)


def test_valid_html_attributes_do_not_add_note_references() -> None:
    source = '<span data-x="[^hidden]">Text</span> <https://example.test/[^other]>\n\nBody[^real].\n\n[^real]: Real.\n'
    _ast, notes = process_md_body_with_notes(source)
    assert len(notes._footnote_children) == 1


@pytest.mark.parametrize("literal", ["prefixhttps://example.test/%%draft", "https://%%", r"Text \%%"])
def test_invalid_urls_and_obsidian_percent_escape_do_not_invent_literal_owners(literal: str) -> None:
    source = _source(literal)
    _assert_visible(_analyze(source), source, 0)
    _ast, notes = process_md_body_with_notes(source)
    assert len(notes._endnote_children) == 0


@pytest.mark.parametrize(
    "literal",
    [*_DESTINATIONS, *_WIKILINKS, "![[Chapter%%Draft]]", r"Text \<!-- literal", "当 x < 2 @[[#^target]] > 0 时成立。"],
)
def test_actual_docx_preserves_metadata_following_refs_and_typed_notes(literal: str, tmp_path: Path) -> None:
    source = tmp_path / "metadata-owners.md"
    authored = _source(literal)
    if literal.startswith("![["):
        # A missing embed cannot serve as the Figure's image object. Keep
        # this notes/presentation test separate from that admission contract.
        authored = authored.replace(literal, "Ordinary separator paragraph.\r\n\r\n" + literal)
    original = authored.encode()
    source.write_bytes(original)
    context, _workspace = make_context(
        str(source),
        options={"markdown_extensions": {"input": {"captions_references": True, "typed_endnotes": True}}},
    )
    result = MdToDocxConverter().convert(context)
    assert result.success, result.error
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    with ZipFile(result.artifacts[0].staging_path) as package:
        document = etree.fromstring(package.read("word/document.xml"))
        visible = "".join(node.text or "" for node in document.findall(".//w:t", ns))
        for target in ("Chapter%%Draft", "Chapter<!--Draft"):
            if target in literal:
                assert f"[File not found: {target}]" in visible
        assert len(document.findall(".//w:footnoteReference", ns)) == 1
        assert len(document.findall(".//w:endnoteReference", ns)) == 1
        instructions = document.findall(".//w:instrText", ns)
        assert sum("REF " in (node.text or "") for node in instructions) == (2 if "当 x" in literal else 1)
        for part, element in [("footnotes", "footnote"), ("endnotes", "endnote")]:
            root = etree.fromstring(package.read(f"word/{part}.xml"))
            assert root.xpath(f"count(//w:{element}[@w:id > 0])", namespaces=ns) == 1
    assert source.read_bytes() == original

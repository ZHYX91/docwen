"""Number Suite note syntax interoperability regressions."""

from __future__ import annotations

from pathlib import Path

import pytest

from docwen_plugin_markdown.to_docx.notes import normalize_note_syntax, process_md_body_with_notes

pytestmark = pytest.mark.contract


def test_unicode_note_ids_use_nfc_plus_lowercase_not_casefold() -> None:
    markdown = "First[^Straße], second[^Strasse].\n\n[^Straße]: Sharp-s identity.\n[^Strasse]: Latin ss identity.\n"

    _cleaned_ast, note_ctx = process_md_body_with_notes(markdown)

    assert len(note_ctx._footnote_children) == 2
    assert note_ctx._endnote_children == {}


def test_note_scanner_ignores_protected_markdown_regions() -> None:
    markdown = (
        "Visible[^ok].\n\n"
        "<!-- [^html-comment] -->\n"
        "%% [^obsidian-comment] %%\n"
        "[Link](https://example.test/[^destination])\n"
        '<span data-note="[^attribute]">literal</span>\n'
        "https://example.test/[^raw-url]\n"
        "`[^code]`\n"
        "    [^indented-code]\n\n"
        "[^ok]: Visible definition.\n"
    )

    projection = normalize_note_syntax(markdown)
    _cleaned_ast, note_ctx = process_md_body_with_notes(markdown)

    assert len(note_ctx._footnote_children) == 1
    for literal in (
        "[^html-comment]",
        "[^obsidian-comment]",
        "[^destination]",
        "[^attribute]",
        "[^raw-url]",
        "`[^code]`",
    ):
        assert literal in projection


def test_note_scanner_ignores_multiline_comment_definitions() -> None:
    markdown = (
        "Body[^real].\n\n"
        "<!--\n"
        "[^hidden]: not a definition\n"
        "-->\n"
        "%%\n"
        "[^also-hidden]: not a definition\n"
        "%%\n\n"
        "[^real]: Real definition.\n"
    )

    _cleaned_ast, note_ctx = process_md_body_with_notes(markdown)

    assert len(note_ctx._footnote_children) == 1


@pytest.mark.parametrize("literal", ["%%", "<!--", "`%%`", "`<!--`"])
def test_code_comment_delimiters_do_not_hide_following_typed_notes(literal: str) -> None:
    source = f"~~~text\n{literal}\n~~~\n\nFoot[^f], end[^endnote:e].\n\n[^f]: Foot.\n[^endnote:e]: End.\n"

    projection = normalize_note_syntax(source)
    _ast, note_ctx = process_md_body_with_notes(source)

    assert f"~~~text\n{literal}\n~~~" in projection
    assert len(note_ctx._footnote_children) == 1
    assert len(note_ctx._endnote_children) == 1


@pytest.mark.parametrize("code", ["`%%`", "`<!--`", "> ~~~text\n> %%\n> ~~~", "- Item\n\n  ~~~text\n  %%\n  ~~~"])
def test_inline_and_container_code_do_not_hide_following_typed_notes(code: str) -> None:
    source = f"{code}\n\nFoot[^f], end[^endnote:e].\n\n[^f]: Foot.\n[^endnote:e]: End.\n"
    _ast, note_ctx = process_md_body_with_notes(source)

    assert len(note_ctx._footnote_children) == 1
    assert len(note_ctx._endnote_children) == 1


def test_docx_pipeline_keeps_code_literals_later_fields_and_separate_note_domains(tmp_path: Path) -> None:
    from zipfile import ZipFile

    from lxml import etree

    from docwen_plugin_markdown.to_docx.converter import MdToDocxConverter

    from .conftest import make_context

    source = tmp_path / "combination.md"
    authored = (
        "Table: Sales ^sales\n\n| A | B |\n| --- | --- |\n| 0 | False |\n\n"
        "Code: Literal ^literal\n\n~~~text\n%%\n<!--\n~~~\n\n"
        "See @[[#^sales]] and @[[#^literal]]. Foot[^f], end[^endnote:e], repeated[^f].\n\n"
        "[^f]: Foot **bold** and *italic*.\n[^endnote:e]: End ~~strike~~ and `value`.\n"
    )
    original = authored.encode()
    source.write_bytes(original)
    context, _workspace = make_context(
        str(source),
        options={
            "markdown_extensions": {
                "input": {"structural_tables": True, "captions_references": True, "typed_endnotes": True}
            }
        },
    )

    result = MdToDocxConverter().convert(context)

    assert result.success, result.error
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    with ZipFile(result.artifacts[0].staging_path) as archive:
        document = etree.fromstring(archive.read("word/document.xml"))
        assert document.xpath(".//w:endnoteReference/@w:id", namespaces=ns) == ["1"]
        assert document.xpath(".//w:footnoteReference/@w:id", namespaces=ns) == ["1", "1"]
        instructions = [node.text or "" for node in document.findall(".//w:instrText", ns)]
        assert sum("REF " in value for value in instructions) == 2
        text = "".join(node.text or "" for node in document.findall(".//w:t", ns))
        assert "%%" in text and "<!--" in text
        assert "@[[" not in text
        endnotes = etree.fromstring(archive.read("word/endnotes.xml"))
        assert endnotes.xpath(".//w:strike", namespaces=ns)
    assert source.read_bytes() == original

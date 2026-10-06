"""Number Suite note syntax interoperability regressions."""

from __future__ import annotations

import pytest

from docwen_plugin_markdown.to_docx.notes import normalize_note_syntax, process_md_body_with_notes

pytestmark = pytest.mark.contract


def test_unicode_note_ids_use_nfc_plus_lowercase_not_casefold() -> None:
    markdown = (
        "First[^Straße], second[^Strasse].\n\n"
        "[^Straße]: Sharp-s identity.\n"
        "[^Strasse]: Latin ss identity.\n"
    )

    _cleaned_ast, note_ctx = process_md_body_with_notes(markdown)

    assert len(note_ctx._footnote_children) == 2
    assert note_ctx._endnote_children == {}


def test_note_scanner_ignores_protected_markdown_regions() -> None:
    markdown = (
        "Visible[^ok].\n\n"
        "<!-- [^html-comment] -->\n"
        "%% [^obsidian-comment] %%\n"
        "[Link](https://example.test/[^destination])\n"
        "<span data-note=\"[^attribute]\">literal</span>\n"
        "https://example.test/[^raw-url]\n"
        "`[^code]`\n\n"
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

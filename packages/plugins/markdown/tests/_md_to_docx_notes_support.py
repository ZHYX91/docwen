"""Foundation tests for MD→DOCX footnote/endnote model.

Covers: F-F1-022 (NoteContext), F-F1-023 (OOXML element creation),
F-F1-024 (reference runs), F-F1-027 (inline reference rendering),
F-F3-007 (process_md_body_with_notes pipeline).
"""

from __future__ import annotations

import zipfile
from pathlib import Path
from typing import Any

import pytest
from docx import Document
from lxml import etree

from docwen_plugin_markdown.mistune_extensions import parse_markdown_text
from docwen_plugin_markdown.renderer import MdToDocxRenderer
from docwen_plugin_markdown.to_docx.notes import (
    WML_NS,
    NoteContext,
    NoteWritebackError,
    _create_endnote_element,
    _create_endnote_ref_run,
    _create_footnote_element,
    _create_footnote_ref_run,
    extract_notes_from_ast,
    normalize_note_syntax,
    process_md_body_with_notes,
)

pytestmark = pytest.mark.contract

NSMAP = {"w": WML_NS}


def _q(name: str) -> str:
    """Qualify a tag name with the WordprocessingML namespace."""
    return f"{{{WML_NS}}}{name}"


def _render_plain(doc) -> str:
    """Return the plain-text content of a python-docx Document."""
    return "\n".join(p.text for p in doc.paragraphs)


def _mk_children(text: str) -> list[list[dict[str, Any]]]:
    """Build a single-paragraph inline-children list from plain text."""
    return [[{"type": "text", "raw": text}]]


def _def_plain(text: str) -> str:
    """Extract plain text from the first paragraph's first text node."""
    return text


MD_WITH_FOOTNOTES = """# Document Title

This is a paragraph with a footnote[^1] reference.

Another paragraph with multiple[^2] footnotes[^3] inline.

## Section Two

More content with an endnote[^endnote:x] reference.

Final text.

[^1]: First footnote definition.
[^2]: Second footnote with **bold** text.
[^3]: Third footnote
    with multiline content.
[^endnote:x]: Endnote definition goes here.
"""

MD_WITH_FOOTNOTES_ONLY = """# Simple Doc

Text with one note[^a].

[^a]: Just a footnote.
"""

MD_NO_NOTES = """# Plain Doc

Just a paragraph with no notes at all.

Another line.
"""


def _list_zip_entries(docx_path: str) -> set[str]:
    """Return the set of entry names in a DOCX ZIP."""
    with zipfile.ZipFile(docx_path, "r") as zf:
        return set(zf.namelist())


def _read_zip_entry(docx_path: str, name: str) -> bytes:
    """Read a raw entry from a DOCX ZIP."""
    with zipfile.ZipFile(docx_path, "r") as zf:
        return zf.read(name)


MD_FOOTNOTE_ONLY = """# Test

Para with a footnote[^f1].

[^f1]: Footnote body text here.
"""

MD_ENDNOTE_ONLY = """# Test

Para with an endnote[^endnote:e1].

[^endnote:e1]: Endnote body text here.
"""

MD_MIXED_NOTES = """# Mixed

Footnote[^footnote:a] and endnote[^endnote:b].

[^footnote:a]: A footnote.
[^endnote:b]: An endnote.
"""

MD_MULTI_FOOTNOTES = """# Multi

A[^1] B[^2] C[^3].

[^1]: First note.
[^2]: Second note.
[^3]: Third note with
    multiline content.
"""

__all__ = (
    "MD_ENDNOTE_ONLY",
    "MD_FOOTNOTE_ONLY",
    "MD_MIXED_NOTES",
    "MD_MULTI_FOOTNOTES",
    "MD_NO_NOTES",
    "MD_WITH_FOOTNOTES",
    "MD_WITH_FOOTNOTES_ONLY",
    "WML_NS",
    "Any",
    "Document",
    "MdToDocxRenderer",
    "NoteContext",
    "NoteWritebackError",
    "Path",
    "_create_endnote_element",
    "_create_endnote_ref_run",
    "_create_footnote_element",
    "_create_footnote_ref_run",
    "_list_zip_entries",
    "_mk_children",
    "_q",
    "_read_zip_entry",
    "_render_plain",
    "etree",
    "extract_notes_from_ast",
    "normalize_note_syntax",
    "parse_markdown_text",
    "process_md_body_with_notes",
    "pytest",
    "pytestmark",
)

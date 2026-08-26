"""Tests for footnote/endnote style ID resolution (plan-footnote-styles).

Covers:
- ``_resolve_style_id_by_name``: lookup, case-insensitivity, not-found
- ``NoteContext.resolve_note_styles``: blank-doc defaults, custom styleIds
- ``MdToDocxRenderer.render``: calls resolve_note_styles before dispatch
"""

from __future__ import annotations

from typing import Any
from unittest.mock import MagicMock, patch

import pytest
from lxml import etree

from docwen_plugin_markdown.to_docx.notes import (
    _NOTE_STYLE_NAMES,
    _WORD_NS,
    NoteContext,
    _resolve_style_id_by_name,
)

WML_NS = _WORD_NS

pytestmark = pytest.mark.unit


# ── Helpers ──────────────────────────────────────────────────────────────


def _mk_children(text: str) -> list[list[dict[str, Any]]]:
    """Build a single-paragraph inline-children list from plain text."""
    return [[{"type": "text", "raw": text}]]


def _make_style_element(style_id: str, style_name: str) -> etree._Element:
    """Build a minimal ``w:style`` element with a ``w:name`` child."""
    el = etree.Element(f"{{{WML_NS}}}style", {f"{{{WML_NS}}}styleId": style_id})
    etree.SubElement(el, f"{{{WML_NS}}}name", {f"{{{WML_NS}}}val": style_name})
    return el


def _make_mock_style(style_id: str, style_name: str):
    """Build a mock style object whose ``_element`` is an lxml element."""
    style = MagicMock()
    style._element = _make_style_element(style_id, style_name)
    return style


def _make_mock_doc(styles: list):
    """Build a mock python-docx Document with given styles.

    Each item in *styles* is a mock style (from ``_make_mock_style``).
    """
    doc = MagicMock()
    doc.styles = styles
    return doc


# ── _resolve_style_id_by_name ────────────────────────────────────────────


class TestResolveStyleIdByName:
    """Unit tests for ``_resolve_style_id_by_name``."""

    def test_returns_style_id_when_name_matches(self):
        """styleId returned when w:name matches exactly."""
        doc = _make_mock_doc(
            [
                _make_mock_style("FootnoteText", "footnote text"),
            ]
        )
        result = _resolve_style_id_by_name(doc, "footnote text")
        assert result == "FootnoteText"

    def test_case_insensitive_match(self):
        """Matching is case-insensitive."""
        doc = _make_mock_doc(
            [
                _make_mock_style("a5", "Footnote Text"),
            ]
        )
        result = _resolve_style_id_by_name(doc, "footnote text")
        assert result == "a5"

    def test_returns_none_when_not_found(self):
        """Returns None when no style with that name exists."""
        doc = _make_mock_doc(
            [
                _make_mock_style("Heading1", "heading 1"),
            ]
        )
        result = _resolve_style_id_by_name(doc, "footnote text")
        assert result is None

    def test_returns_none_when_no_styles(self):
        """Returns None for a document with no styles."""
        doc = _make_mock_doc([])
        result = _resolve_style_id_by_name(doc, "footnote text")
        assert result is None

    def test_skips_styles_without_name_element(self):
        """Styles without a w:name element are skipped."""
        el = etree.Element(f"{{{WML_NS}}}style", {f"{{{WML_NS}}}styleId": "Foo"})
        # No w:name child
        style = MagicMock()
        style._element = el
        doc = _make_mock_doc([style])
        result = _resolve_style_id_by_name(doc, "footnote text")
        assert result is None

    def test_finds_match_among_many_styles(self):
        """Finds the correct styleId among multiple styles."""
        doc = _make_mock_doc(
            [
                _make_mock_style("Heading1", "heading 1"),
                _make_mock_style("Normal", "Normal"),
                _make_mock_style("a5", "footnote text"),
                _make_mock_style("Heading2", "heading 2"),
            ]
        )
        result = _resolve_style_id_by_name(doc, "footnote text")
        assert result == "a5"

    def test_no_style_id_returns_none(self):
        """Style without a styleId attribute returns None."""
        el = etree.Element(f"{{{WML_NS}}}style")
        etree.SubElement(el, f"{{{WML_NS}}}name", {f"{{{WML_NS}}}val": "footnote text"})
        style = MagicMock()
        style._element = el
        doc = _make_mock_doc([style])
        result = _resolve_style_id_by_name(doc, "footnote text")
        assert result is None


# ── NoteContext.resolve_note_styles ──────────────────────────────────────


class TestResolveNoteStyles:
    """Tests for ``NoteContext.resolve_note_styles(doc)``."""

    def test_blank_doc_retains_english_defaults(self):
        """Empty document: no styles → English defaults kept."""
        ctx = NoteContext()
        doc = _make_mock_doc([])
        ctx.resolve_note_styles(doc)

        assert ctx.footnote_text_style == "FootnoteText"
        assert ctx.footnote_ref_style == "FootnoteReference"
        assert ctx.endnote_text_style == "EndnoteText"
        assert ctx.endnote_ref_style == "EndnoteReference"

    def test_resolves_all_four_styles_from_doc(self):
        """All four note style IDs are resolved from the document."""
        ctx = NoteContext()
        doc = _make_mock_doc(
            [
                _make_mock_style("a1", "footnote text"),
                _make_mock_style("a2", "footnote reference"),
                _make_mock_style("a3", "endnote text"),
                _make_mock_style("a4", "endnote reference"),
            ]
        )
        ctx.resolve_note_styles(doc)

        assert ctx.footnote_text_style == "a1"
        assert ctx.footnote_ref_style == "a2"
        assert ctx.endnote_text_style == "a3"
        assert ctx.endnote_ref_style == "a4"

    def test_partial_resolve_keeps_defaults_for_unmatched(self):
        """Only matched styles are updated; unmatched keep defaults."""
        ctx = NoteContext()
        doc = _make_mock_doc(
            [
                _make_mock_style("a5", "footnote text"),
                # footnote reference, endnote text, endnote reference not present
            ]
        )
        ctx.resolve_note_styles(doc)

        assert ctx.footnote_text_style == "a5"
        assert ctx.footnote_ref_style == "FootnoteReference"  # default kept
        assert ctx.endnote_text_style == "EndnoteText"  # default kept
        assert ctx.endnote_ref_style == "EndnoteReference"  # default kept

    def test_does_not_change_other_attributes(self):
        """resolve_note_styles only touches the four style attrs."""
        ctx = NoteContext()
        ctx._footnote_children = {"a": _mk_children("Content A")}
        doc = _make_mock_doc(
            [
                _make_mock_style("cn_fnt", "footnote text"),
            ]
        )
        ctx.resolve_note_styles(doc)

        # Style updated
        assert ctx.footnote_text_style == "cn_fnt"
        # Other state untouched
        assert ctx._footnote_children == {"a": _mk_children("Content A")}
        assert ctx.has_footnotes

    def test_mapping_covers_all_note_style_attrs(self):
        """_NOTE_STYLE_NAMES has entries for all four style attributes."""
        ctx = NoteContext()
        for attr_name in _NOTE_STYLE_NAMES.values():
            # Each attr_name must exist on a fresh NoteContext
            assert hasattr(ctx, attr_name), f"NoteContext missing attribute: {attr_name}"


# ── MdToDocxRenderer.render integration ─────────────────────────────────


class TestRendererResolveNoteStyles:
    """Tests that ``MdToDocxRenderer.render`` calls ``resolve_note_styles``."""

    def test_render_calls_resolve_note_styles(self):
        """render() calls resolve_note_styles(self._doc) when note_ctx provided."""
        from docwen_plugin_markdown.renderer import MdToDocxRenderer

        ctx = NoteContext()
        doc = _make_mock_doc(
            [
                _make_mock_style("a5", "footnote text"),
            ]
        )

        renderer = MdToDocxRenderer(doc, note_ctx=ctx)
        with patch.object(ctx, "resolve_note_styles") as mock_resolve:
            renderer.render([])
            mock_resolve.assert_called_once_with(doc)

    def test_render_skips_resolve_when_no_note_ctx(self):
        """render() does not call resolve_note_styles when note_ctx is None."""
        from docwen_plugin_markdown.renderer import MdToDocxRenderer

        doc = _make_mock_doc([])
        renderer = MdToDocxRenderer(doc, note_ctx=None)
        # Should not raise — no resolve_note_styles call attempted
        result = renderer.render([])
        assert result == []

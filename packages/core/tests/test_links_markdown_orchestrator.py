"""Tests for ``process_markdown_links`` orchestrator, non-embed link
processing (``_process_non_embed_links``), and wiki embed patterns.

Covers:
- F-H2-019: ``process_markdown_links`` orchestrator
- F-H2-020: ``_process_non_embed_links``
- F-H2-023: ``WIKI_EMBED_PATTERN``
- F-H2-024: ``WIKI_EMBED_SIZE_PATTERN``
- F-H2-031: escaped-pipe restoration inside parsed wiki-link fields
"""

from __future__ import annotations

import re
from dataclasses import replace
from pathlib import Path

import pytest

from docwen_core.export_semantics import LinkRuntimeConfig
from docwen_core.links import (
    TABLE_CELL_BR_TOKEN,
    WIKI_EMBED_PATTERN,
    WIKI_EMBED_SIZE_PATTERN,
    _process_non_embed_links,
    process_markdown_links,
    restore_table_safe_breaks,
)
from docwen_core.links._non_embed import _unescape_pipe

pytestmark = pytest.mark.unit

_ORCHESTRATOR_POLICY = replace(
    LinkRuntimeConfig(),
    max_depth=10,
    non_embed_wiki_mode="keep",
    non_embed_markdown_mode="keep",
)


@pytest.mark.parametrize(
    ("value", "expected"),
    [
        (r"plain", "plain"),
        (r"target\|display", "target|display"),
        (r"target\\|display", r"target\|display"),
        ("", ""),
    ],
)
def test_wiki_link_fields_restore_each_escaped_pipe(value: str, expected: str) -> None:
    """F-H2-031: parsed wiki target/display text restores escaped pipes exactly once."""
    assert _unescape_pipe(value) == expected


# ═══════════════════════════════════════════════════════════════════════════
# Helpers
# ═══════════════════════════════════════════════════════════════════════════


def _write(path: Path, text: str) -> str:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")
    return str(path)


# ═══════════════════════════════════════════════════════════════════════════
# WIKI_EMBED_PATTERN  (F-H2-023)
# ═══════════════════════════════════════════════════════════════════════════


class TestWikiEmbedPattern:
    """Verify that the canonical ``WIKI_EMBED_PATTERN`` matches all expected
    wiki-embed syntax variants."""

    _RE = re.compile(WIKI_EMBED_PATTERN)

    def test_simple_embed(self) -> None:
        m = self._RE.search("![[file.md]]")
        assert m is not None
        assert m.group(1).strip() == "file.md"
        assert m.group(2) is None

    def test_embed_with_display_text(self) -> None:
        m = self._RE.search("![[file.md|display text]]")
        assert m is not None
        assert m.group(1).strip() == "file.md"
        assert m.group(2).strip() == "display text"

    def test_embed_with_heading(self) -> None:
        m = self._RE.search("![[file.md#section]]")
        assert m is not None
        assert m.group(1).strip() == "file.md#section"

    def test_embed_with_heading_and_display(self) -> None:
        m = self._RE.search("![[file.md#section|alt]]")
        assert m is not None
        assert m.group(1).strip() == "file.md#section"
        assert m.group(2).strip() == "alt"

    def test_embed_with_block_id(self) -> None:
        m = self._RE.search("![[file.md#^abc123]]")
        assert m is not None
        assert m.group(1).strip() == "file.md#^abc123"

    def test_embed_with_image(self) -> None:
        m = self._RE.search("![[photo.png]]")
        assert m is not None
        assert m.group(1).strip() == "photo.png"

    def test_embed_with_dimension_display(self) -> None:
        m = self._RE.search("![[photo.png|200x150]]")
        assert m is not None
        assert m.group(1).strip() == "photo.png"
        assert m.group(2).strip() == "200x150"

    def test_embed_with_escaped_pipe_in_target(self) -> None:
        """Backslash-escaped pipe before the separator is treated as the
        separator itself (``\\|`` → literal ``|`` in display).  The target
        stops before the escaped pipe."""
        m = self._RE.search(r"![[target\|display]]")
        assert m is not None
        assert m.group(1).strip() == "target"
        assert m.group(2).strip() == "display"

    def test_non_embed_wiki_link_not_matched(self) -> None:
        """``[[link]]`` without ``!`` is NOT matched."""
        assert self._RE.search("[[plain link]]") is None
        assert self._RE.search("[[link|text]]") is None

    def test_markdown_image_syntax_not_matched(self) -> None:
        """Standard ``![alt](url)`` is NOT matched by the wiki pattern."""
        assert self._RE.search("![alt](img.png)") is None


# ═══════════════════════════════════════════════════════════════════════════
# WIKI_EMBED_SIZE_PATTERN  (F-H2-024)
# ═══════════════════════════════════════════════════════════════════════════


class TestWikiEmbedSizePattern:
    """Verify that ``WIKI_EMBED_SIZE_PATTERN`` matches wiki embeds with
    explicit pixel dimensions."""

    _RE = re.compile(WIKI_EMBED_SIZE_PATTERN)

    def test_width_and_height(self) -> None:
        m = self._RE.search("![[img.png|200x150]]")
        assert m is not None
        assert m.group(1).strip() == "img.png"
        assert m.group(2) == "200"
        assert m.group(3) == "150"

    def test_width_only(self) -> None:
        m = self._RE.search("![[img.png|200]]")
        assert m is not None
        assert m.group(1).strip() == "img.png"
        assert m.group(2) == "200"
        assert m.group(3) is None

    def test_width_with_trailing_x_not_matched(self) -> None:
        """``x`` without trailing digits is not matched — the pattern requires
        at least one digit after ``x``."""
        assert self._RE.search("![[img.png|200x]]") is None

    def test_embed_without_dimensions_not_matched(self) -> None:
        """A plain embed without dimensions is NOT matched by the size pattern."""
        assert self._RE.search("![[file.md]]") is None
        assert self._RE.search("![[file.md|text]]") is None

    def test_non_embed_not_matched(self) -> None:
        assert self._RE.search("[[img.png|200x150]]") is None


# ═══════════════════════════════════════════════════════════════════════════
# process_markdown_links orchestrator  (F-H2-019)
# ═══════════════════════════════════════════════════════════════════════════


class TestProcessMarkdownLinksOrchestrator:
    """Full-chain tests for the ``process_markdown_links`` orchestrator."""

    def test_chains_embed_then_non_embed(self, tmp_path: Path) -> None:
        """Embed resolution runs first, then non-embed processing."""
        src = _write(tmp_path / "doc.md", "")
        _write(tmp_path / "embedded.md", "# Embedded\n\nContent.\n")
        _write(tmp_path / "target.md", "# Target\n\nTarget content.\n")

        text = "![[embedded.md]]\n\nSee [[target.md|Target Page]] for more."

        result = process_markdown_links(
            text,
            src,
            link_config=replace(
                _ORCHESTRATOR_POLICY,
                non_embed_wiki_mode="extract_text",
            ),
        )

        # Embed was resolved
        assert "Content." in result
        assert "![[embedded.md]]" not in result
        # Non-embed wiki link was processed (extract_text → display text)
        assert "Target Page" in result
        assert "[[target.md|Target Page]]" not in result

    def test_embed_with_nested_and_non_embed(self, tmp_path: Path) -> None:
        """Orchestrator handles nested embeds + non-embed links in one pass."""
        src = _write(tmp_path / "doc.md", "")
        _write(tmp_path / "outer.md", "# Outer\n\n![[inner.md]]\n\nSee [[ref.md|Ref]]")
        _write(tmp_path / "inner.md", "# Inner\n\nDeep.\n")
        _write(tmp_path / "ref.md", "# Ref\n\nReference.\n")

        text = "![[outer.md]]"

        result = process_markdown_links(
            text,
            src,
            link_config=replace(
                _ORCHESTRATOR_POLICY,
                non_embed_wiki_mode="extract_text",
            ),
        )

        assert "Deep." in result
        # The non-embed wiki link inside outer.md → extract_text
        assert "Ref" in result
        assert "[[ref.md|Ref]]" not in result

    def test_table_safe_embedded_markdown_flattens_newlines_and_escapes_pipes(self, tmp_path: Path) -> None:
        """table_safe mode preserves embedded Markdown inside a single pipe-table cell."""
        src = _write(tmp_path / "doc.md", "")
        _write(tmp_path / "embedded.md", "First line\nSecond | value")

        result = process_markdown_links(
            "| Col |\n| --- |\n| ![[embedded.md]] |\n",
            src,
            link_config=_ORCHESTRATOR_POLICY,
            table_safe=True,
        )

        assert TABLE_CELL_BR_TOKEN not in result
        assert re.search(
            r"\| First line\{\{DOCWEN_BR@[A-Za-z0-9_-]{32}\.[0-9a-f]{16}\}\}"
            r"Second \\\| value \|",
            result,
        )
        assert "First line\nSecond \\| value" in restore_table_safe_breaks(result)

    def test_no_links_passthrough(self, tmp_path: Path) -> None:
        """Plain text without links passes through unchanged."""
        src = _write(tmp_path / "doc.md", "")
        text = "# Title\n\nJust some text.\n"
        assert process_markdown_links(text, src, link_config=_ORCHESTRATOR_POLICY) == text

    def test_empty_text(self, tmp_path: Path) -> None:
        src = _write(tmp_path / "doc.md", "")
        assert process_markdown_links("", src, link_config=_ORCHESTRATOR_POLICY) == ""

    def test_embed_missing_file_placeholder(self, tmp_path: Path) -> None:
        """Missing embedded file produces a placeholder in orchestrator mode."""
        src = _write(tmp_path / "doc.md", "")
        text = "![[missing.md]]"
        result = process_markdown_links(text, src, link_config=_ORCHESTRATOR_POLICY)
        assert "File not found" in result

    def test_wiki_mode_keep(self, tmp_path: Path) -> None:
        """With wiki_mode='keep', non-embed wiki links are untouched."""
        src = _write(tmp_path / "doc.md", "")
        text = "See [[page.md|Page]] for info."
        result = process_markdown_links(text, src, link_config=_ORCHESTRATOR_POLICY)
        assert "[[page.md|Page]]" in result

    def test_wiki_mode_extract_text(self, tmp_path: Path) -> None:
        """wiki_mode='extract_text' keeps only display text."""
        src = _write(tmp_path / "doc.md", "")
        text = "See [[page.md|Page Title]] for info."
        result = process_markdown_links(
            text,
            src,
            link_config=replace(
                _ORCHESTRATOR_POLICY,
                non_embed_wiki_mode="extract_text",
            ),
        )
        assert "Page Title" in result
        assert "[[page.md|Page Title]]" not in result

    def test_wiki_mode_remove(self, tmp_path: Path) -> None:
        """wiki_mode='remove' drops non-embed wiki links."""
        src = _write(tmp_path / "doc.md", "")
        text = "See [[page.md]] for info."
        result = process_markdown_links(
            text,
            src,
            link_config=replace(
                _ORCHESTRATOR_POLICY,
                non_embed_wiki_mode="remove",
            ),
        )
        assert "[[page.md]]" not in result
        assert "for info." in result

    def test_wiki_mode_resolve(self, tmp_path: Path) -> None:
        """wiki_mode='resolve' resolves targets to file paths."""
        src = _write(tmp_path / "doc.md", "")
        _write(tmp_path / "page.md", "# Page\n")
        text = "See [[page.md|Page]] for info."
        result = process_markdown_links(
            text,
            src,
            link_config=replace(
                _ORCHESTRATOR_POLICY,
                non_embed_wiki_mode="resolve",
            ),
        )
        assert "[Page](<page.md>)" in result
        assert "[[page.md|Page]]" not in result

    def test_markdown_mode_keep(self, tmp_path: Path) -> None:
        """markdown_mode='keep' leaves standard links untouched."""
        src = _write(tmp_path / "doc.md", "")
        text = "See [example](https://example.com) for more."
        result = process_markdown_links(text, src, link_config=_ORCHESTRATOR_POLICY)
        assert "[example](https://example.com)" in result

    def test_markdown_mode_extract_text(self, tmp_path: Path) -> None:
        """markdown_mode='extract_text' keeps only display text."""
        src = _write(tmp_path / "doc.md", "")
        text = "See [example](https://example.com) for more."
        result = process_markdown_links(
            text,
            src,
            link_config=replace(
                _ORCHESTRATOR_POLICY,
                non_embed_markdown_mode="extract_text",
            ),
        )
        assert "example" in result
        assert "[example](https://example.com)" not in result

    def test_markdown_mode_remove(self, tmp_path: Path) -> None:
        """markdown_mode='remove' drops standard markdown links."""
        src = _write(tmp_path / "doc.md", "")
        text = "See [example](https://example.com) for more."
        result = process_markdown_links(
            text,
            src,
            link_config=replace(
                _ORCHESTRATOR_POLICY,
                non_embed_markdown_mode="remove",
            ),
        )
        assert "[example](https://example.com)" not in result
        assert "for more." in result

    def test_code_blocks_protected(self, tmp_path: Path) -> None:
        """Links inside fenced code blocks are not processed."""
        src = _write(tmp_path / "doc.md", "")
        text = "```\n[[page.md|Link]] inside code\n[text](http://example.com)\n```\nOutside [[page.md|Real]] link.\n"
        result = process_markdown_links(
            text,
            src,
            link_config=replace(
                _ORCHESTRATOR_POLICY,
                non_embed_wiki_mode="extract_text",
                non_embed_markdown_mode="extract_text",
            ),
        )
        # Inside code block — untouched
        assert "[[page.md|Link]]" in result
        assert "[text](http://example.com)" in result
        # Outside code block — processed
        assert "Real" in result
        assert "[[page.md|Real]]" not in result

    def test_inline_code_spans_protected(self, tmp_path: Path) -> None:
        """Markdown links inside inline code spans are not processed."""
        src = _write(tmp_path / "doc.md", "")
        text = "`[text](url)` is a code span. [real](http://a.com) is not."
        result = process_markdown_links(
            text,
            src,
            link_config=replace(
                _ORCHESTRATOR_POLICY,
                non_embed_markdown_mode="extract_text",
            ),
        )
        assert "`[text](url)`" in result
        assert "real" in result
        assert "[real](http://a.com)" not in result

    def test_wiki_link_target_only_no_display(self, tmp_path: Path) -> None:
        """``[[target]]`` without display text."""
        src = _write(tmp_path / "doc.md", "")
        text = "See [[page.md]] for more."
        result = process_markdown_links(
            text,
            src,
            link_config=replace(
                _ORCHESTRATOR_POLICY,
                non_embed_wiki_mode="extract_text",
            ),
        )
        assert "page.md" in result
        assert "[[page.md]]" not in result

    def test_embed_not_confused_with_non_embed(self, tmp_path: Path) -> None:
        """``![[embed.md]]`` is embedded, ``[[link.md]]`` is non-embedded."""
        src = _write(tmp_path / "doc.md", "")
        _write(tmp_path / "embed.md", "# Embed\n\nEmbedded content.\n")
        text = "![[embed.md]] and [[link.md|Link]]"
        result = process_markdown_links(
            text,
            src,
            link_config=replace(
                _ORCHESTRATOR_POLICY,
                non_embed_wiki_mode="extract_text",
            ),
        )
        assert "Embedded content." in result
        assert "![[embed.md]]" not in result
        assert "Link" in result
        assert "[[link.md|Link]]" not in result

    def test_max_depth_reached(self, tmp_path: Path) -> None:
        """When max_depth is 0, embeds are not expanded."""
        src = _write(tmp_path / "doc.md", "")
        _write(tmp_path / "deep.md", "# Deep\n\nContent.\n")
        text = "![[deep.md]]"
        result = process_markdown_links(
            text,
            src,
            link_config=replace(_ORCHESTRATOR_POLICY, max_depth=0),
        )
        assert "Max depth reached" in result
        assert "Content." not in result


# ═══════════════════════════════════════════════════════════════════════════
# _process_non_embed_links  (F-H2-020)
# ═══════════════════════════════════════════════════════════════════════════


class TestProcessNonEmbedLinks:
    """Direct tests for the ``_process_non_embed_links`` function."""

    def test_wiki_extract_text(self, tmp_path: Path) -> None:
        src = _write(tmp_path / "doc.md", "")
        text = "See [[page.md|My Page]] and [[other.md]]."
        result = _process_non_embed_links(text, source_file_path=src, wiki_mode="extract_text")
        assert "My Page" in result
        assert "other.md" in result
        assert "[[page.md|My Page]]" not in result
        assert "[[other.md]]" not in result

    def test_wiki_remove(self, tmp_path: Path) -> None:
        src = _write(tmp_path / "doc.md", "")
        text = "See [[page.md|My Page]] for details."
        result = _process_non_embed_links(text, source_file_path=src, wiki_mode="remove")
        assert "[[page.md|My Page]]" not in result
        assert "My Page" not in result

    def test_wiki_keep(self, tmp_path: Path) -> None:
        src = _write(tmp_path / "doc.md", "")
        text = "See [[page.md|My Page]] for details."
        result = _process_non_embed_links(text, source_file_path=src, wiki_mode="keep")
        assert result == text

    def test_wiki_resolve(self, tmp_path: Path) -> None:
        src = _write(tmp_path / "doc.md", "")
        _write(tmp_path / "page.md", "# Page\n")
        text = "See [[page.md|My Page]] for details."
        result = _process_non_embed_links(text, source_file_path=src, wiki_mode="resolve")
        assert "[My Page](<page.md>)" in result
        assert "[[page.md|My Page]]" not in result

    def test_markdown_extract_text(self, tmp_path: Path) -> None:
        src = _write(tmp_path / "doc.md", "")
        text = "Visit [example](https://example.com) today."
        result = _process_non_embed_links(text, source_file_path=src, markdown_mode="extract_text")
        assert "example" in result
        assert "[example](https://example.com)" not in result

    def test_markdown_remove(self, tmp_path: Path) -> None:
        src = _write(tmp_path / "doc.md", "")
        text = "Visit [example](https://example.com) today."
        result = _process_non_embed_links(text, source_file_path=src, markdown_mode="remove")
        assert "[example](https://example.com)" not in result
        assert "today." in result

    def test_markdown_keep(self, tmp_path: Path) -> None:
        src = _write(tmp_path / "doc.md", "")
        text = "Visit [example](https://example.com) today."
        result = _process_non_embed_links(text, source_file_path=src, markdown_mode="keep")
        assert result == text

    def test_both_modes_independent(self, tmp_path: Path) -> None:
        """Wiki and markdown modes operate independently."""
        src = _write(tmp_path / "doc.md", "")
        text = "Wiki [[page.md|Page]] and MD [link](http://a.com)."
        result = _process_non_embed_links(
            text,
            source_file_path=src,
            wiki_mode="extract_text",
            markdown_mode="remove",
        )
        assert "Page" in result
        assert "[[page.md|Page]]" not in result
        assert "[link](http://a.com)" not in result

    def test_code_blocks_protected(self, tmp_path: Path) -> None:
        src = _write(tmp_path / "doc.md", "")
        text = "```python\n[[not_a_link]] and [text](url)\n```\nReal [[link.md|Link]].\n"
        result = _process_non_embed_links(text, source_file_path=src, wiki_mode="extract_text", markdown_mode="remove")
        # Code block content untouched
        assert "[[not_a_link]]" in result
        assert "[text](url)" in result
        # Outside processed
        assert "Link" in result
        assert "[[link.md|Link]]" not in result

    def test_inline_code_spans_protected(self, tmp_path: Path) -> None:
        src = _write(tmp_path / "doc.md", "")
        text = "`[code](link)` is safe but [real](link) is not."
        result = _process_non_embed_links(text, source_file_path=src, markdown_mode="extract_text")
        assert "`[code](link)`" in result  # code span untouched
        assert "real" in result
        assert "[real](link)" not in result

    def test_wiki_with_escaped_pipe(self, tmp_path: Path) -> None:
        src = _write(tmp_path / "doc.md", "")
        text = r"See [[page\|name.md|Display]] for details."
        result = _process_non_embed_links(text, source_file_path=src, wiki_mode="extract_text")
        assert "Display" in result

    def test_empty_text(self, tmp_path: Path) -> None:
        src = _write(tmp_path / "doc.md", "")
        assert _process_non_embed_links("", source_file_path=src) == ""

    def test_text_without_links(self, tmp_path: Path) -> None:
        src = _write(tmp_path / "doc.md", "")
        text = "# Title\n\nJust paragraph text.\n"
        result = _process_non_embed_links(text, source_file_path=src)
        assert result == text

    def test_wiki_link_to_heading_only(self, tmp_path: Path) -> None:
        """Wiki link targeting only a heading (no file path)."""
        src = _write(tmp_path / "doc.md", "")
        text = "Jump to [[#section-name|Section]]."
        result = _process_non_embed_links(text, source_file_path=src, wiki_mode="extract_text")
        assert "Section" in result
        assert "[[#section-name|Section]]" not in result


# ═══════════════════════════════════════════════════════════════════════════
# User-path integration: orchestrator → real converter scenario
# ═══════════════════════════════════════════════════════════════════════════


class TestOrchestratorUserPath:
    """Simulate real converter usage of the orchestrator with mixed content."""

    def test_pdf_extracted_markdown_scenario(self, tmp_path: Path) -> None:
        """Simulate the layout converter's post-extraction workflow.

        After pymupdf4llm extracts Markdown from a PDF, the converter calls
        ``process_markdown_links`` to resolve any wiki-embed links and handle
        non-embed cross-references.
        """
        src = _write(tmp_path / "output.md", "")

        # Simulated extracted Markdown (from PDF)
        extracted = (
            "# Research Report\n\n"
            "## Introduction\n\n"
            "Please refer to the ![[appendix.md#Methodology]] for details.\n\n"
            "See also [[references.md|References]] and "
            "[online docs](https://example.com/docs).\n\n"
            "![[figure1.png|400x300]]\n"
        )

        # Supporting files exist on disk
        _write(
            tmp_path / "appendix.md",
            (
                "---\ntitle: Appendix\n---\n\n"
                "# Appendix\n\n"
                "Extra material.\n\n"
                "## Methodology\n\n"
                "The study used a mixed-methods approach.\n\n"
                "## Results\n\n"
                "Data were significant.\n"
            ),
        )
        _write(tmp_path / "references.md", "# References\n\n1. Doe, J.\n")
        # Create a dummy image file so the resolver can find it
        _write(tmp_path / "figure1.png", "")

        result = process_markdown_links(
            extracted,
            src,
            link_config=replace(
                _ORCHESTRATOR_POLICY,
                non_embed_wiki_mode="resolve",
                embed_wiki_image_mode="keep",
            ),
        )

        # Embed was resolved with section extraction
        assert "mixed-methods approach" in result
        assert "![[appendix.md#Methodology]]" not in result
        # YAML front matter was stripped from embedded appendix content
        assert "The study used" in result
        assert "Extra material." not in result  # outside the Methodology section

        # Non-embed wiki link was resolved
        assert "[References](<references.md>)" in result

        # Markdown link was kept
        assert "[online docs](https://example.com/docs)" in result

        # Image embed was kept (image_mode="keep")
        assert "![[figure1.png|400x300]]" in result

    def test_no_files_no_crash(self, tmp_path: Path) -> None:
        """When no referenced files exist on disk, the orchestrator does not
        crash and produces placeholders (or keeps links depending on mode)."""
        src = _write(tmp_path / "doc.md", "")
        text = "![[missing_embed.md]]\n[[missing_link.md|Missing]]\n[example](https://example.com)\n"
        result = process_markdown_links(
            text,
            src,
            link_config=replace(
                _ORCHESTRATOR_POLICY,
                non_embed_wiki_mode="extract_text",
            ),
        )
        # Embed missing → placeholder
        assert "File not found" in result
        # Wiki non-embed extract_text works regardless of file existence
        assert "Missing" in result
        # Markdown link kept
        assert "[example](https://example.com)" in result

    def test_circular_reference_detected(self, tmp_path: Path) -> None:
        """Circular embed references are detected and handled."""
        src = _write(tmp_path / "a.md", "![[b.md]]\n")
        _write(tmp_path / "b.md", "![[a.md]]\n")

        result = process_markdown_links(
            "![[a.md]]",
            src,
            link_config=_ORCHESTRATOR_POLICY,
        )
        assert "Circular reference" in result

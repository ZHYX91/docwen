"""Tests for docwen_core.links — anchor resolution, section/block extraction,
front matter stripping, and file path resolution.

Covers F-H2-002, F-H2-003, F-H2-004, F-H2-005, F-H2-015.
"""

from __future__ import annotations

import os
import tempfile
from collections.abc import Generator
from pathlib import Path

import pytest

from docwen_core.links import (
    extract_block_by_id,
    extract_section_by_heading,
    normalize_link_target,
    parse_anchor,
    resolve_file_path,
    strip_yaml_front_matter,
)

pytestmark = pytest.mark.unit


# ═══════════════════════════════════════════════════════════════════════
# parse_anchor  (F-H2-002)
# ═══════════════════════════════════════════════════════════════════════


class TestParseAnchor:
    """Covers F-H2-002: parse_anchor splitting into (file_path, heading, block_id)."""

    # ── basic forms ──────────────────────────────────────────────────

    def test_simple_file_no_anchor(self) -> None:
        assert parse_anchor("file.md") == ("file.md", None, None)

    def test_file_with_heading_anchor(self) -> None:
        fp, heading, block = parse_anchor("file.md#Introduction")
        assert fp == "file.md"
        assert heading == "Introduction"
        assert block is None

    def test_file_with_block_id_anchor(self) -> None:
        fp, heading, block = parse_anchor("file.md#^abc123")
        assert fp == "file.md"
        assert heading is None
        assert block == "abc123"

    def test_only_anchor_hash(self) -> None:
        fp, heading, block = parse_anchor("#heading-only")
        assert fp is None
        assert heading == "heading-only"
        assert block is None

    def test_only_block_anchor(self) -> None:
        fp, heading, block = parse_anchor("#^block-only")
        assert fp is None
        assert heading is None
        assert block == "block-only"

    # ── URL-decoding ─────────────────────────────────────────────────

    def test_url_encoded_path(self) -> None:
        fp, heading, block = parse_anchor("my%20file.md#My%20Heading")
        assert fp == "my file.md"
        assert heading == "My Heading"
        assert block is None

    def test_url_encoded_international(self) -> None:
        fp, heading, block = parse_anchor("%E6%96%87%E4%BB%B6.md#%E6%A0%87%E9%A2%98")
        assert fp == "文件.md"
        assert heading == "标题"
        assert block is None

    # ── query string stripping ───────────────────────────────────────

    def test_query_string_stripped(self) -> None:
        fp, heading, block = parse_anchor("file.md?v=2#heading")
        assert fp == "file.md"
        assert heading == "heading"
        assert block is None

    def test_query_string_only(self) -> None:
        fp, heading, block = parse_anchor("file.md?v=2&x=3")
        assert fp == "file.md"
        assert heading is None
        assert block is None

    # ── edge cases ───────────────────────────────────────────────────

    def test_empty_string(self) -> None:
        fp, heading, block = parse_anchor("")
        assert fp is None
        assert heading is None
        assert block is None

    def test_whitespace_only(self) -> None:
        fp, heading, block = parse_anchor("   ")
        assert fp is None
        assert heading is None
        assert block is None

    def test_hash_with_empty_anchor(self) -> None:
        fp, heading, block = parse_anchor("file.md#")
        assert fp == "file.md"
        assert heading is None
        assert block is None

    def test_hash_with_spaces_anchor(self) -> None:
        fp, heading, block = parse_anchor("file.md#   ")
        assert fp == "file.md"
        assert heading is None
        assert block is None

    def test_subdir_path(self) -> None:
        fp, heading, block = parse_anchor("sub/dir/file.md#heading")
        assert fp == "sub/dir/file.md"
        assert heading == "heading"
        assert block is None

    def test_windows_path(self) -> None:
        fp, heading, _block = parse_anchor(r"sub\dir\file.md#heading")
        assert fp == r"sub\dir\file.md"
        assert heading == "heading"

    def test_heading_with_special_chars(self) -> None:
        fp, heading, block = parse_anchor("file.md#FAQ / Q&A")
        assert fp == "file.md"
        assert heading == "FAQ / Q&A"
        assert block is None

    def test_block_id_with_hyphen(self) -> None:
        fp, heading, block = parse_anchor("file.md#^my-block-id")
        assert fp == "file.md"
        assert heading is None
        assert block == "my-block-id"


# ═══════════════════════════════════════════════════════════════════════
# extract_section_by_heading  (F-H2-003)
# ═══════════════════════════════════════════════════════════════════════

SAMPLE_DOC = """\
# Top Level

Some intro text.

## Section One

Content of section one.

### Subsection A

Deeper content.

## Section Two

Content of section two.

Another paragraph.
"""


class TestExtractSectionByHeading:
    """Covers F-H2-003: section extraction by heading."""

    def test_extract_h2_section(self) -> None:
        result = extract_section_by_heading(SAMPLE_DOC, "Section One")
        assert result is not None
        assert "## Section One" in result
        assert "Content of section one" in result
        assert "### Subsection A" in result
        assert "## Section Two" not in result

    def test_extract_h2_stops_at_same_level(self) -> None:
        result = extract_section_by_heading(SAMPLE_DOC, "Section Two")
        assert result is not None
        assert "## Section Two" in result
        assert "Content of section two" in result
        assert "Another paragraph" in result
        # Should not contain content from Section One
        assert "Section One" not in result

    def test_extract_h1(self) -> None:
        result = extract_section_by_heading(SAMPLE_DOC, "Top Level")
        assert result is not None
        assert "# Top Level" in result
        assert "## Section One" in result
        assert "## Section Two" in result

    def test_extract_h3_subsection(self) -> None:
        result = extract_section_by_heading(SAMPLE_DOC, "Subsection A")
        assert result is not None
        assert "### Subsection A" in result
        assert "Deeper content" in result
        # Stops at next h2 (higher level)
        assert "## Section Two" not in result

    def test_heading_not_found(self) -> None:
        result = extract_section_by_heading(SAMPLE_DOC, "Nonexistent")
        assert result is None

    def test_case_insensitive_match(self) -> None:
        result = extract_section_by_heading(SAMPLE_DOC, "section one")
        assert result is not None
        assert "## Section One" in result

    def test_whitespace_normalized_match(self) -> None:
        result = extract_section_by_heading(SAMPLE_DOC, "  Section   One  ")
        assert result is not None
        assert "## Section One" in result

    def test_empty_content(self) -> None:
        result = extract_section_by_heading("", "Anything")
        assert result is None

    def test_no_headings_at_all(self) -> None:
        result = extract_section_by_heading("Just plain text\nno headings", "Anything")
        assert result is None

    def test_single_heading_no_following_headings(self) -> None:
        content = "# Only Heading\n\nSome content\nmore content"
        result = extract_section_by_heading(content, "Only Heading")
        assert result is not None
        assert "Some content" in result
        assert "more content" in result

    def test_trailing_blank_lines_trimmed(self) -> None:
        content = "# H1\n\nBody line\n\n\n"
        result = extract_section_by_heading(content, "H1")
        assert result is not None
        assert not result.endswith("\n\n")

    def test_stop_at_higher_level_heading(self) -> None:
        content = "### Sub\n\nSub content\n\n# Top\n\nTop content"
        result = extract_section_by_heading(content, "Sub")
        assert result is not None
        assert "# Top" not in result


# ═══════════════════════════════════════════════════════════════════════
# extract_block_by_id  (F-H2-004)
# ═══════════════════════════════════════════════════════════════════════

BLOCK_DOC = """\
# Heading

This is a paragraph with an inline marker. ^myid

Another paragraph here.

## Second Heading

Some text before the block.

The block contents are here.

^block1

After the block marker.

Yet another paragraph with ^inline2

# Next Section
"""


class TestExtractBlockById:
    """Covers F-H2-004: block extraction by ^block-id."""

    # ── inline markers ───────────────────────────────────────────────

    def test_inline_marker_basic(self) -> None:
        result = extract_block_by_id(BLOCK_DOC, "myid")
        assert result is not None
        assert "This is a paragraph with an inline marker." in result
        assert "^myid" not in result  # marker stripped

    def test_inline_marker_collects_adjacent(self) -> None:
        result = extract_block_by_id(BLOCK_DOC, "inline2")
        assert result is not None
        assert "Yet another paragraph with" in result

    # ── standalone markers ───────────────────────────────────────────

    def test_standalone_marker(self) -> None:
        result = extract_block_by_id(BLOCK_DOC, "block1")
        assert result is not None
        assert "The block contents are here." in result
        assert "^block1" not in result  # marker not in output

    # ── edge cases ───────────────────────────────────────────────────

    def test_block_id_not_found(self) -> None:
        result = extract_block_by_id(BLOCK_DOC, "nonexistent")
        assert result is None

    def test_empty_content(self) -> None:
        result = extract_block_by_id("", "anything")
        assert result is None

    def test_marker_only_content(self) -> None:
        """Standalone marker at top of document with preceding blank lines."""
        content = "\n\n^topmarker\n\n# Heading"
        result = extract_block_by_id(content, "topmarker")
        # Should return None — no preceding paragraph
        assert result is None

    def test_block_id_with_hyphen(self) -> None:
        content = "Paragraph with hyphen id. ^my-block-id\n\n# Next"
        result = extract_block_by_id(content, "my-block-id")
        assert result is not None
        assert "Paragraph with hyphen id" in result

    def test_block_id_with_underscore(self) -> None:
        content = "Paragraph with underscore id. ^my_block_id\n\n# Next"
        result = extract_block_by_id(content, "my_block_id")
        assert result is not None
        assert "Paragraph with underscore id" in result

    def test_standalone_no_preceding_paragraph(self) -> None:
        """Standalone marker with only blank lines above — no paragraph to extract."""
        content = "\n\n\n^lonely\n\nText after"
        result = extract_block_by_id(content, "lonely")
        assert result is None

    def test_standalone_with_preceding_paragraph(self) -> None:
        content = "Some paragraph text.\n\n^marker1\n\n# Heading"
        result = extract_block_by_id(content, "marker1")
        assert result is not None
        assert "Some paragraph text." in result

    def test_inline_stops_at_heading_above(self) -> None:
        content = "# Heading\n\nparagraph with ^marker"
        result = extract_block_by_id(content, "marker")
        assert result is not None
        assert "# Heading" not in result

    def test_inline_stops_at_next_block_marker(self) -> None:
        content = "first part ^markerA\nnext line\nline with ^markerB\n"
        result = extract_block_by_id(content, "markerA")
        assert result is not None
        # Should stop at the line with ^markerB (another block id)
        assert "markerB" not in result


# ═══════════════════════════════════════════════════════════════════════
# strip_yaml_front_matter  (F-H2-005)
# ═══════════════════════════════════════════════════════════════════════


class TestStripYamlFrontMatter:
    """Covers F-H2-005: YAML front matter stripping."""

    def test_basic_front_matter(self) -> None:
        content = "---\ntitle: Test\ndate: 2024-01-01\n---\n\n# Heading\nBody"
        result = strip_yaml_front_matter(content)
        assert result.startswith("# Heading")
        assert "title:" not in result

    def test_no_front_matter(self) -> None:
        content = "# Just a heading\nSome content"
        result = strip_yaml_front_matter(content)
        assert result == content

    def test_empty_content(self) -> None:
        assert strip_yaml_front_matter("") == ""

    def test_only_opening_dashes(self) -> None:
        content = "---\nNo closing dashes\n# Heading"
        result = strip_yaml_front_matter(content)
        assert result == content  # unchanged — no closing ---

    def test_dashes_in_body_not_stripped(self) -> None:
        content = "# Heading\n\nSome text with --- in it\n\nMore text"
        result = strip_yaml_front_matter(content)
        assert result == content  # first line not ---

    def test_front_matter_with_crlf(self) -> None:
        content = "---\r\ntitle: Test\r\n---\r\n\r\n# Heading"
        result = strip_yaml_front_matter(content)
        assert result.startswith("# Heading")

    def test_front_matter_with_utf8_bom(self) -> None:
        content = "\ufeff---\r\ntitle: Test\r\n---\r\n\r\n# Heading"
        result = strip_yaml_front_matter(content)
        assert result == "# Heading"

    def test_inline_triple_hyphen_in_yaml_value_is_not_a_close_delimiter(self) -> None:
        content = '---\ntitle: "a---b"\n---\nBody'
        result = strip_yaml_front_matter(content)
        assert result == "Body"

    def test_front_matter_multiline_yaml(self) -> None:
        content = "---\ntitle: Complex\ntags:\n  - a\n  - b\n---\n\nContent"
        result = strip_yaml_front_matter(content)
        assert result == "Content"

    def test_minimal_front_matter(self) -> None:
        content = "------\nContent"
        result = strip_yaml_front_matter(content)
        assert result == content

    def test_front_matter_only_content(self) -> None:
        """Content that is only front matter."""
        content = "---\ntitle: Only\n---"
        result = strip_yaml_front_matter(content)
        assert result == ""


# ═══════════════════════════════════════════════════════════════════════
# normalize_link_target
# ═══════════════════════════════════════════════════════════════════════


class TestNormalizeLinkTarget:
    def test_basic_path(self) -> None:
        assert normalize_link_target("file.md") == "file.md"

    def test_strips_anchor(self) -> None:
        assert normalize_link_target("file.md#heading") == "file.md"

    def test_strips_query(self) -> None:
        assert normalize_link_target("file.md?v=1") == "file.md"

    def test_strips_both(self) -> None:
        assert normalize_link_target("file.md?v=1#heading") == "file.md"

    def test_url_decoded(self) -> None:
        assert normalize_link_target("my%20file.md") == "my file.md"

    @pytest.mark.parametrize(
        ("target", "expected"),
        [
            ("my%23file.png", "my#file.png"),
            ("my%3Ffile.png", "my?file.png"),
        ],
    )
    def test_encoded_filename_delimiter_is_not_structural(
        self,
        target: str,
        expected: str,
    ) -> None:
        assert normalize_link_target(target) == expected

    def test_whitespace_trimmed(self) -> None:
        assert normalize_link_target("  file.md  ") == "file.md"

    def test_subdir_path(self) -> None:
        assert normalize_link_target("sub/dir/file.md#anchor") == "sub/dir/file.md"

    def test_unc_authority_is_not_downgraded_to_a_relative_path(self) -> None:
        assert (
            normalize_link_target(
                r"\\server\share\file.png",
                preserve_absolute=True,
            )
            == "//server/share/file.png"
        )

    def test_posix_absolute_root_is_not_downgraded_to_a_relative_path(self) -> None:
        assert (
            normalize_link_target(
                "/tmp/file.png",
                preserve_absolute=True,
            )
            == "/tmp/file.png"
        )

    def test_default_contract_keeps_markdown_output_href_relative(self) -> None:
        assert normalize_link_target("/images/photo.png") == "images/photo.png"


# ═══════════════════════════════════════════════════════════════════════
# resolve_file_path  (F-H2-015)
# ═══════════════════════════════════════════════════════════════════════


class TestResolveFilePath:
    """Covers F-H2-015: file path resolution with priority tiers."""

    @pytest.fixture
    def temp_workspace(self) -> Generator[str, None, None]:
        """Create a temporary directory tree with known files."""
        tmp = tempfile.mkdtemp()
        base = Path(tmp)

        # Create source file
        source = base / "notes" / "mynote.md"
        source.parent.mkdir(parents=True)
        source.write_text("# My Note", encoding="utf-8")

        # Create a same-name folder with an asset
        asset_dir = base / "notes" / "mynote"
        asset_dir.mkdir(parents=True)
        (asset_dir / "diagram.png").write_text("fake-png", encoding="utf-8")

        # Create assets dir
        assets = base / "notes" / "assets"
        assets.mkdir(parents=True)
        (assets / "photo.jpg").write_text("fake-jpg", encoding="utf-8")
        (assets / "photo with spaces.jpg").write_text("fake-jpg", encoding="utf-8")

        # A same-basename file outside the bounded source-local search area.
        distant = base / "another" / "deep"
        distant.mkdir(parents=True)
        (distant / "distant image.png").write_text("fake-png", encoding="utf-8")

        # Create an md file in same dir
        (source.parent / "linked.md").write_text("# Linked", encoding="utf-8")

        # Create file with extension-less name + .md auto-add
        (source.parent / "other.md").write_text("# Other", encoding="utf-8")

        yield str(source)

        # Cleanup
        import shutil

        shutil.rmtree(tmp, ignore_errors=True)

    def test_relative_path_with_separator(self, temp_workspace: str) -> None:
        result = resolve_file_path("linked.md", temp_workspace)
        assert result is not None
        assert result.endswith("linked.md")

    def test_relative_subdir_path(self, temp_workspace: str) -> None:
        result = resolve_file_path("assets/photo.jpg", temp_workspace)
        assert result is not None
        assert result.endswith("photo.jpg")

    def test_same_name_folder(self, temp_workspace: str) -> None:
        # mynote.md → mynote/ directory exists
        result = resolve_file_path("diagram.png", temp_workspace)
        assert result is not None
        assert "diagram.png" in result

    def test_auto_md_extension(self, temp_workspace: str) -> None:
        # "other" without extension should resolve to "other.md"
        result = resolve_file_path("other", temp_workspace)
        assert result is not None
        assert result.endswith("other.md")

    def test_not_found(self, temp_workspace: str) -> None:
        result = resolve_file_path("nonexistent.file", temp_workspace)
        assert result is None

    def test_absolute_path(self, temp_workspace: str) -> None:
        # Create a file and resolve via absolute path
        import tempfile

        with tempfile.NamedTemporaryFile(suffix=".md", delete=False) as f:
            f.write(b"# Temp")
            abs_path = f.name.replace("\\", "/")

        try:
            result = resolve_file_path(abs_path, temp_workspace)
            assert result is not None
            # On Windows normpath may differ from input; we just check it resolved
            assert result is not None
        finally:
            os.unlink(abs_path)

    def test_strips_anchor_from_target(self, temp_workspace: str) -> None:
        """Anchor fragments should be stripped before resolution."""
        result = resolve_file_path("linked.md#heading", temp_workspace)
        assert result is not None
        assert result.endswith("linked.md")

    def test_url_encoded_target(self, temp_workspace: str) -> None:
        result = resolve_file_path("linked%2Emd", temp_workspace)
        assert result is not None
        assert result.endswith("linked.md")

    def test_custom_search_dirs(self, temp_workspace: str) -> None:
        # photo.jpg is in assets/ dir — custom search dirs should find it
        result = resolve_file_path("photo.jpg", temp_workspace, search_dirs=["assets"])
        assert result is not None
        assert result.endswith("photo.jpg")

    def test_default_search_dirs(self, temp_workspace: str) -> None:
        # With default search_dirs (includes "assets"), photo.jpg should be found
        result = resolve_file_path("photo.jpg", temp_workspace)
        assert result is not None
        assert result.endswith("photo.jpg")

    def test_spaces_are_supported_in_a_conventional_directory(self, temp_workspace: str) -> None:
        result = resolve_file_path("photo with spaces.jpg", temp_workspace)
        assert result is not None
        assert result.endswith("photo with spaces.jpg")

    def test_short_name_does_not_trigger_recursive_or_workspace_search(self, temp_workspace: str) -> None:
        assert resolve_file_path("distant image.png", temp_workspace) is None

    def test_search_directory_cannot_escape_the_source_directory(self, temp_workspace: str) -> None:
        assert (
            resolve_file_path(
                "distant image.png",
                temp_workspace,
                search_dirs=["../../another/deep"],
            )
            is None
        )


# ═══════════════════════════════════════════════════════════════════════
# Integration: user-path wiring  (H2 user-path evidence)
# ═══════════════════════════════════════════════════════════════════════


class TestUserPathIntegration:
    """Demonstrate the full user path: parse → extract section → strip front matter."""

    def test_wiki_link_to_section(self) -> None:
        """Simulate resolving a wiki link with heading anchor."""
        content = """\
---
title: Source Doc
---
# Introduction

Welcome text.

## Features

- Feature A
- Feature B

## Limitations

- Limitation 1
"""
        # Strip front matter first
        cleaned = strip_yaml_front_matter(content)
        assert not cleaned.startswith("---")

        # Parse anchor like "doc.md#Features"
        _fp, heading, block = parse_anchor("doc.md#Features")
        assert heading == "Features"
        assert block is None

        # Extract the section
        section = extract_section_by_heading(cleaned, heading)
        assert section is not None
        assert "## Features" in section
        assert "- Feature A" in section
        assert "## Limitations" not in section

    def test_wiki_link_to_block(self) -> None:
        """Simulate resolving a wiki link with block-id anchor."""
        content = """\
# Notes

First paragraph with key insight. ^key1

Second paragraph.

Another important note here. ^key2

# Appendix
"""
        _, __, block_id = parse_anchor("doc.md#^key1")
        assert block_id == "key1"

        block = extract_block_by_id(content, block_id)
        assert block is not None
        assert "First paragraph with key insight" in block
        assert "^key1" not in block

    def test_full_pipeline_with_file(self) -> None:
        """End-to-end: parse anchor, strip front matter, extract section."""
        doc = """\
---
author: test
---
# Overview

General overview.

## Details

Specific details here.

## Summary

Summary content.
"""
        fp, heading, _block = parse_anchor("notes/doc.md#Details")
        assert fp == "notes/doc.md"
        assert heading == "Details"

        cleaned = strip_yaml_front_matter(doc)
        section = extract_section_by_heading(cleaned, heading)
        assert section is not None
        assert "## Details" in section
        assert "Specific details here" in section

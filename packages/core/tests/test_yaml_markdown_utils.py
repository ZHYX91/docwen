"""Tests for ``docwen_core.yaml_tools`` and ``docwen_core.markdown_utils``.

Covers:
- F-I2b-001: ``extract_yaml`` — YAML front matter extraction
- F-I2b-001: ``generate_basic_yaml_frontmatter`` — YAML front matter generation
- F-I2b-004: ``format_image_link`` — image link formatting with sanitization
- F-I2b-004: ``sanitize_filename`` / ``sanitize_for_wiki_link`` — filename helpers
"""

from __future__ import annotations

import pytest
import yaml

from docwen_core.markdown_utils import (
    format_sanitized_image_link,
    sanitize_filename,
    sanitize_for_wiki_link,
)
from docwen_core.yaml_tools import (
    extract_yaml,
    generate_basic_yaml_frontmatter,
)

pytestmark = pytest.mark.unit

# ═══════════════════════════════════════════════════════════════════════════
# F-I2b-001: extract_yaml
# ═══════════════════════════════════════════════════════════════════════════


class TestExtractYaml:
    """YAML front matter extraction from Markdown content."""

    def test_basic_extraction(self) -> None:
        content = "---\ntitle: Hello\naliases:\n  - hello\n---\n\n# Section\n\nBody text.\n"
        yaml_str, body = extract_yaml(content)
        assert yaml_str == "title: Hello\naliases:\n  - hello"
        assert body == "# Section\n\nBody text."

    def test_no_yaml_front_matter(self) -> None:
        content = "# Just a heading\n\nSome text."
        yaml_str, body = extract_yaml(content)
        assert yaml_str == ""
        assert body == content

    def test_empty_content(self) -> None:
        yaml_str, body = extract_yaml("")
        assert yaml_str == ""
        assert body == ""

    def test_only_yaml_no_body(self) -> None:
        content = "---\ntitle: Only\n---\n"
        yaml_str, body = extract_yaml(content)
        assert yaml_str == "title: Only"
        assert body == ""

    def test_yaml_with_crlf(self) -> None:
        content = "---\r\ntitle: CRLF\r\ntags: [a, b]\r\n---\r\n\r\nBody\r\n"
        yaml_str, body = extract_yaml(content)
        assert yaml_str == "title: CRLF\ntags: [a, b]"
        assert body == "Body"

    def test_yaml_with_bom(self) -> None:
        content = "﻿---\ntitle: BOM\n---\n\nBody with BOM.\n"
        yaml_str, body = extract_yaml(content)
        assert yaml_str == "title: BOM"
        assert body == "Body with BOM."

    def test_yaml_with_complex_fields(self) -> None:
        content = "---\ntitle: 测试文档\ntags:\n  - tag1\n  - tag2\ndate: 2024-01-15\n---\n\n## 章节一\n\n内容。\n"
        yaml_str, body = extract_yaml(content)
        assert "title: 测试文档" in yaml_str
        assert "tags:" in yaml_str
        assert body.startswith("## 章节一")

    def test_roundtrip_generate_extract(self) -> None:
        """User-path: generate YAML front matter, then extract it back."""
        yaml_block = generate_basic_yaml_frontmatter("MyDocument")
        content = yaml_block + "# Hello\n\nWorld.\n"
        yaml_str, body = extract_yaml(content)
        assert "title: MyDocument" in yaml_str
        assert "aliases:" in yaml_str
        assert body == "# Hello\n\nWorld."

    def test_thematic_break_not_confused_as_yaml(self) -> None:
        """A ``---`` later in the document should not be mistaken for YAML."""
        content = "# Heading\n\n---\n\nMore text.\n"
        yaml_str, body = extract_yaml(content)
        # The first line is "# Heading", not "---", so no YAML extracted
        assert yaml_str == ""
        assert "# Heading" in body


# ═══════════════════════════════════════════════════════════════════════════
# F-I2b-001: generate_basic_yaml_frontmatter
# ═══════════════════════════════════════════════════════════════════════════


class TestGenerateBasicYamlFrontmatter:
    """YAML front matter generation."""

    def test_basic_generation(self) -> None:
        result = generate_basic_yaml_frontmatter("TestFile")
        assert result.startswith("---\n")
        assert "aliases:\n  - TestFile\n" in result
        assert "title: TestFile\n" in result
        assert result.endswith("---\n\n")

    def test_chinese_filename(self) -> None:
        result = generate_basic_yaml_frontmatter("年度报告")
        assert "title: 年度报告" in result
        assert "  - 年度报告" in result

    def test_with_extra_fields(self) -> None:
        result = generate_basic_yaml_frontmatter("image", extra={"source_format": "png"})
        assert "title: image" in result
        assert "source_format: png" in result

    def test_uses_resolved_title_label(self) -> None:
        result = generate_basic_yaml_frontmatter(
            "Workbook",
            yaml_key_labels={"title": "Titel"},
        )
        assert "Titel: Workbook" in result
        assert "title: Workbook" not in result
        assert "aliases:" in result

    def test_with_multiple_extra_fields(self) -> None:
        result = generate_basic_yaml_frontmatter(
            "doc",
            extra={"author": "ZYX", "date": "2024-01-01"},
        )
        yaml_str, _body = extract_yaml(result)
        assert yaml.safe_load(yaml_str) == {
            "title": "doc",
            "aliases": ["doc"],
            "author": "ZYX",
            "date": "2024-01-01",
        }

    def test_string_scalars_roundtrip_without_yaml_type_or_syntax_drift(self) -> None:
        values = (
            "普通文档",
            "[2024]年报",
            "a: b",
            "#tag",
            "true",
            "null",
            "01",
            "O'Brien",
            "2024-01-01",
            "yes",
            " leading",
            "trailing ",
            "line\nbreak",
            "delete\u007fseparator",
            "nel\u0085separator",
            "control\u009fseparator",
            "line\u2028separator",
            "paragraph\u2029separator",
            "noncharacter\ufffeseparator",
            "noncharacter\uffffseparator",
            "emoji😀separator",
            "document📄separator",
            "plane-end\U0001fffe separator",
            "unicode-max\U0010ffff separator",
            '他说"你好"',
            "",
        )

        for value in values:
            result = generate_basic_yaml_frontmatter(value, extra={"source": value})
            yaml_str, _body = extract_yaml(result)
            parsed = yaml.safe_load(yaml_str)

            assert parsed == {
                "title": value,
                "aliases": [value],
                "source": value,
            }
            assert isinstance(parsed["title"], str)
            assert isinstance(parsed["aliases"][0], str)
            assert isinstance(parsed["source"], str)

    def test_boolean_extra_field_keeps_its_existing_yaml_type(self) -> None:
        result = generate_basic_yaml_frontmatter("image_ocr", extra={"ocr": True})
        yaml_str, _body = extract_yaml(result)
        parsed = yaml.safe_load(yaml_str)

        assert parsed["ocr"] is True
        assert "ocr: True" in result

    def test_extra_fields_appear_after_title(self) -> None:
        result = generate_basic_yaml_frontmatter("x", extra={"key": "val"})
        title_pos = result.index("title:")
        extra_pos = result.index("key: val")
        assert title_pos < extra_pos, "extra fields should appear after title"

    def test_title_before_aliases_field_order(self) -> None:
        """R1 regression: ``title:`` must appear before ``aliases:`` to
        match the old converter's ad‑hoc YAML output order."""
        result = generate_basic_yaml_frontmatter("DocName")
        title_pos = result.index("title:")
        aliases_pos = result.index("aliases:")
        assert title_pos < aliases_pos, f"title: (pos {title_pos}) must come before aliases: (pos {aliases_pos})"


# ═══════════════════════════════════════════════════════════════════════════
# F-I2b-004: sanitize_filename
# ═══════════════════════════════════════════════════════════════════════════


class TestSanitizeFilename:
    """File-system filename sanitization."""

    def test_illegal_chars_replaced(self) -> None:
        assert sanitize_filename("file:name*test?.txt") == "file_name_test_.txt"

    def test_spaces_compressed(self) -> None:
        assert sanitize_filename("file   name  .txt") == "file name .txt"

    def test_leading_trailing_dots_stripped(self) -> None:
        assert sanitize_filename("...file.name...") == "file.name"

    def test_leading_trailing_spaces_stripped(self) -> None:
        assert sanitize_filename("  file.txt  ") == "file.txt"

    def test_chinese_characters_preserved(self) -> None:
        assert sanitize_filename("报告2024.txt") == "报告2024.txt"

    def test_backslash_replaced(self) -> None:
        assert sanitize_filename("path\\to\\file.txt") == "path_to_file.txt"


class TestSanitizeForWikiLink:
    """Wiki link filename sanitization."""

    def test_wiki_chars_stripped(self) -> None:
        # [ ] are NOT file-system-illegal → sanitize_filename leaves them;
        # sanitize_for_wiki_link then STRIPS them (does not replace).
        assert sanitize_for_wiki_link("file[name]test.txt") == "filenametest.txt"

    def test_pipe_stripped(self) -> None:
        # | IS file-system-illegal → sanitize_filename replaces with _ first
        assert sanitize_for_wiki_link("file|alt.txt") == "file_alt.txt"

    def test_hash_stripped(self) -> None:
        assert sanitize_for_wiki_link("section#1.txt") == "section1.txt"

    def test_caret_stripped(self) -> None:
        assert sanitize_for_wiki_link("v2^final.txt") == "v2final.txt"

    def test_combined_wiki_and_fs_chars(self) -> None:
        # / is FS-illegal → _; [ ] # are wiki-sensitive → stripped
        assert sanitize_for_wiki_link("path/to[file]#1.txt") == "path_tofile1.txt"


# ═══════════════════════════════════════════════════════════════════════════
# F-I2b-004: format_sanitized_image_link
# ═══════════════════════════════════════════════════════════════════════════


class TestFormatSanitizedImageLink:
    """Image link formatting with style-dependent encoding."""

    # ── Wiki styles ──────────────────────────────────────────────────

    def test_wiki_embed(self) -> None:
        result = format_sanitized_image_link("my image.png", style="wiki_embed")
        assert result == "![[my image.png]]"

    def test_wiki_link(self) -> None:
        result = format_sanitized_image_link("my image.png", style="wiki_link")
        assert result == "[[my image.png]]"

    def test_wiki_strips_sensitive_chars(self) -> None:
        """Wiki-sensitive characters [ ] | # ^ are stripped."""
        result = format_sanitized_image_link("img [test] #1.png", style="wiki_embed")
        assert result == "![[img test 1.png]]"

    def test_wiki_strips_pipe(self) -> None:
        # | is FS-illegal → replaced with _ by sanitize_filename,
        # then sanitize_for_wiki_link would strip it (but already replaced).
        result = format_sanitized_image_link("img|alt.png", style="wiki_embed")
        assert result == "![[img_alt.png]]"

    # ── Markdown styles ──────────────────────────────────────────────

    def test_markdown_embed(self) -> None:
        result = format_sanitized_image_link("my image.png", style="markdown_embed")
        # Spaces preserved in alt text but URL-encoded in link target
        assert result == "![my image.png](my%20image.png)"

    def test_markdown_link(self) -> None:
        result = format_sanitized_image_link("my image.png", style="markdown_link")
        assert result == "[my image.png](my%20image.png)"

    def test_markdown_url_encodes_special_chars(self) -> None:
        result = format_sanitized_image_link("报告 2024.png", style="markdown_embed")
        assert "报告 2024.png" in result  # alt text preserved
        assert "%E6%8A%A5%E5%91%8A%202024.png" in result  # URL-encoded

    def test_markdown_replaces_fs_illegal_chars(self) -> None:
        result = format_sanitized_image_link("file:name*.png", style="markdown_embed")
        assert "file_name_.png" in result

    # ── Fallback behaviour ───────────────────────────────────────────

    def test_unknown_style_falls_back_to_wiki_embed(self) -> None:
        result = format_sanitized_image_link("img.png", style="invalid_style")
        assert result == "![[img.png]]"

    def test_default_style_is_wiki_embed(self) -> None:
        result = format_sanitized_image_link("img.png")
        assert result == "![[img.png]]"

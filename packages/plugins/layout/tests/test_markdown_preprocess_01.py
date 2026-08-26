"""Focused tests split from test_markdown_preprocess.py."""

from __future__ import annotations

import pytest

from ._markdown_preprocess_support import (
    Path,
    _write,
    resolve_embedded_links,
)

pytestmark = pytest.mark.unit


class TestEmbedResolutionInMarkdownOutput:
    """Simulate the embed-resolution step that the layout converter performs
    after extracting Markdown from a PDF.

    The converter produces Markdown text; before writing the artifact,
    ``resolve_embedded_links`` scans for ``![[...]]`` patterns and expands
    each one.
    """

    def test_embed_single_md_file(self, tmp_path: Path) -> None:
        """A Markdown document embedding another file is fully resolved."""
        src = _write(tmp_path / "output.md", "")

        # Simulate extracted Markdown from PDF with an embed link
        extracted = "# Extracted PDF\n\nSee also: ![[notes.md]]\n"

        # Target file exists on disk
        _write(tmp_path / "notes.md", "# Notes\n\nImportant details.\n")

        resolved = resolve_embedded_links(extracted, src)
        assert "Important details." in resolved
        assert "![[notes.md]]" not in resolved

    def test_embed_with_heading(self, tmp_path: Path) -> None:
        """Section-level precision embed after PDF extraction."""
        src = _write(tmp_path / "output.md", "")
        extracted = "# Report\n\n![[ref.md#Summary]]\n"
        _write(
            tmp_path / "ref.md",
            ("---\ntitle: Ref\n---\n\n# Ref\n\nIntro.\n\n## Summary\n\nKey takeaway.\n\n## Appendix\n\nExtra.\n"),
        )

        resolved = resolve_embedded_links(extracted, src)
        assert "Key takeaway." in resolved
        assert "Extra." not in resolved
        assert "title:" not in resolved  # YAML stripped

    def test_embed_with_block_id(self, tmp_path: Path) -> None:
        """Block-level precision embed in post-processing."""
        src = _write(tmp_path / "output.md", "")
        extracted = "# Report\n\n![[data.md#^k1]]\n"
        _write(tmp_path / "data.md", ("# Data\n\nFirst note. ^k1\n\nSecond note.\n"))

        resolved = resolve_embedded_links(extracted, src)
        assert "First note." in resolved
        assert "Second note." not in resolved

    def test_nested_embed_resolution(self, tmp_path: Path) -> None:
        """Recursive: A embeds B, B embeds C — all expanded."""
        src = _write(tmp_path / "output.md", "")
        extracted = "# PDF Output\n\n![[outer.md]]\n"
        _write(tmp_path / "outer.md", "# Outer\n\n![[inner.md]]\n")
        _write(tmp_path / "inner.md", "# Inner\n\nDeep content.\n")

        resolved = resolve_embedded_links(extracted, src)
        assert "Deep content." in resolved
        assert "![[outer.md]]" not in resolved
        assert "![[inner.md]]" not in resolved

    def test_embed_missing_file_placeholder(self, tmp_path: Path) -> None:
        """Missing embedded file produces a placeholder."""
        src = _write(tmp_path / "output.md", "")
        extracted = "# PDF\n\n![[missing.md]]\n"

        resolved = resolve_embedded_links(extracted, src, on_not_found="placeholder")
        assert "File not found" in resolved

    def test_no_embeds_passthrough(self, tmp_path: Path) -> None:
        """Markdown without embeds passes through unchanged."""
        src = _write(tmp_path / "output.md", "")
        text = "# Plain\n\nJust text, no links.\n"
        assert resolve_embedded_links(text, src) == text


class TestConverterEmbedIntegration:
    """Verify that the layout converter's markdown output can be post-processed
    through the shared embed resolver without deep imports."""

    def test_embed_resolver_accessible_from_plugin(self) -> None:
        """The layout converter can import and use ``resolve_embedded_links``."""
        from docwen_core.links import resolve_embedded_links as fn

        assert callable(fn)

    def test_resolver_returns_string(self, tmp_path: Path) -> None:
        """Smoke: resolver returns a string for any input."""
        src = _write(tmp_path / "dummy.md", "# X\n")
        result = resolve_embedded_links("text", src)
        assert isinstance(result, str)

    def test_layout_markdown_postprocess_keeps_image_syntax(self, tmp_path: Path) -> None:
        """Final Markdown must not expose internal image capability markers."""
        from docwen_plugin_layout.to_markdown.converter import _resolve_md_links

        src = _write(tmp_path / "output.md", "")
        _write(tmp_path / "pixel.png", "image fixture")
        text = "![[pixel.png|Alt|20x10]] and ![Alt](pixel.png)"

        result = _resolve_md_links(text, src)

        assert result == text
        assert "{{IMAGE" not in result


class TestEmbedEdgeCases:
    """Edge-case behaviour for embed resolution in Markdown content."""

    def test_empty_extracted_text(self, tmp_path: Path) -> None:
        src = _write(tmp_path / "output.md", "")
        assert resolve_embedded_links("", src) == ""

    def test_embed_with_whitespace_target(self, tmp_path: Path) -> None:
        src = _write(tmp_path / "output.md", "")
        result = resolve_embedded_links("![[   ]]", src, on_not_found="placeholder")
        assert isinstance(result, str)
        # Whitespace-only target with no recognisable file type is left as-is
        # or produces a placeholder depending on resolution outcome.
        assert result in ("![[   ]]", "") or "File not found" in result

    def test_embed_link_not_wiki_format_untouched(self, tmp_path: Path) -> None:
        """Standard Markdown image syntax ``![alt](url)`` is not touched."""
        src = _write(tmp_path / "output.md", "")
        text = "![alt](img.png)"
        assert resolve_embedded_links(text, src) == text

    def test_plain_link_untouched(self, tmp_path: Path) -> None:
        """``[[link]]`` without ``!`` prefix is not touched."""
        src = _write(tmp_path / "output.md", "")
        text = "See [[page]] for more."
        assert resolve_embedded_links(text, src) == text

    def test_multiple_embeds_with_mixed_anchors(self, tmp_path: Path) -> None:
        """Mix of full-file, section, and block embeds in one document."""
        src = _write(tmp_path / "out.md", "")
        _write(tmp_path / "a.md", "Content A.\n")
        _write(tmp_path / "b.md", "# B\n\n## S1\n\nSection 1 text.\n\n## S2\n\nS2.\n")
        _write(tmp_path / "c.md", "# C\n\nBlock text. ^bk\n")

        text = "![[a.md]]\n\n![[b.md#S1]]\n\n![[c.md#^bk]]"
        resolved = resolve_embedded_links(text, src)

        assert "Content A." in resolved
        assert "Section 1 text." in resolved
        assert "Block text." in resolved
        assert "![[a.md]]" not in resolved
        assert "![[b.md#S1]]" not in resolved
        assert "![[c.md#^bk]]" not in resolved


class TestOrchestratorIntegration:
    """Verify that the layout converter can use ``process_markdown_links``
    as the single entry point for all link processing."""

    def test_orchestrator_accessible_from_plugin(self) -> None:
        """The layout converter can import ``process_markdown_links``."""
        from docwen_core.links import process_markdown_links as fn

        assert callable(fn)

    def test_orchestrator_handles_embeds_and_non_embeds(self, tmp_path: Path) -> None:
        """Single call resolves embeds AND processes non-embed links."""
        from docwen_core.export_semantics import LinkRuntimeConfig
        from docwen_core.links import process_markdown_links

        src = _write(tmp_path / "doc.md", "")
        _write(tmp_path / "embed.md", "# E\n\nEmbedded.\n")
        _write(tmp_path / "ref.md", "# Ref\n")

        text = "![[embed.md]]\n\nSee [[ref.md|Reference]]."
        result = process_markdown_links(
            text,
            src,
            link_config=LinkRuntimeConfig(non_embed_wiki_mode="extract_text"),
            target_format="md",
        )

        assert "Embedded." in result
        assert "Reference" in result
        assert "![[embed.md]]" not in result
        assert "[[ref.md|Reference]]" not in result

    def test_orchestrator_respects_modes(self, tmp_path: Path) -> None:
        """Wiki and markdown modes are independently configurable."""
        from docwen_core.export_semantics import LinkRuntimeConfig
        from docwen_core.links import process_markdown_links

        src = _write(tmp_path / "doc.md", "")
        text = "Wiki [[p.md|P]] and MD [link](http://x.com)."
        result = process_markdown_links(
            text,
            src,
            link_config=LinkRuntimeConfig(
                non_embed_wiki_mode="extract_text",
                non_embed_markdown_mode="remove",
            ),
            target_format="md",
        )
        assert "P" in result
        assert "[[p.md|P]]" not in result
        assert "[link](http://x.com)" not in result


class TestNonEmbedWikiLinks:
    """Verify non-embed ``[[link]]`` processing through the shared utilities
    as the layout converter would invoke them."""

    def test_wiki_non_embed_extract_text(self, tmp_path: Path) -> None:
        """``[[target|display]]`` → display text."""
        from docwen_core.links import _process_non_embed_links

        src = _write(tmp_path / "doc.md", "")
        text = "See [[page.md|The Page]] for details."
        result = _process_non_embed_links(text, source_file_path=src, wiki_mode="extract_text")
        assert "The Page" in result
        assert "[[page.md|The Page]]" not in result

    def test_wiki_non_embed_remove(self, tmp_path: Path) -> None:
        """``[[target|display]]`` → removed."""
        from docwen_core.links import _process_non_embed_links

        src = _write(tmp_path / "doc.md", "")
        text = "See [[page.md|The Page]]."
        result = _process_non_embed_links(text, source_file_path=src, wiki_mode="remove")
        assert "[[page.md|The Page]]" not in result

    def test_wiki_non_embed_keep(self, tmp_path: Path) -> None:
        """wiki_mode='keep' leaves links unchanged."""
        from docwen_core.links import _process_non_embed_links

        src = _write(tmp_path / "doc.md", "")
        text = "See [[page.md|The Page]]."
        result = _process_non_embed_links(text, source_file_path=src, wiki_mode="keep")
        assert result == text

    def test_wiki_non_embed_resolve(self, tmp_path: Path) -> None:
        """wiki_mode='resolve' produces ``[display](path)`` Markdown links."""
        from docwen_core.links import _process_non_embed_links

        src = _write(tmp_path / "doc.md", "")
        _write(tmp_path / "page.md", "# Page\n")
        text = "See [[page.md|The Page]]."
        result = _process_non_embed_links(text, source_file_path=src, wiki_mode="resolve")
        assert "[The Page](<page.md>)" in result
        assert "[[page.md|The Page]]" not in result

    def test_wiki_link_without_display(self, tmp_path: Path) -> None:
        """``[[target]]`` without display text."""
        from docwen_core.links import _process_non_embed_links

        src = _write(tmp_path / "doc.md", "")
        text = "See [[page.md]]."
        result = _process_non_embed_links(text, source_file_path=src, wiki_mode="extract_text")
        assert "page.md" in result
        assert "[[page.md]]" not in result

    def test_multiple_wiki_links(self, tmp_path: Path) -> None:
        """Multiple wiki links in one text are all processed."""
        from docwen_core.links import _process_non_embed_links

        src = _write(tmp_path / "doc.md", "")
        text = "See [[a.md|A]] and [[b.md|B]]."
        result = _process_non_embed_links(text, source_file_path=src, wiki_mode="extract_text")
        assert "A" in result
        assert "B" in result
        assert "[[" not in result


class TestNonEmbedMarkdownLinks:
    """Verify non-embed ``[text](url)`` processing through shared utilities."""

    def test_markdown_non_embed_extract_text(self, tmp_path: Path) -> None:
        """``[text](url)`` → text."""
        from docwen_core.links import _process_non_embed_links

        src = _write(tmp_path / "doc.md", "")
        text = "Visit [example](https://example.com)."
        result = _process_non_embed_links(text, source_file_path=src, markdown_mode="extract_text")
        assert "example" in result
        assert "[example](https://example.com)" not in result

    def test_markdown_non_embed_remove(self, tmp_path: Path) -> None:
        """``[text](url)`` → removed."""
        from docwen_core.links import _process_non_embed_links

        src = _write(tmp_path / "doc.md", "")
        text = "Visit [example](https://example.com)."
        result = _process_non_embed_links(text, source_file_path=src, markdown_mode="remove")
        assert "[example](https://example.com)" not in result

    def test_markdown_non_embed_keep(self, tmp_path: Path) -> None:
        """markdown_mode='keep' leaves links unchanged."""
        from docwen_core.links import _process_non_embed_links

        src = _write(tmp_path / "doc.md", "")
        text = "Visit [example](https://example.com)."
        result = _process_non_embed_links(text, source_file_path=src, markdown_mode="keep")
        assert result == text

    def test_markdown_links_with_nested_parens(self, tmp_path: Path) -> None:
        """URLs with parentheses are handled correctly."""
        from docwen_core.links import _process_non_embed_links

        src = _write(tmp_path / "doc.md", "")
        text = "See [wiki](https://en.wikipedia.org/wiki/Python_(programming_language))."
        result = _process_non_embed_links(text, source_file_path=src, markdown_mode="extract_text")
        assert "wiki" in result
        assert not result.startswith("See (")

    def test_code_blocks_protect_links(self, tmp_path: Path) -> None:
        """Wiki and markdown links inside fenced code blocks are untouched."""
        from docwen_core.links import _process_non_embed_links

        src = _write(tmp_path / "doc.md", "")
        text = "```\n[[not_a_link]] and [text](url)\n```\nReal [[link.md|Link]].\n"
        result = _process_non_embed_links(
            text,
            source_file_path=src,
            wiki_mode="extract_text",
            markdown_mode="extract_text",
        )
        assert "[[not_a_link]]" in result
        assert "[text](url)" in result
        assert "Link" in result


class TestIsRemoteUrl:
    """Tests for ``_is_remote_url`` — URL scheme detection."""

    def test_http_is_remote(self) -> None:
        from docwen_plugin_layout.preprocess import _is_remote_url

        assert _is_remote_url("http://example.com/img.png")
        assert _is_remote_url("https://cdn.example.com/photo.jpg")

    def test_local_path_is_not_remote(self) -> None:
        from docwen_plugin_layout.preprocess import _is_remote_url

        assert not _is_remote_url("images/photo.png")
        assert not _is_remote_url("/abs/path/img.jpg")
        assert not _is_remote_url("file:///C:/Users/test/img.png")

    def test_data_uri_is_not_remote(self) -> None:
        from docwen_plugin_layout.preprocess import _is_remote_url

        assert not _is_remote_url("data:image/png;base64,iVBORw0KGgo=")

    def test_ftp_is_not_http_remote(self) -> None:
        from docwen_plugin_layout.preprocess import _is_remote_url

        assert not _is_remote_url("ftp://files.example.com/img.png")


class TestExtractBaseHref:
    """Tests for ``_extract_base_href`` — ``<base href>`` extraction."""

    def test_extracts_href(self) -> None:
        from docwen_plugin_layout.preprocess import _extract_base_href

        html = '<html><head><base href="http://example.com/docs/"></head><body></body></html>'
        assert _extract_base_href(html) == "http://example.com/docs/"

    def test_no_base_tag_returns_none(self) -> None:
        from docwen_plugin_layout.preprocess import _extract_base_href

        assert _extract_base_href("<html><head></head><body></body></html>") is None

    def test_empty_html_returns_none(self) -> None:
        from docwen_plugin_layout.preprocess import _extract_base_href

        assert _extract_base_href("") is None

    def test_empty_href_returns_none(self) -> None:
        from docwen_plugin_layout.preprocess import _extract_base_href

        html = '<html><head><base href=""></head><body></body></html>'
        assert _extract_base_href(html) is None

    def test_single_quoted_href(self) -> None:
        from docwen_plugin_layout.preprocess import _extract_base_href

        html = "<html><head><base href='http://x.com/'></head></html>"
        assert _extract_base_href(html) == "http://x.com/"

    def test_file_scheme_href(self) -> None:
        from docwen_plugin_layout.preprocess import _extract_base_href

        html = '<html><head><base href="file:///C:/docs/"></head></html>'
        assert _extract_base_href(html) == "file:///C:/docs/"


class TestResolveLocalPath:
    """Tests for ``_resolve_local_path`` — local image path resolution."""

    def test_relative_to_html_dir(self, tmp_path: Path) -> None:
        from docwen_plugin_layout.preprocess import _resolve_local_path

        html_file = tmp_path / "doc.html"
        html_file.write_text("")
        img = tmp_path / "photo.png"
        img.write_text("fake")

        result = _resolve_local_path(
            src="photo.png",
            html_path=str(html_file),
            base_href=None,
            resource_dir=None,
        )
        assert result == img.resolve()

    def test_relative_with_resource_dir(self, tmp_path: Path) -> None:
        from docwen_plugin_layout.preprocess import _resolve_local_path

        html_file = tmp_path / "doc.html"
        html_file.write_text("")
        res_dir = tmp_path / "resources"
        res_dir.mkdir()
        img = res_dir / "img.gif"
        img.write_text("fake")

        result = _resolve_local_path(
            src="img.gif",
            html_path=str(html_file),
            base_href=None,
            resource_dir=str(res_dir),
        )
        assert result == img.resolve()

    def test_file_scheme_uri(self, tmp_path: Path) -> None:
        from docwen_plugin_layout.preprocess import _resolve_local_path

        img = tmp_path / "x.png"
        img.write_text("fake")
        src_uri = img.as_uri()

        result = _resolve_local_path(
            src=src_uri,
            html_path=str(tmp_path / "doc.html"),
            base_href=None,
            resource_dir=None,
        )
        assert result is not None
        assert result.exists()

    def test_non_file_scheme_returns_none(self) -> None:
        from docwen_plugin_layout.preprocess import _resolve_local_path

        result = _resolve_local_path(
            src="ftp://server/img.png",
            html_path="/tmp/doc.html",
            base_href=None,
            resource_dir=None,
        )
        assert result is None

    def test_absolute_path_under_base_href(self, tmp_path: Path) -> None:
        from docwen_plugin_layout.preprocess import _resolve_local_path

        root = tmp_path / "root"
        root.mkdir()
        img = root / "assets" / "logo.png"
        img.parent.mkdir()
        img.write_text("fake")

        result = _resolve_local_path(
            src="/assets/logo.png",
            html_path=str(tmp_path / "doc.html"),
            base_href=(root / "index.html").as_uri(),
            resource_dir=None,
        )
        assert result == img

    def test_encoded_path(self, tmp_path: Path) -> None:
        from docwen_plugin_layout.preprocess import _resolve_local_path

        html_file = tmp_path / "doc.html"
        html_file.write_text("")
        img = tmp_path / "my photo.png"
        img.write_text("fake")

        result = _resolve_local_path(
            src="my%20photo.png",
            html_path=str(html_file),
            base_href=None,
            resource_dir=None,
        )
        assert result == img.resolve()

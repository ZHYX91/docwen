"""Focused tests split from test_markdown_preprocess.py."""

from __future__ import annotations

import pytest

from ._markdown_preprocess_support import (
    Path,
    _write_image,
)

pytestmark = pytest.mark.unit


class TestConvertHtmlToMarkdownText:
    """Integration tests for ``convert_html_to_markdown_text`` — the
    complete HTML→Markdown pipeline with image preprocessing."""

    def test_html_without_images(self, tmp_path: Path) -> None:
        from docwen_plugin_layout.to_markdown.converter import (
            convert_html_to_markdown_text,
        )

        html = "<html><body><h1>Hello</h1><p>World.</p></body></html>"
        out_dir = tmp_path / "out"
        out_dir.mkdir()

        result = convert_html_to_markdown_text(
            html_text=html,
            html_path=str(tmp_path / "doc.html"),
            output_folder=str(out_dir),
        )
        assert "# Hello" in result or "Hello" in result
        assert "World" in result

    def test_html_with_title_extraction(self, tmp_path: Path) -> None:
        from docwen_plugin_layout.to_markdown.converter import (
            convert_html_to_markdown_text,
        )

        html = "<html><head><title>My Document</title></head><body><p>Content.</p></body></html>"
        out_dir = tmp_path / "out"
        out_dir.mkdir()

        result = convert_html_to_markdown_text(
            html_text=html,
            html_path=str(tmp_path / "doc.html"),
            output_folder=str(out_dir),
        )
        assert result.startswith("# My Document")

    def test_html_with_local_image(self, tmp_path: Path) -> None:
        from docwen_plugin_layout.to_markdown.converter import (
            convert_html_to_markdown_text,
        )

        _write_image(tmp_path / "photo.png", "PNG")
        html_file = tmp_path / "doc.html"
        html_file.write_text("")
        out_dir = tmp_path / "out"
        out_dir.mkdir()

        html = '<html><body><p>See below:</p><img src="photo.png"></body></html>'
        result = convert_html_to_markdown_text(
            html_text=html,
            html_path=str(html_file),
            output_folder=str(out_dir),
            image_link_style="wiki_embed",
        )
        # The result should contain a wiki-embed image link
        assert "![[" in result
        assert ".png]]" in result
        # The image should be copied to output
        assert len(list(out_dir.iterdir())) >= 1

    def test_html_with_remote_image(self, tmp_path: Path) -> None:
        from docwen_plugin_layout.to_markdown.converter import (
            convert_html_to_markdown_text,
        )

        html = '<html><body><img src="https://cdn.example.com/photo.jpg"></body></html>'
        out_dir = tmp_path / "out"
        out_dir.mkdir()

        result = convert_html_to_markdown_text(
            html_text=html,
            html_path=str(tmp_path / "doc.html"),
            output_folder=str(out_dir),
            image_link_style="markdown_embed",
        )
        assert "![https://cdn.example.com/photo.jpg](https://cdn.example.com/photo.jpg)" in result

    def test_html_with_data_uri_image(self, tmp_path: Path) -> None:
        from docwen_plugin_layout.to_markdown.converter import (
            convert_html_to_markdown_text,
        )

        png_b64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=="
        data_uri = f"data:image/png;base64,{png_b64}"
        html = f'<html><body><img src="{data_uri}"></body></html>'
        out_dir = tmp_path / "out"
        out_dir.mkdir()

        result = convert_html_to_markdown_text(
            html_text=html,
            html_path=str(tmp_path / "doc.html"),
            output_folder=str(out_dir),
            keep_images=True,
        )
        assert "![[" in result or "![" in result
        # A PNG file should have been materialised
        png_files = list(out_dir.glob("*.png"))
        assert len(png_files) >= 1

    def test_html_with_base_href(self, tmp_path: Path) -> None:
        from docwen_plugin_layout.to_markdown.converter import (
            convert_html_to_markdown_text,
        )

        html = (
            '<html><head><base href="https://cdn.example.com/site/"></head>'
            '<body><img src="images/logo.png"></body></html>'
        )
        out_dir = tmp_path / "out"
        out_dir.mkdir()

        result = convert_html_to_markdown_text(
            html_text=html,
            html_path=str(tmp_path / "doc.html"),
            output_folder=str(out_dir),
        )
        assert "cdn.example.com/site/images/logo.png" in result

    def test_keep_images_false_skips_local_save(self, tmp_path: Path) -> None:
        from docwen_plugin_layout.to_markdown.converter import (
            convert_html_to_markdown_text,
        )

        _write_image(tmp_path / "pic.png", "PNG")
        html_file = tmp_path / "doc.html"
        html_file.write_text("")
        out_dir = tmp_path / "out"
        out_dir.mkdir()

        html = '<html><body><img src="pic.png"></body></html>'
        convert_html_to_markdown_text(
            html_text=html,
            html_path=str(html_file),
            output_folder=str(out_dir),
            keep_images=False,
        )
        # No new files in output (keep_images=False → no copy)
        png_files = list(out_dir.glob("*.png"))
        assert len(png_files) == 0

    def test_user_path_custom_link_styles(self, tmp_path: Path) -> None:
        """End-to-end user path: HTML with multiple images, different
        link styles, verify all images appear in the output."""
        from docwen_plugin_layout.to_markdown.converter import (
            convert_html_to_markdown_text,
        )

        _write_image(tmp_path / "a.png", "PNG")
        _write_image(tmp_path / "b.jpg", "JPEG")
        html_file = tmp_path / "page.html"
        html_file.write_text("")
        out_dir = tmp_path / "out"
        out_dir.mkdir()

        html = (
            "<html><head><title>Gallery</title></head><body>"
            "<h1>Images</h1>"
            '<img src="a.png">'
            "<p>Middle text.</p>"
            '<img src="https://remote.example.com/banner.gif">'
            "</body></html>"
        )
        result = convert_html_to_markdown_text(
            html_text=html,
            html_path=str(html_file),
            output_folder=str(out_dir),
            image_link_style="wiki_embed",
            md_file_link_style="wiki_embed",
        )
        # Title
        assert result.startswith("# Gallery")
        # Local image (wiki embed)
        assert ".png]]" in result
        # Remote image (wiki embed)
        assert "remote.example.com/banner.gif]]" in result
        # Local file was materialised
        png_files = list(out_dir.glob("*.png"))
        assert len(png_files) >= 1

"""Tests for build_image_ocr_sidecar helper."""

from __future__ import annotations

import pytest

from docwen_core.text.image_markdown import build_image_ocr_sidecar

pytestmark = pytest.mark.unit


class TestBuildImageOcrSidecar:
    """Unit tests for the shared sidecar builder."""

    def test_returns_two_strings(self):
        sidecar_text, replacement_link = build_image_ocr_sidecar(
            sidecar_stem="test_img_001_ocr",
            source_format="docx",
            image_markdown="![[foo.png]]\n",
            ocr_text="Hello world",
            md_link_style="wiki_embed",
        )
        assert isinstance(sidecar_text, str)
        assert isinstance(replacement_link, str)
        assert len(sidecar_text) > 0
        assert len(replacement_link) > 0

    def test_sidecar_contains_front_matter(self):
        sidecar_text, _ = build_image_ocr_sidecar(
            sidecar_stem="test_img_001_ocr",
            source_format="docx",
            image_markdown="![[img.png]]\n",
            ocr_text="Sample OCR",
            md_link_style="wiki_embed",
        )
        assert "---" in sidecar_text
        assert "title: test_img_001_ocr" in sidecar_text
        assert "source_format: docx" in sidecar_text
        assert "ocr: True" in sidecar_text

    def test_sidecar_contains_image_link(self):
        sidecar_text, _ = build_image_ocr_sidecar(
            sidecar_stem="img_ocr",
            source_format="docx",
            image_markdown="![[foo/bar.png]]\n",
            ocr_text="Text",
            md_link_style="wiki_embed",
        )
        assert "![[foo/bar.png]]" in sidecar_text

    def test_sidecar_contains_ocr_blockquote(self):
        sidecar_text, _ = build_image_ocr_sidecar(
            sidecar_stem="img_ocr",
            source_format="docx",
            image_markdown="![[img.png]]\n",
            ocr_text="Line 1\nLine 2",
            md_link_style="wiki_embed",
        )
        assert "> Line 1" in sidecar_text
        assert "> Line 2" in sidecar_text

    def test_replacement_link_wiki_embed(self):
        _, link = build_image_ocr_sidecar(
            sidecar_stem="img_001_ocr",
            source_format="docx",
            image_markdown="![[img.png]]\n",
            ocr_text="Text",
            md_link_style="wiki_embed",
        )
        assert link.startswith("![[")

    def test_replacement_link_markdown_link(self):
        _, link = build_image_ocr_sidecar(
            sidecar_stem="img_001_ocr",
            source_format="docx",
            image_markdown="![[img.png]]\n",
            ocr_text="Text",
            md_link_style="markdown_link",
        )
        assert link.startswith("[")

    def test_replacement_link_contains_sidecar_name(self):
        _, link = build_image_ocr_sidecar(
            sidecar_stem="my_doc__img_001_ocr",
            source_format="docx",
            image_markdown="![[img.png]]\n",
            ocr_text="Text",
            md_link_style="wiki_embed",
        )
        assert "my_doc__img_001_ocr" in link

    def test_pure_function_no_side_effects(self):
        """Calling twice with same inputs produces identical outputs."""
        kwargs = {
            "sidecar_stem": "s",
            "source_format": "xlsx",
            "image_markdown": "![[a.png]]\n",
            "ocr_text": "T",
            "md_link_style": "markdown_link",
        }
        a_text, a_link = build_image_ocr_sidecar(**kwargs)
        b_text, b_link = build_image_ocr_sidecar(**kwargs)
        assert a_text == b_text
        assert a_link == b_link

    def test_ocr_blockquote_title(self):
        sidecar_text, _ = build_image_ocr_sidecar(
            sidecar_stem="s",
            source_format="png",
            image_markdown="![[img.png]]\n",
            ocr_text="Hello",
            md_link_style="wiki_embed",
            ocr_blockquote_title="OCR结果",
        )
        assert "> **OCR结果**" in sidecar_text
        assert "> Hello" in sidecar_text

    def test_no_title_produces_plain_blockquote(self):
        sidecar_text, _ = build_image_ocr_sidecar(
            sidecar_stem="s",
            source_format="png",
            image_markdown="![[img.png]]\n",
            ocr_text="Hello",
            md_link_style="wiki_embed",
        )
        assert "> **" not in sidecar_text
        assert "> Hello" in sidecar_text

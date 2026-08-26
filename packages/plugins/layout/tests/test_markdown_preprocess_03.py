"""Focused tests split from test_markdown_preprocess.py."""

from __future__ import annotations

from ._markdown_preprocess_support import (
    Path,
    _write_image,
    pytest,
)

pytestmark = pytest.mark.unit


class TestBuildImageMarkdown:
    """Tests for ``build_image_markdown`` — image→Markdown-link conversion."""

    def test_remote_url_markdown_embed_style(self) -> None:
        from docwen_plugin_layout.preprocess import build_image_markdown

        result = build_image_markdown(
            src="https://example.com/img.png",
            html_path="/tmp/doc.html",
            base_href=None,
            resource_dir=None,
            output_folder="/tmp/out",
            original_basename="doc",
            image_index=1,
            unified_timestamp_desc="export",
            keep_images=True,
            image_link_style="markdown_embed",
        )
        assert result == "![https://example.com/img.png](https://example.com/img.png)"

    def test_remote_url_wiki_embed_style(self) -> None:
        from docwen_plugin_layout.preprocess import build_image_markdown

        result = build_image_markdown(
            src="https://example.com/img.png",
            html_path="/tmp/doc.html",
            base_href=None,
            resource_dir=None,
            output_folder="/tmp/out",
            original_basename="doc",
            image_index=1,
            unified_timestamp_desc="export",
            keep_images=True,
            image_link_style="wiki_embed",
        )
        assert result == "![[https://example.com/img.png]]"

    def test_keep_images_false_remote_returns_empty(self) -> None:
        from docwen_plugin_layout.preprocess import build_image_markdown

        result = build_image_markdown(
            src="https://example.com/img.png",
            html_path="/tmp/doc.html",
            base_href=None,
            resource_dir=None,
            output_folder="/tmp/out",
            original_basename="doc",
            image_index=1,
            unified_timestamp_desc="export",
            keep_images=False,
        )
        assert result == ""

    def test_local_image_wiki_embed(self, tmp_path: Path) -> None:
        from docwen_plugin_layout.preprocess import build_image_markdown

        _write_image(tmp_path / "photo.png", "PNG")
        html_file = tmp_path / "doc.html"
        html_file.write_text("")
        out_dir = tmp_path / "out"
        out_dir.mkdir()

        result = build_image_markdown(
            src="photo.png",
            html_path=str(html_file),
            base_href=None,
            resource_dir=None,
            output_folder=str(out_dir),
            original_basename="doc",
            image_index=1,
            unified_timestamp_desc="export",
            keep_images=True,
            image_link_style="wiki_embed",
        )
        assert result.startswith("![[")
        assert result.endswith(".png]]")

    def test_local_image_copies_on_disk(self, tmp_path: Path) -> None:
        from docwen_plugin_layout.preprocess import build_image_markdown

        _write_image(tmp_path / "logo.jpg", "JPEG")
        html_file = tmp_path / "doc.html"
        html_file.write_text("")
        out_dir = tmp_path / "out"
        out_dir.mkdir()

        result = build_image_markdown(
            src="logo.jpg",
            html_path=str(html_file),
            base_href=None,
            resource_dir=None,
            output_folder=str(out_dir),
            original_basename="doc",
            image_index=2,
            unified_timestamp_desc="v2",
            keep_images=True,
            image_link_style="markdown_embed",
        )
        # Extract filename from markdown link
        import re

        m = re.search(r"!\[([^\]]*)\]\(([^)]+)\)", result)
        assert m is not None
        filename = m.group(2)
        assert (out_dir / filename).exists()

    def test_missing_local_falls_back_to_src(self, tmp_path: Path) -> None:
        from docwen_plugin_layout.preprocess import build_image_markdown

        html_file = tmp_path / "doc.html"
        html_file.write_text("")
        out_dir = tmp_path / "out"
        out_dir.mkdir()

        result = build_image_markdown(
            src="nonexistent.jpg",
            html_path=str(html_file),
            base_href=None,
            resource_dir=None,
            output_folder=str(out_dir),
            original_basename="doc",
            image_index=1,
            unified_timestamp_desc="export",
            keep_images=True,
            image_link_style="wiki_embed",
        )
        # Falls back to src-as-link
        assert "nonexistent.jpg" in result

    def test_data_uri_with_keep_images(self, tmp_path: Path) -> None:
        from docwen_plugin_layout.preprocess import build_image_markdown

        png_b64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=="
        data_uri = f"data:image/png;base64,{png_b64}"

        html_file = tmp_path / "doc.html"
        html_file.write_text("")
        out_dir = tmp_path / "out"
        out_dir.mkdir()

        result = build_image_markdown(
            src=data_uri,
            html_path=str(html_file),
            base_href=None,
            resource_dir=None,
            output_folder=str(out_dir),
            original_basename="doc",
            image_index=1,
            unified_timestamp_desc="export",
            keep_images=True,
            image_link_style="wiki_embed",
        )
        assert result.startswith("![[")
        # A file should have been created in output
        files = list(out_dir.iterdir())
        assert len(files) >= 1

    def test_base_href_urljoin_local(self, tmp_path: Path) -> None:
        from docwen_plugin_layout.preprocess import build_image_markdown

        html_file = tmp_path / "doc.html"
        html_file.write_text("")
        out_dir = tmp_path / "out"
        out_dir.mkdir()

        result = build_image_markdown(
            src="images/pic.png",
            html_path=str(html_file),
            base_href="https://cdn.example.com/site/",
            resource_dir=None,
            output_folder=str(out_dir),
            original_basename="doc",
            image_index=1,
            unified_timestamp_desc="export",
            keep_images=True,
            image_link_style="markdown_embed",
        )
        assert "cdn.example.com/site/images/pic.png" in result

    def test_markdown_link_style(self) -> None:
        from docwen_plugin_layout.preprocess import build_image_markdown

        result = build_image_markdown(
            src="https://x.com/a.png",
            html_path="/tmp/doc.html",
            base_href=None,
            resource_dir=None,
            output_folder="/tmp/out",
            original_basename="doc",
            image_index=1,
            unified_timestamp_desc="export",
            keep_images=True,
            image_link_style="markdown_link",
        )
        assert result == "[https://x.com/a.png](https://x.com/a.png)"

    def test_wiki_link_style(self) -> None:
        from docwen_plugin_layout.preprocess import build_image_markdown

        result = build_image_markdown(
            src="https://x.com/a.png",
            html_path="/tmp/doc.html",
            base_href=None,
            resource_dir=None,
            output_folder="/tmp/out",
            original_basename="doc",
            image_index=1,
            unified_timestamp_desc="export",
            keep_images=True,
            image_link_style="wiki_link",
        )
        assert result == "[[https://x.com/a.png]]"

    def test_ocr_uses_request_language_and_locale(
        self,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        import docwen_core.text.ocr as core_ocr
        from docwen_plugin_layout.preprocess import build_image_markdown

        calls: list[tuple[str, str, str]] = []

        def fake_ocr(
            image_path: str | Path,
            ocr_language: str | None = None,
            *,
            source_format: str,
            current_locale: str = "zh_CN",
            model_dir: str | Path | None = None,
        ) -> core_ocr.OcrOutcome:
            del model_dir
            assert source_format == "png"
            calls.append((Path(image_path).name, ocr_language or "", current_locale))
            return core_ocr.OcrOutcome(core_ocr.OcrStatus.SUCCESS, text="LAYOUT HTML OCR")

        monkeypatch.setattr(core_ocr, "run_ocr_outcome", fake_ocr)

        _write_image(tmp_path / "pic.png", "PNG")
        html_file = tmp_path / "doc.html"
        html_file.write_text("")
        out_dir = tmp_path / "out"
        out_dir.mkdir()

        result = build_image_markdown(
            src="pic.png",
            html_path=str(html_file),
            base_href=None,
            resource_dir=None,
            output_folder=str(out_dir),
            original_basename="doc",
            image_index=1,
            unified_timestamp_desc="export",
            keep_images=True,
            enable_ocr=True,
            ocr_language="latin",
            current_locale="de_DE",
        )

        assert "LAYOUT HTML OCR" in result
        assert calls == [(next(out_dir.iterdir()).name, "latin", "de_DE")]

    def test_ocr_failure_preserves_local_image_link(
        self,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        import docwen_core.text.ocr as core_ocr
        from docwen_plugin_layout.preprocess import build_image_markdown

        def missing_model(*_args: object, **_kwargs: object) -> core_ocr.OcrOutcome:
            return core_ocr.OcrOutcome(core_ocr.OcrStatus.MODEL_MISSING, message="rapidocr model missing")

        monkeypatch.setattr(core_ocr, "run_ocr_outcome", missing_model)
        _write_image(tmp_path / "pic.png", "PNG")
        html_file = tmp_path / "doc.html"
        html_file.write_text("")
        out_dir = tmp_path / "out"
        out_dir.mkdir()

        result = build_image_markdown(
            src="pic.png",
            html_path=str(html_file),
            base_href=None,
            resource_dir=None,
            output_folder=str(out_dir),
            original_basename="doc",
            image_index=1,
            unified_timestamp_desc="export",
            keep_images=True,
            enable_ocr=True,
        )

        assert next(out_dir.iterdir()).name in result


class TestPreprocessHtmlImages:
    """Tests for ``preprocess_html_images`` — HTML ``<img>`` preprocessing."""

    def test_single_img_replaced_with_token(self) -> None:
        from docwen_plugin_layout.preprocess import preprocess_html_images

        html = '<html><body><img src="photo.png"></body></html>'
        result = preprocess_html_images(
            html_text=html,
            html_path="/tmp/doc.html",
            output_folder="/tmp/out",
        )
        assert "<img" not in result["html"]
        assert "DOCWENIMG0001" in result["html"]
        assert "DOCWENIMG0001" in result["token_map"]

    def test_multiple_imgs_get_unique_tokens(self) -> None:
        from docwen_plugin_layout.preprocess import preprocess_html_images

        html = '<html><body><img src="a.png"><img src="b.png"><img src="c.png"></body></html>'
        result = preprocess_html_images(
            html_text=html,
            html_path="/tmp/doc.html",
            output_folder="/tmp/out",
        )
        assert "DOCWENIMG0001" in result["html"]
        assert "DOCWENIMG0002" in result["html"]
        assert "DOCWENIMG0003" in result["html"]
        assert len(result["token_map"]) == 3
        for token in result["token_map"]:
            assert token in result["html"]

    def test_img_without_src_left_unchanged(self) -> None:
        from docwen_plugin_layout.preprocess import preprocess_html_images

        html = '<html><body><img alt="no src"></body></html>'
        result = preprocess_html_images(
            html_text=html,
            html_path="/tmp/doc.html",
            output_folder="/tmp/out",
        )
        assert "<img" in result["html"]
        assert len(result["token_map"]) == 0

    def test_no_img_tags_passthrough(self) -> None:
        from docwen_plugin_layout.preprocess import preprocess_html_images

        html = "<html><body><p>No images here.</p></body></html>"
        result = preprocess_html_images(
            html_text=html,
            html_path="/tmp/doc.html",
            output_folder="/tmp/out",
        )
        assert result["html"] == html
        assert len(result["token_map"]) == 0

    def test_remote_img_produces_remote_link(self) -> None:
        from docwen_plugin_layout.preprocess import preprocess_html_images

        html = '<html><body><img src="https://cdn.example.com/logo.png"></body></html>'
        result = preprocess_html_images(
            html_text=html,
            html_path="/tmp/doc.html",
            output_folder="/tmp/out",
            image_link_style="wiki_embed",
        )
        token_map: object = result["token_map"]  # type helper
        assert isinstance(token_map, dict)
        token = token_map.get("DOCWENIMG0001", "")
        assert "cdn.example.com/logo.png" in token

    def test_base_href_resolution_in_html(self, tmp_path: Path) -> None:
        from docwen_plugin_layout.preprocess import preprocess_html_images

        html = (
            '<html><head><base href="https://cdn.example.com/site/"></head>'
            '<body><img src="images/logo.png"></body></html>'
        )
        out_dir = tmp_path / "out"
        out_dir.mkdir()

        result = preprocess_html_images(
            html_text=html,
            html_path=str(tmp_path / "doc.html"),
            output_folder=str(out_dir),
            image_link_style="markdown_embed",
        )
        token_map2: object = result["token_map"]  # type helper
        assert isinstance(token_map2, dict)
        token = token_map2.get("DOCWENIMG0001", "")
        assert "cdn.example.com/site/images/logo.png" in token

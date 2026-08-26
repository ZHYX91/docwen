"""Focused tests split from test_markdown_preprocess.py."""

from __future__ import annotations

from ._markdown_preprocess_support import (
    Path,
    _write_image,
    pytest,
)

pytestmark = pytest.mark.unit


class TestMakeImageFilename:
    """Tests for ``_make_image_filename`` — structured filename generation."""

    def test_basic_filename(self, tmp_path: Path) -> None:
        from docwen_plugin_layout.preprocess import _make_image_filename

        source = _write_image(tmp_path / "photo.png", "PNG")
        name = _make_image_filename(
            source,
            original_basename="report",
            image_index=1,
            unified_timestamp_desc="20260612",
        )
        assert name.startswith("report_image1_20260612")
        assert name.endswith(".png")

    def test_jpg_extension(self, tmp_path: Path) -> None:
        from docwen_plugin_layout.preprocess import _make_image_filename

        source = _write_image(tmp_path / "img.jpg", "JPEG")
        name = _make_image_filename(
            source,
            original_basename="doc",
            image_index=3,
            unified_timestamp_desc="export",
        )
        assert name.endswith(".jpg")
        assert "doc_image3_export" in name

    def test_unknown_extension_uses_detected_png(self, tmp_path: Path) -> None:
        from docwen_plugin_layout.preprocess import _make_image_filename

        source = _write_image(tmp_path / "data", "PNG")
        name = _make_image_filename(
            source,
            original_basename="file",
            image_index=5,
            unified_timestamp_desc="v1",
        )
        assert name.endswith(".png")

    def test_sanitizes_illegal_chars(self, tmp_path: Path) -> None:
        from docwen_plugin_layout.preprocess import _make_image_filename

        source = _write_image(tmp_path / "img.png", "PNG")
        name = _make_image_filename(
            source,
            original_basename="doc:version",
            image_index=2,
            unified_timestamp_desc="2026/06/12",
        )
        assert ":" not in name
        assert "/" not in name

    def test_misleading_jpg_suffix_uses_detected_png_extension(self, tmp_path: Path) -> None:
        from docwen_plugin_layout.preprocess import _make_image_filename

        source = _write_image(tmp_path / "actually-png.jpg", "PNG")

        assert _make_image_filename(source, "doc", 1, "export").endswith(".png")

    def test_non_image_is_rejected(self, tmp_path: Path) -> None:
        from docwen_plugin_layout.preprocess import _make_image_filename

        source = tmp_path / "not-an-image.png"
        source.write_text("plain text", encoding="utf-8")

        with pytest.raises(ValueError, match="not a supported image"):
            _make_image_filename(source, "doc", 1, "export")


class TestExtractedImageAdmission:
    def test_detected_image_format_ignores_misleading_suffix(self, tmp_path: Path) -> None:
        from docwen_plugin_layout.to_markdown.converter import _detected_image_format

        source = _write_image(tmp_path / "extracted.txt", "JPEG")

        assert _detected_image_format(source) == "jpeg"

    def test_detected_image_format_rejects_non_image_with_png_suffix(self, tmp_path: Path) -> None:
        from docwen_plugin_layout.to_markdown.converter import _detected_image_format

        source = tmp_path / "fake.png"
        source.write_text("not an image", encoding="utf-8")

        assert _detected_image_format(source) is None


class TestCopyLocalImage:
    """Tests for ``_copy_local_image`` — local file copy to output."""

    def test_copies_to_output(self, tmp_path: Path) -> None:
        from docwen_plugin_layout.preprocess import _copy_local_image

        src_img = _write_image(tmp_path / "source.png", "PNG")
        out_dir = tmp_path / "output"
        out_dir.mkdir()

        filename = _copy_local_image(
            src_path=src_img,
            output_folder=str(out_dir),
            original_basename="doc",
            image_index=1,
            unified_timestamp_desc="export",
        )
        assert filename is not None
        copied = out_dir / filename
        assert copied.exists()
        assert copied.read_bytes() == src_img.read_bytes()

    def test_skips_when_target_exists(self, tmp_path: Path) -> None:
        from docwen_plugin_layout.preprocess import _copy_local_image

        src_img = _write_image(tmp_path / "source.png", "PNG", color=(1, 1, 1))
        out_dir = tmp_path / "output"
        out_dir.mkdir()

        # First copy
        filename = _copy_local_image(
            src_path=src_img,
            output_folder=str(out_dir),
            original_basename="doc",
            image_index=1,
            unified_timestamp_desc="export",
        )
        assert filename is not None
        # Modify source
        first_bytes = (out_dir / filename).read_bytes()
        _write_image(src_img, "PNG", color=(2, 2, 2))
        # Second copy should skip
        _copy_local_image(
            src_path=src_img,
            output_folder=str(out_dir),
            original_basename="doc",
            image_index=1,
            unified_timestamp_desc="export",
        )
        copied = out_dir / filename
        # Content should be v1 (first copy), not v2
        assert copied.read_bytes() == first_bytes


class TestMaterializeImageTarget:
    """Tests for ``materialize_image_target`` — image file materialisation
    **without** OCR."""

    def test_remote_url_returns_none_path(self) -> None:
        from docwen_plugin_layout.preprocess import materialize_image_target

        result = materialize_image_target(
            src="https://example.com/photo.jpg",
            html_path="/tmp/doc.html",
            base_href=None,
            resource_dir=None,
            output_folder="/tmp/out",
            original_basename="doc",
            image_index=1,
            unified_timestamp_desc="export",
        )
        assert result["filename"] == "https://example.com/photo.jpg"
        assert result["image_path"] is None

    def test_local_image_copies_to_output(self, tmp_path: Path) -> None:
        from docwen_plugin_layout.preprocess import materialize_image_target

        _write_image(tmp_path / "photo.png", "PNG")
        html_file = tmp_path / "doc.html"
        html_file.write_text("")
        out_dir = tmp_path / "out"
        out_dir.mkdir()

        result = materialize_image_target(
            src="photo.png",
            html_path=str(html_file),
            base_href=None,
            resource_dir=None,
            output_folder=str(out_dir),
            original_basename="doc",
            image_index=1,
            unified_timestamp_desc="export",
        )
        assert result["filename"] is not None
        assert result["image_path"] is not None
        assert Path(result["image_path"]).exists()  # type: ignore[arg-type]

    def test_local_image_uses_content_extension_despite_wrong_suffix(self, tmp_path: Path) -> None:
        from docwen_plugin_layout.preprocess import materialize_image_target

        _write_image(tmp_path / "photo.txt", "JPEG")
        html_file = tmp_path / "doc.html"
        html_file.write_text("", encoding="utf-8")
        out_dir = tmp_path / "out"
        out_dir.mkdir()

        result = materialize_image_target(
            src="photo.txt",
            html_path=str(html_file),
            base_href=None,
            resource_dir=None,
            output_folder=str(out_dir),
            original_basename="doc",
            image_index=1,
            unified_timestamp_desc="export",
        )

        assert str(result["filename"]).endswith(".jpg")
        assert result["image_path"] is not None
        assert Path(result["image_path"]).read_bytes() == (tmp_path / "photo.txt").read_bytes()  # type: ignore[arg-type]

    def test_local_non_image_with_png_suffix_is_not_copied(self, tmp_path: Path) -> None:
        from docwen_plugin_layout.preprocess import materialize_image_target

        source = tmp_path / "fake.png"
        source.write_text("plain text", encoding="utf-8")
        html_file = tmp_path / "doc.html"
        html_file.write_text("", encoding="utf-8")
        out_dir = tmp_path / "out"
        out_dir.mkdir()

        result = materialize_image_target(
            src=source.name,
            html_path=str(html_file),
            base_href=None,
            resource_dir=None,
            output_folder=str(out_dir),
            original_basename="doc",
            image_index=1,
            unified_timestamp_desc="export",
        )

        assert result == {"filename": source.name, "image_path": None}
        assert list(out_dir.iterdir()) == []

    def test_missing_local_returns_original_src(self, tmp_path: Path) -> None:
        from docwen_plugin_layout.preprocess import materialize_image_target

        html_file = tmp_path / "doc.html"
        html_file.write_text("")
        out_dir = tmp_path / "out"
        out_dir.mkdir()

        result = materialize_image_target(
            src="nonexistent.png",
            html_path=str(html_file),
            base_href=None,
            resource_dir=None,
            output_folder=str(out_dir),
            original_basename="doc",
            image_index=1,
            unified_timestamp_desc="export",
        )
        assert result["filename"] == "nonexistent.png"
        assert result["image_path"] is None

    def test_data_uri_png_saves_to_output(self, tmp_path: Path) -> None:
        from docwen_plugin_layout.preprocess import materialize_image_target

        # Minimal valid PNG data URI (1x1 pixel transparent PNG)
        png_b64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=="
        data_uri = f"data:image/png;base64,{png_b64}"

        html_file = tmp_path / "doc.html"
        html_file.write_text("")
        out_dir = tmp_path / "out"
        out_dir.mkdir()

        result = materialize_image_target(
            src=data_uri,
            html_path=str(html_file),
            base_href=None,
            resource_dir=None,
            output_folder=str(out_dir),
            original_basename="doc",
            image_index=1,
            unified_timestamp_desc="export",
        )
        # Should produce a valid filename and existing image_path
        assert result["filename"] != data_uri
        assert result["image_path"] is not None
        assert Path(result["image_path"]).exists()  # type: ignore[arg-type]

    def test_base_href_urljoin(self, tmp_path: Path) -> None:
        from docwen_plugin_layout.preprocess import materialize_image_target

        html_file = tmp_path / "doc.html"
        html_file.write_text("")
        out_dir = tmp_path / "out"
        out_dir.mkdir()

        # When base_href is a remote URL, the relative src gets urljoin'd
        result = materialize_image_target(
            src="images/photo.png",
            html_path=str(html_file),
            base_href="https://cdn.example.com/docs/",
            resource_dir=None,
            output_folder=str(out_dir),
            original_basename="doc",
            image_index=1,
            unified_timestamp_desc="export",
        )
        assert result["filename"] == "https://cdn.example.com/docs/images/photo.png"
        assert result["image_path"] is None

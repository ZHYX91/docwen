"""Focused tests split from test_layout_conversions.py."""

from __future__ import annotations

import pytest

from ._layout_conversions_support import (
    Path,
    _build_fake_context,
    os,
    tempfile,
)

pytestmark = pytest.mark.contract


class TestLayoutToPng:
    def test_pdf_to_png_uses_admitted_format_despite_xps_suffix(self, sample_pdf_path: Path, tmp_path: Path) -> None:
        """Layout rendering consumes FileRef.format rather than the filename."""
        from docwen_plugin_layout.to_image.converter import LayoutToImageConverter

        misleading_path = tmp_path / "admitted-pdf.xps"
        misleading_path.write_bytes(sample_pdf_path.read_bytes())

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(misleading_path),
                staging,
                "png",
                options={"render_dpi": 72},
                source_format="pdf",
            )
            result = LayoutToImageConverter("png").convert(context)

        assert result.success is True, f"unexpected error: {result.error}"
        assert result.artifacts[0].media_type == "image/png"

    def test_single_page_pdf_to_png(self, sample_pdf_path: Path) -> None:
        """Single-page PDF → PNG should produce a valid PNG artifact."""
        from docwen_plugin_layout.to_image.converter import LayoutToImageConverter

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_pdf_path),
                staging,
                "png",
                options={"render_dpi": 150},
            )
            result = LayoutToImageConverter("png").convert(context)

            assert result.success is True, f"unexpected error: {result.error}"
            assert len(result.artifacts) == 1
            artifact = result.artifacts[0]
            assert artifact.kind == "image"
            assert artifact.media_type == "image/png"
            assert artifact.is_primary is True
            assert os.path.isfile(artifact.staging_path)

            # Verify it's a valid PNG
            header = Path(artifact.staging_path).read_bytes()[:8]
            assert header[:4] == b"\x89PNG", f"Not a valid PNG: {header[:4]!r}"

            # Metrics
            assert result.metrics.input_bytes > 0
            assert result.metrics.output_bytes > 0
            assert result.metrics.extra.get("page_count") == 1
            assert result.metrics.extra.get("dpi") == 150
            assert any(d.code == "LAYOUT2IMG-RENDERED" for d in result.diagnostics)

    def test_multi_page_pdf_to_png(self, sample_multi_page_pdf_path: Path) -> None:
        """Multi-page PDF → PNG should produce one artifact per page."""
        from docwen_plugin_layout.to_image.converter import LayoutToImageConverter

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_multi_page_pdf_path),
                staging,
                "png",
                options={"render_dpi": 150},
            )
            result = LayoutToImageConverter("png").convert(context)

            assert result.success is True
            assert len(result.artifacts) == 3

            # First artifact is primary, rest are not
            assert result.artifacts[0].is_primary is True
            assert result.artifacts[1].is_primary is False
            assert result.artifacts[2].is_primary is False

            for i, artifact in enumerate(result.artifacts):
                assert artifact.media_type == "image/png"
                assert f"_page_{i + 1:02d}" in artifact.suggested_name
                assert os.path.isfile(artifact.staging_path)
                header = Path(artifact.staging_path).read_bytes()[:8]
                assert header[:4] == b"\x89PNG"

            assert result.metrics.extra.get("page_count") == 3

    def test_render_dpi_option(self, sample_pdf_path: Path) -> None:
        """render_dpi should affect output size."""
        from docwen_plugin_layout.to_image.converter import LayoutToImageConverter

        with tempfile.TemporaryDirectory() as staging:
            context_72 = _build_fake_context(
                str(sample_pdf_path),
                staging,
                "png",
                options={"render_dpi": 72},
            )
            result_72 = LayoutToImageConverter("png").convert(context_72)
            assert result_72.success is True
            size_72 = os.path.getsize(result_72.artifacts[0].staging_path)

        with tempfile.TemporaryDirectory() as staging:
            context_300 = _build_fake_context(
                str(sample_pdf_path),
                staging,
                "png",
                options={"render_dpi": 300},
            )
            result_300 = LayoutToImageConverter("png").convert(context_300)
            assert result_300.success is True
            size_300 = os.path.getsize(result_300.artifacts[0].staging_path)

            # Higher DPI should produce larger files
            assert size_300 > size_72, f"72 DPI size={size_72}, 300 DPI size={size_300}"

    def test_direct_non_pdf_input_is_invalid(self, tmp_path: Path) -> None:
        """The internal image converter requires its preprocessed PDF input."""
        from docwen_plugin_layout.to_image.converter import LayoutToImageConverter

        xps_path = tmp_path / "test.xps"
        xps_path.write_text("fake xps content")

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(xps_path),
                staging,
                "png",
                source_format="xps",
            )
            result = LayoutToImageConverter("png").convert(context)

            assert result.success is False
            assert result.error is not None
            assert result.error.error_type == "invalid_input"
            assert result.error.diagnostic_code == "LAYOUT2IMG-INVALID-INPUT"


class TestLayoutToJpg:
    def test_single_page_pdf_to_jpg(self, sample_pdf_path: Path) -> None:
        """Single-page PDF → JPG should produce a valid JPEG artifact."""
        from docwen_plugin_layout.to_image.converter import LayoutToImageConverter

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_pdf_path),
                staging,
                "jpg",
                options={"render_dpi": 150},
            )
            result = LayoutToImageConverter("jpg").convert(context)

            assert result.success is True
            assert len(result.artifacts) == 1
            artifact = result.artifacts[0]
            assert artifact.kind == "image"
            assert artifact.media_type == "image/jpeg"
            assert artifact.is_primary is True
            assert os.path.isfile(artifact.staging_path)

            # Verify JPEG magic bytes
            header = Path(artifact.staging_path).read_bytes()[:3]
            assert header == b"\xff\xd8\xff", f"Not a valid JPEG: {header!r}"

    def test_jpg_no_alpha(self, sample_pdf_path: Path) -> None:
        """JPG output should not have alpha channel — verify via PIL."""
        from docwen_plugin_layout.to_image.converter import LayoutToImageConverter

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_pdf_path),
                staging,
                "jpg",
                options={"render_dpi": 150},
            )
            result = LayoutToImageConverter("jpg").convert(context)
            assert result.success is True

            staging_path = result.artifacts[0].staging_path
            # Verify JPEG magic bytes
            header = Path(staging_path).read_bytes()[:3]
            assert header == b"\xff\xd8\xff", f"Not a valid JPEG: {header!r}"

            # Verify no alpha channel (JPG is always RGB)
            from PIL import Image

            with Image.open(staging_path) as img:
                assert img.mode == "RGB", f"JPG should be RGB, got {img.mode}"
            assert os.path.getsize(staging_path) > 0


class TestLayoutToTif:
    def test_single_page_pdf_to_tif(self, sample_pdf_path: Path) -> None:
        """Single-page PDF → TIF should produce a valid TIFF artifact."""
        from docwen_plugin_layout.to_image.converter import LayoutToImageConverter

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_pdf_path),
                staging,
                "tif",
                options={"render_dpi": 150},
            )
            result = LayoutToImageConverter("tif").convert(context)

            assert result.success is True
            assert len(result.artifacts) == 1
            artifact = result.artifacts[0]
            assert artifact.kind == "primary"
            assert artifact.media_type == "image/tiff"
            assert artifact.is_primary is True
            assert os.path.isfile(artifact.staging_path)

            # Verify TIFF magic bytes (II or MM)
            header = Path(artifact.staging_path).read_bytes()[:2]
            assert header in (b"II", b"MM"), f"Not a valid TIFF: {header!r}"

            # Verify page count
            from PIL import Image

            with Image.open(artifact.staging_path) as img:
                assert getattr(img, "n_frames", 1) == 1

    def test_multi_page_pdf_to_tif(self, sample_multi_page_pdf_path: Path) -> None:
        """Multi-page PDF → TIF should produce a multi-page TIFF."""
        from docwen_plugin_layout.to_image.converter import LayoutToImageConverter

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_multi_page_pdf_path),
                staging,
                "tif",
                options={"render_dpi": 150},
            )
            result = LayoutToImageConverter("tif").convert(context)

            assert result.success is True
            assert len(result.artifacts) == 1
            artifact = result.artifacts[0]
            assert artifact.suggested_name == "multi.tif"
            assert os.path.isfile(artifact.staging_path)

            # Verify multi-page TIFF content
            from PIL import Image

            with Image.open(artifact.staging_path) as img:
                assert getattr(img, "n_frames", 1) == 3

            assert result.metrics.extra.get("page_count") == 3


class TestCancellation:
    def test_png_cancellation_before_execution(self, sample_pdf_path: Path) -> None:
        """A pre-cancelled context should return a cancelled result."""
        from docwen_plugin_layout.to_image.converter import LayoutToImageConverter

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(sample_pdf_path),
                staging,
                "png",
                pre_cancelled=True,
            )
            result = LayoutToImageConverter("png").convert(context)

            assert result.success is False
            assert result.error is not None
            assert result.error.error_type == "cancelled"
            assert result.error.diagnostic_code == "LAYOUT2IMG-CANCELLED"

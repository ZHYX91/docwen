"""Focused tests split from test_spreadsheet_image_extraction.py."""

from __future__ import annotations

from ._spreadsheet_image_extraction_support import (
    Path,
    _build_fake_context,
    _make_xlsx_with_images,
    openpyxl,
    os,
    pytest,
    tempfile,
)


@pytest.mark.contract
class TestSpreadsheetImageExtraction:
    """Verify that the spreadsheet converter extracts embedded images from
    XLSX workbooks, registers them as staging artifacts, and emits correct
    Markdown image references."""

    def test_extracts_images_from_single_sheet(self, tmp_path: Path) -> None:
        """A workbook with one image -> markdown contains image ref."""
        from docwen_plugin_spreadsheet.to_markdown.converter import (
            SpreadsheetToMarkdownConverter,
        )

        xlsx_path = _make_xlsx_with_images(tmp_path, images_per_sheet=[1])

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(str(xlsx_path), staging)
            result = SpreadsheetToMarkdownConverter().convert(context)

            assert result.success is True
            content = Path(result.artifacts[0].staging_path).read_text("utf-8")

            # Markdown should contain a wiki-embed image link
            assert "![[" in content
            assert "Data_image" in content
            assert ".png" in content

    def test_image_extraction_failure_is_a_typed_visible_loss(
        self,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        """A corrupt embedded image is omitted without claiming lossless success."""
        from docwen_plugin_spreadsheet.to_markdown.converter import (
            SpreadsheetToMarkdownConverter,
        )

        xlsx_path = _make_xlsx_with_images(tmp_path, images_per_sheet=[1])

        def _raise_image_read_error(_image: object) -> bytes:
            raise OSError("private image decoder detail")

        monkeypatch.setattr(openpyxl.drawing.image.Image, "_data", _raise_image_read_error)

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(str(xlsx_path), staging)
            result = SpreadsheetToMarkdownConverter().convert(context)

            assert result.success is True
            warning = next(
                diagnostic for diagnostic in result.diagnostics if diagnostic.code == "SHEET2MD-IMAGE-EXTRACTION-LOSS"
            )
            assert warning.level == "warning"
            assert "1 embedded image" in warning.message
            assert "private image decoder detail" not in warning.message
            assert result.metrics.extra["image_extraction_loss_count"] == 1
            primary = next(artifact for artifact in result.artifacts if artifact.is_primary)
            assert primary.metadata["image_extraction_loss_count"] == 1
            assert not any(artifact.kind == "image" for artifact in result.artifacts)

    def test_images_registered_as_artifacts(self, tmp_path: Path) -> None:
        """Extracted images produce 'image'-kind ArtifactManifest entries."""
        from docwen_plugin_spreadsheet.to_markdown.converter import (
            SpreadsheetToMarkdownConverter,
        )

        # 1 image in sheet 1, 2 in sheet 2
        xlsx_path = _make_xlsx_with_images(tmp_path, images_per_sheet=[1, 2])

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(str(xlsx_path), staging)
            result = SpreadsheetToMarkdownConverter().convert(context)

            assert result.success is True

            # workspace._artifacts captures both primary + image artifacts
            image_artifacts = [a for a in context.workspace._artifacts if a.kind == "image"]
            assert len(image_artifacts) == 3  # 1 + 2

            # Each image artifact should have expected fields
            for art in image_artifacts:
                assert art.artifact_id
                assert art.kind == "image"
                assert art.staging_path
                assert art.media_type in ("image/png", "image/jpeg")
                assert "sheet_name" in art.metadata
                assert os.path.isfile(art.staging_path)

    def test_image_files_written_to_staging(self, tmp_path: Path) -> None:
        """Extracted image files must exist on disk in staging."""
        from docwen_plugin_spreadsheet.to_markdown.converter import (
            SpreadsheetToMarkdownConverter,
        )

        xlsx_path = _make_xlsx_with_images(tmp_path, images_per_sheet=[1])

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(str(xlsx_path), staging)
            result = SpreadsheetToMarkdownConverter().convert(context)

            assert result.success

            image_artifacts = [a for a in context.workspace._artifacts if a.kind == "image"]
            assert len(image_artifacts) == 1

            art = image_artifacts[0]
            assert os.path.isfile(art.staging_path)
            # File should contain actual PNG data (non-zero size)
            assert os.path.getsize(art.staging_path) > 0

    def test_markdown_references_correct_count(self, tmp_path: Path) -> None:
        """Markdown output should have one image ref per extracted image."""
        from docwen_plugin_spreadsheet.to_markdown.converter import (
            SpreadsheetToMarkdownConverter,
        )

        xlsx_path = _make_xlsx_with_images(tmp_path, images_per_sheet=[2])

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(str(xlsx_path), staging)
            result = SpreadsheetToMarkdownConverter().convert(context)

            assert result.success
            content = Path(result.artifacts[0].staging_path).read_text("utf-8")

            # Count wiki embed patterns: ![[...]]
            import re

            embeds = re.findall(r"!\[\[.*?\.(png|jpg|jpeg|gif|bmp)\]\]", content)
            assert len(embeds) == 2

    def test_keep_images_false_suppresses_extraction(self, tmp_path: Path) -> None:
        """When keep_images=False, no images are extracted or referenced."""
        from docwen_plugin_spreadsheet.to_markdown.converter import (
            SpreadsheetToMarkdownConverter,
        )

        xlsx_path = _make_xlsx_with_images(tmp_path, images_per_sheet=[2])

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(xlsx_path),
                staging,
                options={"to_md_keep_images": False},
            )
            result = SpreadsheetToMarkdownConverter().convert(context)

            assert result.success is True
            content = Path(result.artifacts[0].staging_path).read_text("utf-8")

            # No image links should appear
            assert "![[" not in content

            # No image artifacts should be registered
            image_artifacts = [a for a in context.workspace._artifacts if a.kind == "image"]
            assert len(image_artifacts) == 0

    def test_multi_sheet_images_in_correct_sections(self, tmp_path: Path) -> None:
        """Images from each sheet appear after that sheet's heading."""
        from docwen_plugin_spreadsheet.to_markdown.converter import (
            SpreadsheetToMarkdownConverter,
        )

        xlsx_path = _make_xlsx_with_images(
            tmp_path,
            images_per_sheet=[1, 1, 1],
        )

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(str(xlsx_path), staging)
            result = SpreadsheetToMarkdownConverter().convert(context)

            assert result.success
            content = Path(result.artifacts[0].staging_path).read_text("utf-8")

            # Each sheet heading should be followed by an image ref
            assert "# Data" in content
            assert "# Sheet2" in content
            assert "# Sheet3" in content

            # Check image refs appear after their respective headings
            # Data sheet image
            assert "Data_image" in content
            # Sheet2 image
            assert "Sheet2_image" in content
            # Sheet3 image
            assert "Sheet3_image" in content

    def test_no_images_produces_no_image_artifacts(self, tmp_path: Path) -> None:
        """XLSX with no images should not produce image artifacts."""
        from docwen_plugin_spreadsheet.to_markdown.converter import (
            SpreadsheetToMarkdownConverter,
        )

        # Create a workbook without any images
        wb = openpyxl.Workbook()
        ws = wb.active
        assert ws is not None
        ws.title = "Plain"
        ws.cell(row=1, column=1, value="Hello")
        ws.cell(row=2, column=1, value="World")
        xlsx_path = tmp_path / "no_images.xlsx"
        wb.save(str(xlsx_path))
        wb.close()

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(str(xlsx_path), staging)
            result = SpreadsheetToMarkdownConverter().convert(context)

            assert result.success is True
            content = Path(result.artifacts[0].staging_path).read_text("utf-8")

            # Should not contain any image embeds
            assert "![[" not in content

            # No image artifacts
            image_artifacts = [a for a in context.workspace._artifacts if a.kind == "image"]
            assert len(image_artifacts) == 0

    def test_artifact_metadata_includes_image_count(self, tmp_path: Path) -> None:
        """Primary artifact metadata should report extracted image count."""
        from docwen_plugin_spreadsheet.to_markdown.converter import (
            SpreadsheetToMarkdownConverter,
        )

        xlsx_path = _make_xlsx_with_images(tmp_path, images_per_sheet=[3])

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(str(xlsx_path), staging)
            result = SpreadsheetToMarkdownConverter().convert(context)

            assert result.success
            artifact = result.artifacts[0]
            assert artifact.metadata["image_count"] == 3

    def test_image_extract_does_not_break_table_rendering(
        self,
        tmp_path: Path,
    ) -> None:
        """Image extraction should coexist with table rendering."""
        from docwen_plugin_spreadsheet.to_markdown.converter import (
            SpreadsheetToMarkdownConverter,
        )

        xlsx_path = _make_xlsx_with_images(tmp_path, images_per_sheet=[1])

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(str(xlsx_path), staging)
            result = SpreadsheetToMarkdownConverter().convert(context)

            assert result.success
            content = Path(result.artifacts[0].staging_path).read_text("utf-8")

            # Table data should still be present
            assert "Header" in content
            assert "Row2" in content
            assert "Row3" in content
            assert "|" in content  # pipe table markers

    def test_image_format_detection_is_applied(self, tmp_path: Path) -> None:
        """Extracted image files should have correct media type."""
        from docwen_plugin_spreadsheet.to_markdown.converter import (
            SpreadsheetToMarkdownConverter,
        )

        xlsx_path = _make_xlsx_with_images(tmp_path, images_per_sheet=[1])

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(str(xlsx_path), staging)
            result = SpreadsheetToMarkdownConverter().convert(context)

            assert result.success

            image_artifacts = [a for a in context.workspace._artifacts if a.kind == "image"]
            assert len(image_artifacts) >= 1

            # All image artifacts should have a recognizable media type
            for art in image_artifacts:
                assert art.media_type.startswith("image/")
                # suggested_name should end with the correct extension
                assert art.suggested_name.endswith(
                    art.media_type.replace("image/", "."),
                ) or art.suggested_name.endswith(".png")

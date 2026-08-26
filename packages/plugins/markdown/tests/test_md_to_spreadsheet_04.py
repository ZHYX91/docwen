"""Focused tests split from test_md_to_spreadsheet.py."""

from __future__ import annotations

from ._md_to_spreadsheet_support import (
    SAMPLE_MD_TABLES,
    MdToCsvConverter,
    Path,
    _bundled_spreadsheet_template_id,
    _execute_markdown_runtime,
    _old_system_spreadsheet_smoke_fixture,
    _png_bytes,
    make_context,
    pytest,
    write_temp_md,
)


class TestMdToCsv:
    """Tests for MD → CSV conversion."""

    @pytest.mark.contract
    def test_canonical_english_sheet_template_exports_old_project_smoke_csvs(self, tmp_path: Path):
        """The bundled XLSX template keeps the old Tk/PySide6 CSV sheet chain."""
        import csv

        project_root = Path(__file__).resolve().parents[4]
        fixture = _old_system_spreadsheet_smoke_fixture()
        template = project_root / "templates" / "English Sample Sheet Template.xlsx"
        md_path = tmp_path / "canonical-template-smoke.md"
        md_path.write_text(fixture["input_markdown"], encoding="utf-8")
        ctx, _workspace = make_context(
            str(md_path),
            target_format="csv",
            options={"template_name": str(template)},
        )

        converter = MdToCsvConverter()
        result = converter.convert(ctx)

        assert result.success
        assert [artifact.suggested_name for artifact in result.artifacts] == fixture["csv_suggested_names"]

        with Path(result.artifacts[0].staging_path).open(encoding="utf-8-sig", newline="") as handle:
            rows = list(csv.reader(handle))
        assert rows[:11] == fixture["csv_sheet1_rows"]

        for artifact in result.artifacts[1:]:
            with Path(artifact.staging_path).open(encoding="utf-8-sig", newline="") as handle:
                assert list(csv.reader(handle)) == []

    @pytest.mark.integration
    def test_markdown_csv_old_system_fixture_finalizes_through_runtime(self, tmp_path: Path):
        """Runtime finalizer places the old-system MD->CSV smoke output chain."""
        import csv
        import os

        from docwen_runtime.path_io import filesystem_path

        fixture = _old_system_spreadsheet_smoke_fixture()
        md_path = tmp_path / "canonical-template-smoke.md"
        md_path.write_text(fixture["input_markdown"], encoding="utf-8")
        output_dir = tmp_path / "final-output"

        result = _execute_markdown_runtime(
            md_path,
            "csv",
            output_dir,
            options={"template_name": _bundled_spreadsheet_template_id()},
        )

        assert result.success
        assert any(diagnostic.code == "FINALIZER_DONE" for diagnostic in result.diagnostics)
        assert [artifact.suggested_name for artifact in result.artifacts] == [
            os.path.normpath(name) for name in fixture["csv_suggested_names"]
        ]
        assert len(result.artifacts) == len(fixture["csv_suggested_names"])

        final_paths = [Path(artifact.staging_path) for artifact in result.artifacts]
        for path in final_paths:
            assert filesystem_path(path).exists()
            assert path.suffix == ".csv"
            assert path.is_relative_to(output_dir)
            assert not path.is_relative_to(tmp_path / "workspace")

        with filesystem_path(final_paths[0]).open(encoding="utf-8-sig", newline="") as handle:
            rows = list(csv.reader(handle))
        assert rows[:11] == fixture["csv_sheet1_rows"]

        for path in final_paths[1:]:
            with filesystem_path(path).open(encoding="utf-8-sig", newline="") as handle:
                assert list(csv.reader(handle)) == []

    @pytest.mark.contract
    def test_converts_tables_to_csv(self):
        """Markdown tables are converted to CSV files."""
        md_path = write_temp_md(SAMPLE_MD_TABLES)
        ctx, _workspace = make_context(md_path, target_format="csv")

        converter = MdToCsvConverter()
        result = converter.convert(ctx)

        assert result.success, f"Conversion failed: {result.error.message if result.error else 'unknown'}"
        # Two tables should produce two CSV files
        assert len(result.artifacts) == 2
        for artifact in result.artifacts:
            assert artifact.kind == "primary"
            assert artifact.media_type == "text/csv"
            assert artifact.suggested_name.endswith(".csv")

        # Verify first CSV file content
        import csv

        output_path = Path(result.artifacts[0].staging_path)
        assert output_path.exists()
        with open(output_path, encoding="utf-8") as f:
            reader = csv.reader(f)
            rows = list(reader)
        assert rows[0] == ["Name", "Age", "City"]
        assert rows[1] == ["Alice", "30", "Beijing"]
        assert rows[2] == ["Bob", "25", "Shanghai"]

        # Verify second CSV file content
        output_path2 = Path(result.artifacts[1].staging_path)
        assert output_path2.exists()
        with open(output_path2, encoding="utf-8") as f:
            reader = csv.reader(f)
            rows2 = list(reader)
        assert rows2[0] == ["Product", "Price", "Qty"]

    @pytest.mark.contract
    def test_single_table_csv(self):
        """Single table Markdown produces one CSV file."""
        md_content = "# Table\n\n| A | B |\n|---|---|\n| 1 | 2 |\n"
        md_path = write_temp_md(md_content)
        ctx, _workspace = make_context(md_path, target_format="csv")

        converter = MdToCsvConverter()
        result = converter.convert(ctx)

        assert result.success
        assert len(result.artifacts) == 1
        artifact = result.artifacts[0]
        assert artifact.suggested_name.endswith(".csv")

    @pytest.mark.contract
    def test_csv_embed_images_downgrade_to_filenames_without_internal_placeholders(
        self,
        tmp_path: Path,
    ) -> None:
        """CSV exposes image filenames because it has no binary drawing channel."""
        import csv

        image_bytes = _png_bytes()
        (tmp_path / "wiki.png").write_bytes(image_bytes)
        (tmp_path / "markdown.png").write_bytes(image_bytes)
        source = tmp_path / "csv-images.md"
        source.write_text(
            "| Wiki | Markdown |\n| --- | --- |\n| ![[wiki.png]] | ![Markdown alt](markdown.png) |\n",
            encoding="utf-8",
        )
        context, _workspace = make_context(
            str(source),
            target_format="csv",
            config_values={
                "link": {
                    "embed_links": {
                        "wiki_image_mode": "embed",
                        "markdown_image_mode": "embed",
                        "md_file_mode": "embed",
                    }
                }
            },
        )

        result = MdToCsvConverter().convert(context)

        assert result.success is True, result.error
        output = Path(result.artifacts[0].staging_path)
        with output.open(encoding="utf-8-sig", newline="") as handle:
            rows = list(csv.reader(handle))
        assert rows[1] == ["wiki.png", "markdown.png"]
        serialized = output.read_text(encoding="utf-8-sig")
        assert "{{IMAGE:" not in serialized
        assert "![[" not in serialized
        assert "![" not in serialized

    @pytest.mark.contract
    def test_no_table_md_fails(self):
        """MD with no tables fails with clear error."""
        md_content = "# No tables\n\nJust text."
        md_path = write_temp_md(md_content)
        ctx, _workspace = make_context(md_path, target_format="csv")

        converter = MdToCsvConverter()
        result = converter.convert(ctx)

        assert not result.success
        assert result.error is not None
        assert result.error.error_type == "conversion_failed"
        assert "no tables" in result.error.message.lower()

    @pytest.mark.contract
    def test_csv_artifact_metadata(self):
        """CSV artifact metadata includes row and header counts."""
        md_path = write_temp_md(SAMPLE_MD_TABLES)
        ctx, _workspace = make_context(md_path, target_format="csv")

        converter = MdToCsvConverter()
        result = converter.convert(ctx)

        assert result.success
        artifact = result.artifacts[0]
        assert artifact.metadata.get("row_count", 0) >= 2
        assert artifact.metadata.get("header_count", 0) == 3

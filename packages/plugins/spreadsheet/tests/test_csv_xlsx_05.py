"""Focused tests split from test_csv_xlsx.py."""

from __future__ import annotations

from typing import ClassVar

from ._csv_xlsx_support import (
    Any,
    Path,
    _build_fake_context,
    _write_multisheet_marker_workbook,
    _write_workbook,
    pytest,
    tempfile,
)

pytestmark = [pytest.mark.golden, pytest.mark.contract]


class TestSmartSheetConverter:
    """ROUTE-SHEETFMT-* routes are backed by the external office bridge."""

    ALL_SMARTSHEET_PAIRS: ClassVar[list[tuple[str, str]]] = [
        ("xlsx", "xls"),
        ("xlsx", "ods"),
        ("xlsx", "et"),
        ("xls", "xlsx"),
        ("ods", "xlsx"),
        ("et", "xlsx"),
        ("xls", "ods"),
        ("xls", "et"),
        ("ods", "xls"),
        ("ods", "et"),
        ("et", "xls"),
        ("et", "ods"),
        ("csv", "xls"),
        ("csv", "ods"),
        ("xls", "csv"),
        ("ods", "csv"),
        ("et", "csv"),
    ]

    def _write_input_for_source(self, tmp_path: Path, source_format: str, sample_xlsx_path: Path) -> Path:
        if source_format == "xlsx":
            return sample_xlsx_path
        input_path = tmp_path / f"smart-source.{source_format}"
        if source_format == "csv":
            input_path.write_text("Name,Value\nAlpha,1\nBeta,2\n", encoding="utf-8")
        else:
            input_path.write_bytes(f"fake {source_format} workbook".encode("ascii"))
        return input_path

    @pytest.mark.parametrize("target_format", ["xls", "ods"])
    def test_csv_to_binary_preserves_numeric_cells_like_old_systems(
        self,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
        target_format: str,
    ) -> None:
        """The private CSV hub must use the canonical numeric-cell semantics."""
        import openpyxl

        from docwen_core.office_bridge import BridgeResult
        from docwen_plugin_spreadsheet.format_conversion import converter as sheet_module
        from docwen_plugin_spreadsheet.format_conversion.converter import SmartSheetConverter

        captured_rows: list[list[list[Any]]] = []

        def fake_convert(input_path: str, output_path: str, **kwargs: Any) -> BridgeResult:
            wb = openpyxl.load_workbook(input_path, data_only=True)
            ws = wb.active
            assert ws is not None
            captured_rows.append([list(row) for row in ws.iter_rows(values_only=True)])
            wb.close()
            Path(output_path).write_bytes(b"converted")
            return BridgeResult(True, output_path=output_path, backend="fake-office")

        monkeypatch.setattr(sheet_module, "convert_with_backend_priority", fake_convert)

        source_path = tmp_path / "typed.csv"
        source_path.write_text("Name,Value\nAlpha,11\nBeta,22\n", encoding="utf-8")
        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(str(source_path), staging, target_format=target_format)
            context.request.input_refs[0] = type(context.request.input_refs[0])(
                path=str(source_path), format="csv", category="spreadsheet"
            )
            result = SmartSheetConverter().convert(context)

        assert result.success is True
        assert captured_rows == [
            [
                ["Name", "Value"],
                ["Alpha", 11],
                ["Beta", 22],
            ]
        ]
        assert result.artifacts[0].metadata["backend"] == "openpyxl -> fake-office"

    def test_binary_to_csv_pipeline_finalizes_every_sheet(
        self,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        """Plugin routing and the runtime finalizer must keep every CSV artifact."""
        import csv

        from docwen_core.models.file_ref import FileRef
        from docwen_core.models.request import ConversionRequest, OutputPolicy
        from docwen_core.office_bridge import BridgeResult
        from docwen_plugin_spreadsheet.format_conversion import converter as sheet_module
        from docwen_plugin_spreadsheet.plugin import SpreadsheetPlugin
        from docwen_runtime.engine.route_resolver import RouteResolver
        from docwen_runtime.engine.task_manager import TaskManager
        from docwen_runtime.output.finalizer import OutputFinalizer
        from docwen_runtime.plugin_registry.registry import PluginRegistry
        from docwen_runtime.workspace.manager import WorkspaceManager

        def fake_convert(input_path: str, output_path: str, **kwargs: Any) -> BridgeResult:
            _write_multisheet_marker_workbook(Path(output_path))
            return BridgeResult(True, output_path=output_path, backend="fake-office")

        monkeypatch.setattr(sheet_module, "convert_with_backend_priority", fake_convert)

        source_path = tmp_path / "legacy-multi.xls"
        source_path.write_bytes(b"fake xls workbook")
        output_dir = tmp_path / "out"
        workspace_root = tmp_path / "workspace"
        registry = PluginRegistry()
        registry.register(SpreadsheetPlugin())
        task_manager = TaskManager(
            plugin_registry=registry,
            route_resolver=RouteResolver(registry),
            workspace_manager=WorkspaceManager(str(workspace_root)),
            output_finalizer=OutputFinalizer(),
        )
        request = ConversionRequest(
            request_id="smart-csv-finalize",
            input_refs=[
                FileRef(
                    path=str(source_path),
                    format="xls",
                    category="spreadsheet",
                    size_bytes=source_path.stat().st_size,
                )
            ],
            target_format="csv",
            output_policy=OutputPolicy(output_dir=str(output_dir)),
        )

        result = task_manager.execute_single(request)

        assert result.success is True
        assert [Path(artifact.staging_path).name for artifact in result.artifacts] == [
            "legacy-multi_Alpha_Data.csv",
            "legacy-multi_Beta-Summary.csv",
            "legacy-multi_Hidden_Sheet.csv",
        ]
        assert [artifact.is_primary for artifact in result.artifacts] == [True, False, False]
        assert all(Path(artifact.staging_path).parent == output_dir for artifact in result.artifacts)
        assert all(artifact.metadata["sheet_count"] == 3 for artifact in result.artifacts)
        assert all(artifact.metadata["source_format"] == "xls" for artifact in result.artifacts)
        assert all(artifact.metadata["target_format"] == "csv" for artifact in result.artifacts)
        assert all(artifact.metadata["backend"] == "fake-office" for artifact in result.artifacts)

        markers: list[str] = []
        for artifact in result.artifacts:
            with open(artifact.staging_path, encoding="utf-8-sig", newline="") as handle:
                markers.append(next(csv.reader(handle))[0])
        assert markers == ["ALPHA_MARKER", "BETA_MARKER", "HIDDEN_MARKER"]
        assert {diagnostic.code for diagnostic in result.diagnostics} >= {"SHEETFMT-OK", "FINALIZER_DONE"}
        assert result.metrics.input_bytes == source_path.stat().st_size
        assert result.metrics.output_bytes == sum(
            Path(artifact.staging_path).stat().st_size for artifact in result.artifacts
        )
        assert result.metrics.extra == {
            "backend": "fake-office",
            "sheet_count": 3,
            "total_rows": 3,
            "output_dir": str(output_dir),
        }
        assert not list(workspace_root.rglob("*.xlsx"))
        assert not list(workspace_root.rglob("*.csv"))

    def test_smart_converter_succeeds_when_bridge_available(
        self, sample_xlsx_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """xlsx→xls should succeed when the bridge can produce the target file."""
        from docwen_core.office_bridge import BridgeResult
        from docwen_plugin_spreadsheet.format_conversion import converter as sheet_module
        from docwen_plugin_spreadsheet.format_conversion.converter import SmartSheetConverter

        def fake_convert(input_path: str, output_path: str, **kwargs: Any) -> BridgeResult:
            Path(output_path).write_bytes(b"converted")
            return BridgeResult(True, output_path=output_path, backend="fake-office")

        monkeypatch.setattr(sheet_module, "convert_with_backend_priority", fake_convert)

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(str(sample_xlsx_path), staging, target_format="xls")
            context.request.input_refs[0] = type(context.request.input_refs[0])(
                path=str(sample_xlsx_path), format="xlsx", category="spreadsheet"
            )
            result = SmartSheetConverter().convert(context)

            assert result.success is True
            assert result.error is None
            assert result.artifacts[0].suggested_name.endswith(".xls")
            assert result.diagnostics[0].code == "SHEETFMT-OK"

    @pytest.mark.parametrize(("source_format", "target_format"), [("xlsx", "ods"), ("ods", "xlsx")])
    def test_ods_routes_skip_wps_spreadsheets_candidate(
        self,
        sample_xlsx_path: Path,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
        source_format: str,
        target_format: str,
    ) -> None:
        """ODS source/target routes should follow the legacy Excel/LibreOffice policy."""
        from docwen_core.office_bridge import BridgeResult
        from docwen_plugin_spreadsheet.format_conversion import converter as sheet_module
        from docwen_plugin_spreadsheet.format_conversion.converter import SmartSheetConverter

        captured: list[tuple[list[tuple[str, str]], str | None]] = []

        def fake_convert(input_path: str, output_path: str, **kwargs: Any) -> BridgeResult:
            candidates = kwargs["com_candidates"].values()
            captured.append(
                (
                    [(candidate.name, candidate.prog_id) for candidate in candidates],
                    kwargs["libreoffice_format"],
                )
            )
            if Path(output_path).suffix.lower() == ".xlsx":
                _write_workbook(Path(output_path), [["Name", "Value"], ["Alpha", 1]])
            else:
                Path(output_path).write_bytes(b"converted")
            return BridgeResult(True, output_path=output_path, backend="fake-office")

        monkeypatch.setattr(sheet_module, "convert_with_backend_priority", fake_convert)

        source_path = sample_xlsx_path
        if source_format == "ods":
            source_path = tmp_path / "smart-source.ods"
            source_path.write_bytes(b"fake ods workbook")

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(str(source_path), staging, target_format=target_format)
            context.request.input_refs[0] = type(context.request.input_refs[0])(
                path=str(source_path), format=source_format, category="spreadsheet"
            )
            result = SmartSheetConverter().convert(context)

        assert result.success is True
        assert captured == [
            (
                [("Microsoft Excel", "Excel.Application")],
                target_format,
            )
        ]

    def test_non_ods_bridge_routes_keep_wps_then_excel_candidate_order(
        self, sample_xlsx_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """Non-ODS binary spreadsheet routes should preserve WPS before Excel."""
        from docwen_core.office_bridge import BridgeResult
        from docwen_plugin_spreadsheet.format_conversion import converter as sheet_module
        from docwen_plugin_spreadsheet.format_conversion.converter import SmartSheetConverter

        captured: list[list[tuple[str, str]]] = []

        def fake_convert(input_path: str, output_path: str, **kwargs: Any) -> BridgeResult:
            candidates = kwargs["com_candidates"].values()
            captured.append([(candidate.name, candidate.prog_id) for candidate in candidates])
            Path(output_path).write_bytes(b"converted")
            return BridgeResult(True, output_path=output_path, backend="fake-office")

        monkeypatch.setattr(sheet_module, "convert_with_backend_priority", fake_convert)

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(str(sample_xlsx_path), staging, target_format="xls")
            context.request.input_refs[0] = type(context.request.input_refs[0])(
                path=str(sample_xlsx_path), format="xlsx", category="spreadsheet"
            )
            result = SmartSheetConverter().convert(context)

        assert result.success is True
        assert captured == [
            [
                ("WPS Spreadsheets", "Ket.Application"),
                ("Microsoft Excel", "Excel.Application"),
            ]
        ]

    def test_smart_converter_surfaces_missing_backend_cleanly(
        self, sample_xlsx_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """xlsx→xls should report dependency_missing when no bridge backend is available."""
        from docwen_core.office_bridge import BridgeResult
        from docwen_plugin_spreadsheet.format_conversion import converter as sheet_module
        from docwen_plugin_spreadsheet.format_conversion.converter import SmartSheetConverter

        def fake_convert(input_path: str, output_path: str, **kwargs: Any) -> BridgeResult:
            return BridgeResult(False, message="Install LibreOffice.")

        monkeypatch.setattr(sheet_module, "convert_with_backend_priority", fake_convert)

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(str(sample_xlsx_path), staging, target_format="xls")
            context.request.input_refs[0] = type(context.request.input_refs[0])(
                path=str(sample_xlsx_path), format="xlsx", category="spreadsheet"
            )
            result = SmartSheetConverter().convert(context)

            assert result.success is False
            assert result.error is not None
            assert result.error.error_type == "dependency_missing"
            assert result.error.diagnostic_code == "SHEETFMT-BACKEND"

    def test_xlsx_ods_surfaces_installed_backend_failure_distinctly(
        self, sample_xlsx_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """An installed backend conversion failure is not a missing dependency."""
        from docwen_core.office_bridge import BridgeResult
        from docwen_plugin_spreadsheet.format_conversion import converter as sheet_module
        from docwen_plugin_spreadsheet.format_conversion.converter import SmartSheetConverter

        def fake_convert(input_path: str, output_path: str, **kwargs: Any) -> BridgeResult:
            return BridgeResult(
                False,
                message="Microsoft Excel SaveAs failed.",
                attempted_backend_ids=("msoffice_excel",),
                available_backend_ids=("msoffice_excel",),
            )

        monkeypatch.setattr(sheet_module, "convert_with_backend_priority", fake_convert)

        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(str(sample_xlsx_path), staging, target_format="ods")
            result = SmartSheetConverter().convert(context)

        assert result.success is False
        assert result.error is not None
        assert result.error.error_type == "conversion_failed"
        assert result.error.diagnostic_code == "SHEETFMT-BACKEND-FAILED"

"""Focused tests split from test_csv_xlsx.py."""

from __future__ import annotations

from typing import ClassVar

from ._csv_xlsx_support import (
    Any,
    Path,
    _add_policy02_workbook_protection,
    _build_fake_context,
    _write_multisheet_marker_workbook,
    _write_policy02_external_link_workbook,
    _write_policy02_ods,
    _write_policy02_protected_workbook,
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

    def test_policy02_delivers_ods_through_legacy_xls_after_direct_backend_failure(
        self,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        from docwen_core.office_bridge import BridgeResult
        from docwen_plugin_spreadsheet.format_conversion import converter as sheet_module
        from docwen_plugin_spreadsheet.format_conversion.converter import SmartSheetConverter
        from docwen_plugin_spreadsheet.format_conversion.xlsx_ods_policy import inspect_xlsx_ods_policy

        source = tmp_path / "owner.xlsx"
        _write_policy02_external_link_workbook(source)
        source_before = source.read_bytes()
        calls: list[tuple[Path, Path]] = []

        def fake_convert(input_path: str, output_path: str, **kwargs: Any) -> BridgeResult:
            bridge_input = Path(input_path)
            bridge_output = Path(output_path)
            calls.append((bridge_input, bridge_output))
            if len(calls) == 1:
                assert bridge_output.suffix == ".ods"
                inspection = inspect_xlsx_ods_policy(bridge_input)
                assert inspection.external_formula_cells == ()
                assert inspection.external_link_parts_present is False
                return BridgeResult(
                    False,
                    message="Direct XLSX to ODS failed.",
                    attempted_backend_ids=("msoffice_excel", "libreoffice"),
                    available_backend_ids=("msoffice_excel", "libreoffice"),
                )
            if len(calls) == 2:
                assert bridge_output.suffix == ".xls"
                assert bridge_input == calls[0][0]
                bridge_output.write_bytes(b"legacy xls")
                return BridgeResult(True, output_path=str(bridge_output), backend="fake-xls")
            if len(calls) == 3:
                assert bridge_input.suffix == ".xls"
                assert bridge_output.suffix == ".xlsx"
                _write_workbook(bridge_output, [["External cached value"], [30]])
                return BridgeResult(True, output_path=str(bridge_output), backend="fake-xlsx")
            assert len(calls) == 4
            assert bridge_input.suffix == ".xlsx"
            assert bridge_output.suffix == ".ods"
            _write_policy02_ods(bridge_output, value="30")
            return BridgeResult(True, output_path=str(bridge_output), backend="fake-ods")

        monkeypatch.setattr(sheet_module, "convert_with_backend_priority", fake_convert)
        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(str(source), staging, target_format="ods")
            result = SmartSheetConverter().convert(context)

        assert result.success is True
        assert source.read_bytes() == source_before
        assert len(calls) == 4
        assert result.artifacts[0].metadata["backend"] == "fake-xls -> fake-xlsx -> fake-ods"
        assert [diagnostic.code for diagnostic in result.diagnostics] == [
            "SHEETFMT-OK",
            "EXTERNAL_LINK_FLATTENED",
            "ODS_LEGACY_XLS_FALLBACK",
        ]
        fallback_warning = result.diagnostics[-1].message
        assert "Direct XLSX-to-ODS conversion failed" in fallback_warning
        assert "legacy XLS" in fallback_warning
        assert "conditional formatting" in fallback_warning
        assert "protection" in fallback_warning

    def test_policy02_composes_external_flattening_and_protection_removal_once(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        from docwen_core.office_bridge import BridgeResult
        from docwen_plugin_spreadsheet.format_conversion import converter as sheet_module
        from docwen_plugin_spreadsheet.format_conversion.converter import SmartSheetConverter
        from docwen_plugin_spreadsheet.format_conversion.xlsx_ods_policy import inspect_xlsx_ods_policy

        source = tmp_path / "linked-protected.xlsx"
        _write_policy02_external_link_workbook(source)
        _add_policy02_workbook_protection(source, "test")
        source_before = source.read_bytes()
        backend_inputs: list[str] = []

        def fake_convert(input_path: str, output_path: str, **kwargs: Any) -> BridgeResult:
            backend_inputs.append(input_path)
            inspection = inspect_xlsx_ods_policy(input_path)
            assert inspection.external_formula_cells == ()
            assert inspection.password_protected_elements == ()
            _write_policy02_ods(Path(output_path), value="30")
            return BridgeResult(True, output_path=output_path, backend="fake-office")

        monkeypatch.setattr(sheet_module, "convert_with_backend_priority", fake_convert)
        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(source),
                staging,
                target_format="ods",
                options={
                    "spreadsheet_password": "test",
                    "allow_spreadsheet_protection_loss": True,
                },
            )
            result = SmartSheetConverter().convert(context)

        assert result.success is True
        assert len(backend_inputs) == 1
        assert source.read_bytes() == source_before
        assert [diagnostic.code for diagnostic in result.diagnostics] == [
            "SHEETFMT-OK",
            "EXTERNAL_LINK_FLATTENED",
            "PROTECTION_REMOVED_FOR_TARGET",
        ]

    @pytest.mark.parametrize(("scope", "password"), [("workbook", "test"), ("sheet", "pwd")])
    def test_policy02_converter_removes_protection_only_with_password_and_consent(
        self,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
        scope: str,
        password: str,
    ) -> None:
        from docwen_core.office_bridge import BridgeResult
        from docwen_plugin_spreadsheet.format_conversion import converter as sheet_module
        from docwen_plugin_spreadsheet.format_conversion.converter import SmartSheetConverter
        from docwen_plugin_spreadsheet.format_conversion.xlsx_ods_policy import inspect_xlsx_ods_policy

        source = tmp_path / f"{scope}-protected.xlsx"
        _write_policy02_protected_workbook(source, scope=scope, password=password)
        source_before = source.read_bytes()

        def fake_convert(input_path: str, output_path: str, **kwargs: Any) -> BridgeResult:
            inspection = inspect_xlsx_ods_policy(input_path)
            assert inspection.password_protected_elements == ()
            assert inspection.unpassworded_protected_elements == ()
            _write_policy02_ods(Path(output_path))
            return BridgeResult(True, output_path=output_path, backend="fake-office")

        monkeypatch.setattr(sheet_module, "convert_with_backend_priority", fake_convert)
        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(source),
                staging,
                target_format="ods",
                options={
                    "spreadsheet_password": password,
                    "allow_spreadsheet_protection_loss": True,
                },
            )
            result = SmartSheetConverter().convert(context)

        assert result.success is True
        assert source.read_bytes() == source_before
        assert [diagnostic.code for diagnostic in result.diagnostics] == [
            "SHEETFMT-OK",
            "PROTECTION_REMOVED_FOR_TARGET",
        ]

    def test_policy02_no_password_protection_control_is_preserved_without_warning(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        from docwen_core.office_bridge import BridgeResult
        from docwen_plugin_spreadsheet.format_conversion import converter as sheet_module
        from docwen_plugin_spreadsheet.format_conversion.converter import SmartSheetConverter

        source = tmp_path / "all-locked.xlsx"
        _write_policy02_protected_workbook(source, scope="sheet", password=None)

        def fake_convert(input_path: str, output_path: str, **kwargs: Any) -> BridgeResult:
            from docwen_plugin_spreadsheet.format_conversion.xlsx_ods_policy import inspect_xlsx_ods_policy

            inspection = inspect_xlsx_ods_policy(input_path)
            assert inspection.password_protected_elements == ()
            assert inspection.unpassworded_protected_elements == ("Sheet",)
            Path(output_path).write_bytes(b"converted ods")
            return BridgeResult(True, output_path=output_path, backend="fake-office")

        monkeypatch.setattr(sheet_module, "convert_with_backend_priority", fake_convert)
        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(source),
                staging,
                target_format="ods",
                options={"spreadsheet_password": "unused", "allow_spreadsheet_protection_loss": True},
            )
            result = SmartSheetConverter().convert(context)

        assert result.success is True
        assert [diagnostic.code for diagnostic in result.diagnostics] == ["SHEETFMT-OK"]

    def test_policy02_final_ods_validator_streams_content_xml(
        self,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        from docwen_plugin_spreadsheet.format_conversion import xlsx_ods_policy
        from docwen_plugin_spreadsheet.format_conversion.xlsx_ods_policy import (
            XlsxOdsPreparation,
            validate_prepared_ods,
        )

        output = tmp_path / "streamed.ods"
        _write_policy02_ods(output, value="30")
        preparation = XlsxOdsPreparation(
            output_path=str(output),
            external_links_flattened=True,
            protection_removed=False,
            flattened_cached_values=("30",),
        )

        def forbidden_whole_tree_parse(payload: bytes) -> Any:
            raise AssertionError(f"validator must stream content.xml, got {len(payload)} bytes")

        monkeypatch.setattr(xlsx_ods_policy, "_parse_xml", forbidden_whole_tree_parse)

        validate_prepared_ods(output, preparation)

    def test_policy02_final_ods_validator_rejects_lost_cached_value(self, tmp_path: Path) -> None:
        from docwen_plugin_spreadsheet.format_conversion.xlsx_ods_policy import (
            XlsxOdsPolicyError,
            XlsxOdsPreparation,
            validate_prepared_ods,
        )

        output = tmp_path / "bad.ods"
        _write_policy02_ods(output, value="29")
        preparation = XlsxOdsPreparation(
            output_path=str(output),
            external_links_flattened=True,
            protection_removed=False,
            flattened_cached_values=("30",),
        )

        with pytest.raises(XlsxOdsPolicyError) as caught:
            validate_prepared_ods(output, preparation)

        assert caught.value.diagnostic_code == "EXTERNAL_LINK_FLATTENING_NOT_DELIVERED"

    def test_policy02_final_ods_validator_does_not_attribute_unrelated_ref_to_name_removal(
        self,
        tmp_path: Path,
    ) -> None:
        from docwen_plugin_spreadsheet.format_conversion.xlsx_ods_policy import (
            XlsxOdsPreparation,
            validate_prepared_ods,
        )

        output = tmp_path / "best-effort.ods"
        _write_policy02_ods(output, value="#REF!")
        preparation = XlsxOdsPreparation(
            output_path=str(output),
            external_links_flattened=True,
            protection_removed=False,
            removed_external_defined_names=("LegacyExternalName",),
        )

        validate_prepared_ods(output, preparation)

    def test_policy02_final_ods_validator_rejects_retained_named_sheet_protection(self, tmp_path: Path) -> None:
        from docwen_plugin_spreadsheet.format_conversion.xlsx_ods_policy import (
            XlsxOdsPolicyError,
            XlsxOdsPreparation,
            validate_prepared_ods,
        )

        output = tmp_path / "bad-protected.ods"
        _write_policy02_ods(output, protected_sheet="Sheet")
        preparation = XlsxOdsPreparation(
            output_path=str(output),
            external_links_flattened=False,
            protection_removed=True,
            removed_protection_elements=("Sheet",),
        )

        with pytest.raises(XlsxOdsPolicyError) as caught:
            validate_prepared_ods(output, preparation)

        assert caught.value.diagnostic_code == "PROTECTION_REMOVAL_NOT_DELIVERED"

    def test_ods_to_xls_resolves_priority_per_bridge_leg(self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
        from docwen_core.office_bridge import BridgeResult
        from docwen_plugin_spreadsheet.format_conversion import converter as sheet_module
        from docwen_plugin_spreadsheet.format_conversion.converter import SmartSheetConverter

        calls: list[dict[str, Any]] = []

        def fake_convert(input_path: str, output_path: str, **kwargs: Any) -> BridgeResult:
            calls.append(kwargs)
            if Path(output_path).suffix.lower() == ".xlsx":
                _write_workbook(Path(output_path), [["Name", "Value"], ["Alpha", 1]])
            else:
                Path(output_path).write_bytes(b"converted")
            return BridgeResult(True, output_path=output_path, backend="fake-office")

        monkeypatch.setattr(sheet_module, "convert_with_backend_priority", fake_convert)

        source_path = tmp_path / "smart-source.ods"
        source_path.write_bytes(b"fake ods workbook")
        config_snapshot = {
            "software": {
                "default_priority": {"spreadsheet_processors": ["wps_spreadsheets", "libreoffice", "msoffice_excel"]},
                "special_conversions": {"ods": ["msoffice_excel", "libreoffice"]},
            }
        }
        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(
                str(source_path), staging, target_format="xls", config_snapshot=config_snapshot
            )
            result = SmartSheetConverter().convert(context)

        assert result.success is True
        assert [call["backend_priority"] for call in calls] == [
            ["msoffice_excel", "libreoffice"],
            ["wps_spreadsheets", "libreoffice", "msoffice_excel"],
        ]
        assert [set(call["com_candidates"]) for call in calls] == [
            {"msoffice_excel"},
            {"wps_spreadsheets", "msoffice_excel"},
        ]

    @pytest.mark.parametrize(("source_format", "target_format"), ALL_SMARTSHEET_PAIRS)
    def test_all_declared_smart_sheet_routes_succeed_when_bridge_available(
        self,
        sample_xlsx_path: Path,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
        source_format: str,
        target_format: str,
    ) -> None:
        """Every declared SmartSheet route should reach the converter contract."""
        from docwen_core.office_bridge import BridgeResult
        from docwen_plugin_spreadsheet.format_conversion import converter as sheet_module
        from docwen_plugin_spreadsheet.format_conversion.converter import SmartSheetConverter

        bridge_calls: list[tuple[str, str, str]] = []

        def fake_convert(input_path: str, output_path: str, **kwargs: Any) -> BridgeResult:
            bridge_calls.append(
                (
                    Path(input_path).suffix.lower(),
                    Path(output_path).suffix.lower(),
                    kwargs["source_format"],
                )
            )
            if Path(output_path).suffix.lower() == ".xlsx":
                _write_workbook(Path(output_path), [["Name", "Value"], ["Alpha", 1], ["Beta", 2]])
            else:
                Path(output_path).write_bytes(b"converted")
            return BridgeResult(True, output_path=output_path, backend="fake-office")

        monkeypatch.setattr(sheet_module, "convert_with_backend_priority", fake_convert)

        source_path = self._write_input_for_source(tmp_path, source_format, sample_xlsx_path)
        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(str(source_path), staging, target_format=target_format)
            context.request.input_refs[0] = type(context.request.input_refs[0])(
                path=str(source_path), format=source_format, category="spreadsheet"
            )
            result = SmartSheetConverter().convert(context)

        assert result.success is True
        assert result.error is None
        assert result.artifacts
        assert result.artifacts[0].suggested_name.endswith(f".{target_format}")
        assert result.artifacts[0].metadata["source_format"] == source_format
        assert result.artifacts[0].metadata["target_format"] == target_format
        assert result.diagnostics[0].code == "SHEETFMT-OK"
        if source_format in {"xls", "ods", "et"} or target_format in {"xls", "ods", "et"}:
            assert bridge_calls
            first_bridge_source = source_format if source_format in {"xls", "ods", "et"} else "xlsx"
            assert [call[2] for call in bridge_calls] == [
                first_bridge_source,
                *(["xlsx"] * (len(bridge_calls) - 1)),
            ]

    @pytest.mark.parametrize("source_format", ["xls", "ods", "et"])
    def test_binary_to_csv_preserves_every_hub_sheet_as_an_artifact(
        self,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
        source_format: str,
    ) -> None:
        """XLS/ODS/ET→CSV must keep the old systems' one-CSV-per-sheet contract."""
        import csv

        from docwen_core.office_bridge import BridgeResult
        from docwen_plugin_spreadsheet.format_conversion import converter as sheet_module
        from docwen_plugin_spreadsheet.format_conversion.converter import SmartSheetConverter

        def fake_convert(input_path: str, output_path: str, **kwargs: Any) -> BridgeResult:
            _write_multisheet_marker_workbook(Path(output_path))
            return BridgeResult(True, output_path=output_path, backend="fake-office")

        monkeypatch.setattr(sheet_module, "convert_with_backend_priority", fake_convert)

        source_path = tmp_path / f"legacy-multi.{source_format}"
        source_path.write_bytes(f"fake {source_format} workbook".encode("ascii"))
        with tempfile.TemporaryDirectory() as staging:
            context = _build_fake_context(str(source_path), staging, target_format="csv")
            context.request.input_refs[0] = type(context.request.input_refs[0])(
                path=str(source_path), format=source_format, category="spreadsheet"
            )
            result = SmartSheetConverter().convert(context)

            assert result.success is True
            assert [artifact.suggested_name for artifact in result.artifacts] == [
                "legacy-multi_Alpha_Data.csv",
                "legacy-multi_Beta-Summary.csv",
                "legacy-multi_Hidden_Sheet.csv",
            ]
            assert [artifact.is_primary for artifact in result.artifacts] == [True, False, False]
            assert [artifact.kind for artifact in result.artifacts] == ["primary", "auxiliary", "auxiliary"]
            assert [artifact.metadata["sheet_name"] for artifact in result.artifacts] == [
                "Alpha Data",
                "Beta-Summary",
                "Hidden Sheet",
            ]
            assert all(artifact.metadata["source_format"] == source_format for artifact in result.artifacts)
            assert all(artifact.metadata["target_format"] == "csv" for artifact in result.artifacts)
            assert all(artifact.metadata["backend"] == "fake-office" for artifact in result.artifacts)

            markers: list[str] = []
            for artifact in result.artifacts:
                with open(artifact.staging_path, encoding="utf-8-sig", newline="") as handle:
                    markers.append(next(csv.reader(handle))[0])
            assert markers == ["ALPHA_MARKER", "BETA_MARKER", "HIDDEN_MARKER"]
            assert result.metrics.extra == {
                "backend": "fake-office",
                "sheet_count": 3,
                "total_rows": 3,
            }
            assert result.diagnostics[0].code == "SHEETFMT-OK"

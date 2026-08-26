"""Evaluate frozen POLICY-02 link=B,password=B against real ODS backends."""

from __future__ import annotations

import argparse
import contextlib
import hashlib
import io
import json
import os
import shutil
import subprocess
import sys
import uuid
import zipfile
from pathlib import Path
from typing import Any
from unittest import mock
from xml.etree import ElementTree

LINK_WARNING = "EXTERNAL_LINK_FLATTENED"
PROTECTION_WARNING = "PROTECTION_REMOVED_FOR_TARGET"
PASSWORD_REQUIRED = "PROTECTION_PASSWORD_REQUIRED"
PASSWORD_INVALID = "PROTECTION_PASSWORD_INVALID"
CONSENT_REQUIRED = "PROTECTION_LOSS_CONSENT_REQUIRED"
TESTED_REPO_PATHS = (
    "packages/apps/cli/src/docwen_cli/commands/convert.py",
    "packages/apps/cli/src/docwen_cli/presenters/json_presenter.py",
    "packages/apps/gui/src/docwen_gui/main_window.py",
    "packages/apps/gui/src/docwen_gui/widgets/conversion_panel.py",
    "packages/core/src/docwen_core/office_bridge.py",
    "packages/plugins/spreadsheet/src/docwen_plugin_spreadsheet/format_conversion/converter.py",
    "packages/plugins/spreadsheet/src/docwen_plugin_spreadsheet/format_conversion/xlsx_ods_policy.py",
    "packages/plugins/spreadsheet/src/docwen_plugin_spreadsheet/manifest.py",
    "tools/validation/evaluate_policy02_delivery_first.py",
)
_ODF_OFFICE_NS = "urn:oasis:names:tc:opendocument:xmlns:office:1.0"
_ODF_TABLE_NS = "urn:oasis:names:tc:opendocument:xmlns:table:1.0"
_ODF_XLINK_NS = "http://www.w3.org/1999/xlink"
_ODF_NS = {
    "office": _ODF_OFFICE_NS,
    "table": _ODF_TABLE_NS,
    "xlink": _ODF_XLINK_NS,
}


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _write_json(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(value, ensure_ascii=False, indent=2),
        encoding="utf-8",
        newline="\n",
    )


def _assert_credentials_redacted(text: str) -> None:
    for credential in ("test", "pwd", "wrong-policy02"):
        assert f'"{credential}"' not in text
        assert f"spreadsheet_password={credential}" not in text


def _backend_availability() -> dict[str, dict[str, Any]]:
    from docwen_core.office_bridge import find_soffice_path

    excel_available = False
    excel_detail = "Excel.Application COM registration not found"
    if sys.platform == "win32":
        import winreg

        try:
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Excel.Application\CLSID") as key:
                clsid, _kind = winreg.QueryValueEx(key, "")
            excel_available = bool(clsid)
            excel_detail = f"Excel.Application CLSID registered: {clsid}"
        except OSError:
            pass
    soffice = find_soffice_path()
    return {
        "msoffice_excel": {
            "configured": True,
            "available": excel_available,
            "detail": excel_detail,
        },
        "libreoffice": {
            "configured": True,
            "available": bool(soffice),
            "detail": str(soffice or "soffice not found"),
        },
    }


def _inspect_ods(path: Path) -> dict[str, Any]:
    with zipfile.ZipFile(path) as package:
        names = sorted(package.namelist())
        content_payload = package.read("content.xml")
    root = ElementTree.fromstring(content_payload)
    numeric_values = [
        str(cell.get(f"{{{_ODF_OFFICE_NS}}}value", ""))
        for cell in root.findall(".//table:table-cell", _ODF_NS)
        if cell.get(f"{{{_ODF_OFFICE_NS}}}value") is not None
    ]
    formulas = [
        str(cell.get(f"{{{_ODF_TABLE_NS}}}formula", ""))
        for cell in root.findall(".//table:table-cell", _ODF_NS)
        if cell.get(f"{{{_ODF_TABLE_NS}}}formula") is not None
    ]
    table_sources = [
        str(source.get(f"{{{_ODF_XLINK_NS}}}href", "")) for source in root.findall(".//table:table-source", _ODF_NS)
    ]
    protected_tables = [
        str(table.get(f"{{{_ODF_TABLE_NS}}}name", ""))
        for table in root.findall(".//table:table", _ODF_NS)
        if str(table.get(f"{{{_ODF_TABLE_NS}}}protected", "")).lower() in {"1", "true", "on"}
    ]
    structure_protected = any(
        str(spreadsheet.get(f"{{{_ODF_TABLE_NS}}}structure-protected", "")).lower() in {"1", "true", "on"}
        for spreadsheet in root.findall(".//office:spreadsheet", _ODF_NS)
    )
    decoded = content_payload.decode("utf-8", errors="replace")
    return {
        "package_entries": names,
        "numeric_values": numeric_values,
        "formulas": formulas,
        "table_sources": table_sources,
        "protected_tables": protected_tables,
        "structure_protected": structure_protected,
        "contains_ref_error": "#REF!" in decoded,
        "contains_file_uri": "file:///" in decoded,
        "content_sha256": hashlib.sha256(content_payload).hexdigest(),
    }


def _diagnostic_codes(result: Any) -> list[str]:
    return [str(diagnostic.code) for diagnostic in result.diagnostics]


def _error_code(result: Any) -> str:
    error = result.error
    return str(error.diagnostic_code if error is not None else "")


def _execute_runtime_case(
    runtime: Any,
    *,
    backend: str,
    case_id: str,
    source: Path,
    output_dir: Path,
    options: dict[str, Any],
    expected_success: bool,
    expected_codes: list[str],
    artifact_assertion: str | None,
) -> tuple[dict[str, Any], Any]:
    from docwen_core.models.file_ref import FileRef
    from docwen_core.models.request import ConversionRequest, OutputPolicy

    source_before = _sha256(source)
    output_dir.mkdir(parents=True, exist_ok=True)
    request = ConversionRequest(
        request_id=f"policy02-{backend}-{case_id}-{uuid.uuid4().hex[:8]}",
        input_refs=[
            FileRef(
                path=str(source),
                format="xlsx",
                category="spreadsheet",
                size_bytes=source.stat().st_size,
            )
        ],
        target_format="ods",
        options=dict(options),
        output_policy=OutputPolicy(output_dir=str(output_dir)),
        config_snapshot={
            "software": {
                "special_conversions": {
                    "ods": [backend],
                }
            }
        },
    )
    result = runtime.execute(request)
    assert not isinstance(result, list)
    assert result.success is expected_success
    artifact_projection: dict[str, Any] | None = None
    if expected_success:
        actual_codes = _diagnostic_codes(result)
        assert actual_codes == ["SHEETFMT-OK", *expected_codes, "FINALIZER_DONE"], (
            case_id,
            backend,
            actual_codes,
        )
        assert len(result.artifacts) == 1
        artifact = Path(result.artifacts[0].staging_path)
        assert artifact.is_file() and artifact.stat().st_size > 0
        artifact_projection = _inspect_ods(artifact)
        if artifact_assertion == "external":
            assert any(float(value) == 30.0 for value in artifact_projection["numeric_values"])
            assert artifact_projection["formulas"] == []
            assert artifact_projection["table_sources"] == []
            assert artifact_projection["contains_ref_error"] is False
            assert artifact_projection["contains_file_uri"] is False
        elif artifact_assertion == "unprotected":
            assert artifact_projection["structure_protected"] is False
            assert artifact_projection["protected_tables"] == []
        elif artifact_assertion == "control":
            assert "Foglio1" in artifact_projection["protected_tables"]
    else:
        assert _error_code(result) == expected_codes[0]
        assert result.artifacts == []
        assert not list(output_dir.glob("*.ods"))

    surfaced = json.dumps(result.to_dict(), ensure_ascii=False)
    _assert_credentials_redacted(surfaced)
    assert _sha256(source) == source_before
    projection = {
        "backend": backend,
        "case_id": case_id,
        "source": str(source),
        "source_sha256_before": source_before,
        "source_sha256_after": _sha256(source),
        "success": bool(result.success),
        "diagnostic_codes": _diagnostic_codes(result),
        "error_code": _error_code(result),
        "artifact": (
            {
                "path": str(result.artifacts[0].staging_path),
                "bytes": Path(result.artifacts[0].staging_path).stat().st_size,
                "sha256": _sha256(Path(result.artifacts[0].staging_path)),
                "projection": artifact_projection,
            }
            if expected_success
            else None
        ),
        "credential_redacted": True,
        "pass": True,
    }
    return projection, result


def _run_backend_matrix(
    source_root: Path,
    evidence_root: Path,
    availability: dict[str, dict[str, Any]],
) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    from docwen_bundle.runtime_factory import create_runtime_port

    link_owner = source_root / "link-external-workbook-b.xlsx"
    workbook_password = source_root / "workbookProtection-workbook_password-2013.xlsx"
    sheet_password = source_root / "workbookProtection-sheet_password-2013.xlsx"
    no_password = source_root / "sheetProtection_allLocked.xlsx"
    absent_dir = evidence_root / "inputs" / "link-target-absent"
    absent_dir.mkdir(parents=True)
    absent_owner = absent_dir / link_owner.name
    shutil.copy2(link_owner, absent_owner)
    assert not (absent_dir / "link-external-workbook-a.xlsx").exists()
    assert _sha256(absent_owner) == _sha256(link_owner)

    case_specs = [
        ("link-target-present", link_owner, {}, True, [LINK_WARNING], "external"),
        ("link-target-absent", absent_owner, {}, True, [LINK_WARNING], "external"),
        (
            "workbook-correct-consent",
            workbook_password,
            {
                "spreadsheet_password": "test",
                "allow_spreadsheet_protection_loss": True,
            },
            True,
            [PROTECTION_WARNING],
            "unprotected",
        ),
        ("workbook-missing-password", workbook_password, {}, False, [PASSWORD_REQUIRED], None),
        (
            "workbook-wrong-password",
            workbook_password,
            {
                "spreadsheet_password": "wrong-policy02",
                "allow_spreadsheet_protection_loss": True,
            },
            False,
            [PASSWORD_INVALID],
            None,
        ),
        (
            "workbook-no-consent",
            workbook_password,
            {
                "spreadsheet_password": "test",
            },
            False,
            [CONSENT_REQUIRED],
            None,
        ),
        (
            "sheet-correct-consent",
            sheet_password,
            {
                "spreadsheet_password": "pwd",
                "allow_spreadsheet_protection_loss": True,
            },
            True,
            [PROTECTION_WARNING],
            "unprotected",
        ),
        ("sheet-missing-password", sheet_password, {}, False, [PASSWORD_REQUIRED], None),
        (
            "sheet-wrong-password",
            sheet_password,
            {
                "spreadsheet_password": "wrong-policy02",
                "allow_spreadsheet_protection_loss": True,
            },
            False,
            [PASSWORD_INVALID],
            None,
        ),
        (
            "sheet-no-consent",
            sheet_password,
            {
                "spreadsheet_password": "pwd",
            },
            False,
            [CONSENT_REQUIRED],
            None,
        ),
        ("no-password-control", no_password, {}, True, [], "control"),
    ]
    slots: list[dict[str, Any]] = []
    retained_results: dict[str, Any] = {}
    runtime = create_runtime_port()
    try:
        for backend, backend_info in availability.items():
            if not backend_info["available"]:
                continue
            for case_id, source, options, success, codes, assertion in case_specs:
                projection, result = _execute_runtime_case(
                    runtime,
                    backend=backend,
                    case_id=case_id,
                    source=source,
                    output_dir=evidence_root / "backend-matrix" / backend / case_id,
                    options=options,
                    expected_success=success,
                    expected_codes=codes,
                    artifact_assertion=assertion,
                )
                slots.append(projection)
                retained_results[f"{backend}:{case_id}"] = result
    finally:
        runtime.shutdown()
    if not any(info["available"] for info in availability.values()):
        raise RuntimeError("No configured XLSX-to-ODS backend is available")
    return slots, retained_results


def _run_cli_case(arguments: list[str], *, password: str | None = None) -> dict[str, Any]:
    from docwen_bundle.cli_entry import main as cli_main

    stdout = io.StringIO()
    stderr = io.StringIO()
    password_patch = (
        mock.patch("getpass.getpass", return_value=password) if password is not None else contextlib.nullcontext()
    )
    with password_patch, contextlib.redirect_stdout(stdout), contextlib.redirect_stderr(stderr):
        return_code = cli_main(arguments)
    stdout_text = stdout.getvalue()
    stderr_text = stderr.getvalue()
    payload = json.loads(stdout_text)
    _assert_credentials_redacted(stdout_text)
    _assert_credentials_redacted(stderr_text)
    return {
        "arguments": ["<source>" if argument.lower().endswith(".xlsx") else argument for argument in arguments],
        "return_code": return_code,
        "payload": payload,
        "stderr": stderr_text,
        "credential_redacted": True,
    }


def _run_cli_projection(source_root: Path, evidence_root: Path) -> list[dict[str, Any]]:
    link = _run_cli_case(
        [
            "--json",
            "--yes",
            "run",
            str(source_root / "link-external-workbook-b.xlsx"),
            "--to",
            "ods",
            "--output",
            str(evidence_root / "cli" / "link"),
        ]
    )
    assert link["return_code"] == 0
    assert [warning["code"] for warning in link["payload"]["warnings"]] == [LINK_WARNING]

    missing = _run_cli_case(
        [
            "--json",
            "--yes",
            "run",
            str(source_root / "workbookProtection-workbook_password-2013.xlsx"),
            "--to",
            "ods",
            "--output",
            str(evidence_root / "cli" / "missing"),
        ]
    )
    assert missing["return_code"] == 2
    assert missing["payload"]["error"]["details"] == PASSWORD_REQUIRED

    correct = _run_cli_case(
        [
            "--json",
            "--yes",
            "run",
            str(source_root / "workbookProtection-workbook_password-2013.xlsx"),
            "--to",
            "ods",
            "--output",
            str(evidence_root / "cli" / "correct"),
            "--spreadsheet-password-prompt",
            "--allow-spreadsheet-protection-loss",
        ],
        password="test",
    )
    assert correct["return_code"] == 0
    assert [warning["code"] for warning in correct["payload"]["warnings"]] == [PROTECTION_WARNING]
    return [link, missing, correct]


def _run_gui_projection(
    source_root: Path,
    retained_results: dict[str, Any],
    availability: dict[str, dict[str, Any]],
) -> list[dict[str, Any]]:
    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
    from PySide6.QtWidgets import QLineEdit

    from docwen_gui.app import create_main_window, create_qapplication

    backend = next(name for name, info in availability.items() if info["available"])
    source = source_root / "workbookProtection-workbook_password-2013.xlsx"
    app = create_qapplication(["policy02-gui-oracle"])
    window = create_main_window(initial_files=[str(source)])
    app.processEvents()
    emitted: list[tuple[str, str, dict[str, Any]]] = []
    try:
        with contextlib.suppress(TypeError, RuntimeError):
            window._conversion_panel_vm.conversion_requested.disconnect(window._handle_conversion_requested)
        window._conversion_panel_vm.conversion_requested.connect(
            lambda target, path, options: emitted.append((target, path, dict(options)))
        )
        window._conversion_panel_vm.set_file_info(
            "spreadsheet",
            "xlsx",
            file_path=str(source),
            ui_mode="single",
        )
        panel = window._conversion_panel
        assert panel._spreadsheet_password_edit.echoMode() == QLineEdit.EchoMode.Password
        panel._conversion_combo.setCurrentText("ODS")
        panel._spreadsheet_password_edit.setText("test")
        panel._spreadsheet_protection_loss_checkbox.setChecked(True)
        panel._conversion_button.click()
        app.processEvents()
        assert emitted and emitted[-1][2] == {
            "spreadsheet_password": "test",
            "allow_spreadsheet_protection_loss": True,
        }
        assert panel._spreadsheet_password_edit.text() == ""
        assert panel._spreadsheet_protection_loss_checkbox.isChecked() is False

        request, context = window._build_request(
            file_path=str(source),
            target_format="ods",
            action_name="",
            options=emitted[-1][2],
        )
        assert request.options["spreadsheet_password"] == "test"
        assert context["options"]["spreadsheet_password"] == "<redacted>"
        assert "test" not in json.dumps(context, ensure_ascii=False)

        real_result = retained_results[f"{backend}:workbook-correct-consent"]
        operation_id = f"policy02-gui-{uuid.uuid4().hex[:8]}"
        real_result.task_id = operation_id
        context.update(
            {
                "request_id": operation_id,
                "file_path": str(source),
                "file_paths": [str(source)],
                "total_count": 1,
                "batch": False,
                "display_name": source.name,
            }
        )
        window._on_execution_finished(real_result, context)
        app.processEvents()
        rows = [row for row in window._info_area_vm.history_rows if row.operation_id == operation_id]
        warning_messages = [row.message for row in rows if row.message_type == "warning"]
        assert any("protection" in message.lower() for message in warning_messages)
        return [
            {
                "backend": backend,
                "masked_password": True,
                "explicit_consent_default_false": True,
                "password_cleared_after_emit": True,
                "context_password": "<redacted>",
                "history_row_types": [row.message_type for row in rows],
                "history_warning_messages": warning_messages,
                "pass": True,
            }
        ]
    finally:
        window.close()
        app.processEvents()


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--repo-root", type=Path, required=True)
    parser.add_argument("--source-root", type=Path, required=True)
    parser.add_argument("--evidence-root", type=Path, required=True)
    args = parser.parse_args()
    repo_root = args.repo_root.resolve()
    source_root = args.source_root.resolve()
    evidence_root = args.evidence_root.resolve()
    evidence_root.mkdir(parents=True, exist_ok=True)
    if any(evidence_root.iterdir()):
        raise RuntimeError(f"Evidence root must be empty: {evidence_root}")

    source_names = (
        "link-external-workbook-a.xlsx",
        "link-external-workbook-b.xlsx",
        "workbookProtection-workbook_password-2013.xlsx",
        "workbookProtection-sheet_password-2013.xlsx",
        "sheetProtection_allLocked.xlsx",
    )
    source_hashes_before = {name: _sha256(source_root / name) for name in source_names}
    availability = _backend_availability()
    slots, retained_results = _run_backend_matrix(
        source_root,
        evidence_root,
        availability,
    )
    cli_slots = _run_cli_projection(source_root, evidence_root)
    gui_slots = _run_gui_projection(source_root, retained_results, availability)
    source_hashes_after = {name: _sha256(source_root / name) for name in source_names}
    assert source_hashes_after == source_hashes_before

    available_backends = [name for name, info in availability.items() if info["available"]]
    expected_matrix_slots = 11 * len(available_backends)
    assert len(slots) == expected_matrix_slots
    assert all(slot["pass"] for slot in slots)
    result = {
        "stage": "VIS-2026-07-23-203",
        "policy": "POLICY-02 link=B,password=B",
        "repo_head_at_start": subprocess.check_output(
            ["git", "rev-parse", "HEAD"],
            cwd=repo_root,
            text=True,
        ).strip(),
        "repo_index_tree_at_start": subprocess.check_output(
            ["git", "write-tree"],
            cwd=repo_root,
            text=True,
        ).strip(),
        "tested_repo_path_hashes": {relative: _sha256(repo_root / relative) for relative in TESTED_REPO_PATHS},
        "source_root": str(source_root),
        "evidence_root": str(evidence_root),
        "source_hashes_before": source_hashes_before,
        "source_hashes_after": source_hashes_after,
        "backend_availability": availability,
        "available_backends": available_backends,
        "backend_matrix": slots,
        "cli_slots": cli_slots,
        "gui_slots": gui_slots,
        "counts": {
            "backend_matrix_expected": expected_matrix_slots,
            "backend_matrix_passed": len(slots),
            "cli_expected": 3,
            "cli_passed": len(cli_slots),
            "gui_expected": 1,
            "gui_passed": len(gui_slots),
        },
        "credential_redaction_checked": True,
        "source_immutability_checked": True,
        "target_absent_execution_checked": True,
        "pass": True,
        "accepted_boundary": (
            "External links become static cached values and cannot update. "
            "Password-protected workbook/sheet controls are removed only from "
            "a private copy after a correct request-scoped password and explicit "
            "loss consent; the published ODS is intentionally unprotected."
        ),
    }
    serialized = json.dumps(result, ensure_ascii=False)
    _assert_credentials_redacted(serialized)
    _write_json(evidence_root / "result.json", result)
    print(json.dumps(result["counts"], ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())

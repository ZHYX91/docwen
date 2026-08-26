from __future__ import annotations

import ast
from pathlib import Path

import pytest
from tools.validation.source_family import read_source_text

pytestmark = pytest.mark.contract

PROJECT_ROOT = Path(__file__).resolve().parents[2]
REPORT_NAME = "successful-result-warning-consumer-parity-2026-07-16.md"


def _read(relative_path: str) -> str:
    return read_source_text(PROJECT_ROOT / relative_path)


def _production_warning_codes() -> set[str]:
    codes: set[str] = set()
    for path in (PROJECT_ROOT / "packages" / "plugins").rglob("*.py"):
        if "tests" in path.parts:
            continue
        tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
        for node in ast.walk(tree):
            if not isinstance(node, ast.Call):
                continue
            name = ""
            if isinstance(node.func, ast.Name):
                name = node.func.id
            elif isinstance(node.func, ast.Attribute):
                name = node.func.attr
            if name not in {"ConversionDiagnostic", "report_diagnostic"}:
                continue
            keywords = {keyword.arg: keyword.value for keyword in node.keywords if keyword.arg is not None}
            level = keywords.get("level")
            code = keywords.get("code")
            if name == "report_diagnostic":
                if level is None and node.args:
                    level = node.args[0]
                if code is None and len(node.args) >= 3:
                    code = node.args[2]
            if (
                isinstance(level, ast.Constant)
                and level.value == "warning"
                and isinstance(code, ast.Constant)
                and isinstance(code.value, str)
            ):
                codes.add(code.value)
    return codes


def test_cli_presenters_consume_final_result_warning_diagnostics() -> None:
    json_presenter = _read("packages/apps/cli/src/docwen_cli/presenters/json_presenter.py")
    text_presenter = _read("packages/apps/cli/src/docwen_cli/presenters/text_presenter.py")
    cli_tests = _read("packages/apps/cli/tests/test_cli_json.py") + _read(
        "packages/apps/cli/tests/test_cli_text_diagnostics.py"
    )

    assert json_presenter.count("self._result_warning_payloads(") == 2
    assert 'payload["file"] = str(file_path)' in json_presenter
    assert 'getattr(diagnostic, "level", "") != "warning"' in json_presenter
    assert text_presenter.count("self._present_warnings(") == 2
    assert "print(text, file=sys.stderr)" in text_presenter
    assert "test_single_projects_only_warning_diagnostics" in cli_tests
    assert "test_batch_projects_result_warning_diagnostics_with_file_context" in cli_tests
    assert "test_single_success_writes_warning_diagnostic_to_stderr" in cli_tests


def test_gui_success_consumers_keep_success_state_with_warning_tone() -> None:
    main_window = _read("packages/apps/gui/src/docwen_gui/main_window.py")
    gui_tests = _read("packages/apps/gui/tests/test_main_window_projection_binding_*.py")

    assert main_window.count("_result_warning_messages(") == 3
    assert 'completion_tone = "warning" if warning_messages else "success"' in main_window
    assert 'tone = "warning" if warning_rows else "success"' in main_window
    assert 'f"{Path(warning_file).name}: {warning_message}"' in main_window
    assert "test_success_callback_projects_warning_diagnostics_to_info_area" in gui_tests
    assert "test_batch_all_success_with_warning_keeps_success_state_and_warning_tone" in gui_tests
    assert 'warning_row.property("infoStatusTone") == "warning"' in gui_tests

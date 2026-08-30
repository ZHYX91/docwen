"""Contract tests for the breaking DocWen CLI protocol 3 foundation."""

from __future__ import annotations

import json
from importlib import import_module

import pytest

pytestmark = pytest.mark.contract


def test_protocol_envelope_has_exact_top_level_shape() -> None:
    from docwen_cli.protocol import make_envelope

    envelope = make_envelope(command="info", success=True, data={"ready": True})

    assert envelope == {
        "protocol_version": 3,
        "product_version": "0.9.1",
        "success": True,
        "command": "info",
        "data": {"ready": True},
        "error": None,
        "warnings": [],
        "meta": {},
    }


def test_json_presenter_emits_typed_protocol_error(capsys: pytest.CaptureFixture[str]) -> None:
    from docwen_cli.presenters.json_presenter import JsonPresenter

    JsonPresenter().present_error(
        "gui status",
        "DocWen GUI is not running",
        error_code="gui_not_running",
        details={"running": False},
        hint="Start the GUI or use gui open.",
    )

    payload = json.loads(capsys.readouterr().out)
    assert payload["protocol_version"] == 3
    assert payload["product_version"] == "0.9.1"
    assert payload["error"] == {
        "category": "unavailable",
        "code": "gui_not_running",
        "message": "DocWen GUI is not running",
        "details": {"running": False},
        "hint": "Start the GUI or use gui open.",
    }


def test_warning_strings_are_normalized(capsys: pytest.CaptureFixture[str]) -> None:
    from docwen_cli.presenters.json_presenter import JsonPresenter

    presenter = JsonPresenter()
    presenter.add_warning("fallback used")
    presenter.present_data("doctor", {"healthy": True})

    payload = json.loads(capsys.readouterr().out)
    assert payload["warnings"] == [{"code": "warning", "message": "fallback used"}]


def test_large_unicode_details_remain_one_complete_json_document(capsys: pytest.CaptureFixture[str]) -> None:
    from docwen_cli.presenters.json_presenter import JsonPresenter, JsonValue

    details: dict[str, JsonValue] = {"path": "C:\\资料\\" + "很长的文件名" * 10_000}
    JsonPresenter().present_error("convert", "无法写入目标", error_code="invalid_input", details=details)

    captured = capsys.readouterr()
    payload = json.loads(captured.out)
    assert captured.err == ""
    assert payload["error"]["details"] == details
    assert captured.out.count('"protocol_version"') == 1


def test_unknown_exception_does_not_leak_message_path_or_traceback(
    capsys: pytest.CaptureFixture[str], monkeypatch: pytest.MonkeyPatch
) -> None:
    from docwen_cli.exit_codes import ExitCode

    cli_main_module = import_module("docwen_cli.main")
    secret = r"C:\Users\private-user\secret-document.docx"

    def _raise(_args: object, _controller: object | None = None) -> int:
        raise RuntimeError(f"sensitive failure at {secret}")

    cli_main_module._init_command_table()
    monkeypatch.setitem(cli_main_module._COMMAND_TABLE, "doctor", (_raise, False))
    monkeypatch.setattr(cli_main_module, "run_security_protections", lambda **_kwargs: None)

    exit_code = cli_main_module.main(["doctor", "--json"])

    captured = capsys.readouterr()
    payload = json.loads(captured.out)
    assert exit_code == int(ExitCode.INTERNAL_ERROR)
    assert payload["error"]["code"] == "unknown_error"
    assert payload["error"]["details"] == {"exception_type": "RuntimeError"}
    assert secret not in captured.out
    assert "sensitive failure" not in captured.out
    assert "traceback" not in captured.out.lower()
    assert captured.err == ""


@pytest.mark.parametrize(
    ("error_code", "exit_code"),
    [
        ("invalid_arguments", 2),
        ("resource_not_found", 3),
        ("dependency_missing", 4),
        ("security_check_failed", 5),
        ("network_access_blocked", 5),
        ("gui_not_running", 6),
        ("gui_control_endpoint_unavailable", 6),
        ("output_exists", 7),
        ("operation_timeout", 8),
        ("batch_partial_failure", 9),
        ("operation_cancelled", 130),
        ("internal_error", 1),
    ],
)
def test_protocol_error_code_exit_mapping(error_code: str, exit_code: int) -> None:
    from docwen_cli.exit_codes import exit_code_from_error_code

    assert int(exit_code_from_error_code(error_code)) == exit_code


def test_network_block_preserves_typed_security_envelope_during_bootstrap(
    capsys: pytest.CaptureFixture[str],
) -> None:
    from argparse import Namespace

    from docwen_cli.exit_codes import ExitCode
    from docwen_cli.main import _handle_bootstrap_error
    from docwen_runtime.security import NetworkAccessBlockedError

    exit_code = _handle_bootstrap_error(
        "doctor",
        Namespace(json=True),
        NetworkAccessBlockedError("socket.getaddrinfo"),
    )

    payload = json.loads(capsys.readouterr().out)
    assert exit_code == int(ExitCode.SECURITY_CHECK_FAILED)
    assert payload["error"]["category"] == "security"
    assert payload["error"]["code"] == "network_access_blocked"

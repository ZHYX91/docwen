"""Tests for the lightweight protocol 3 discovery entry points."""

from __future__ import annotations

import json
import platform

import pytest

pytestmark = pytest.mark.contract


def test_info_json_does_not_initialize_runtime(capsys: pytest.CaptureFixture[str]) -> None:
    from docwen_cli.main import main

    def forbidden_factory():
        raise AssertionError("info must not initialize runtime or config")

    exit_code = main(
        ["info", "--json"],
        runtime_port_factory=forbidden_factory,
        config_port_factory=forbidden_factory,
    )

    captured = capsys.readouterr()
    payload = json.loads(captured.out)
    assert exit_code == 0
    assert captured.err == ""
    assert payload["command"] == "info"
    assert payload["protocol_version"] == 3
    assert payload["data"]["protocol"]["major"] == 3
    assert payload["data"]["product"]["version"] == "0.9.1"
    capabilities = {item["id"]: item for item in payload["data"]["capabilities"]}
    assert capabilities["cli.schema"] == {
        "id": "cli.schema",
        "contract_version": 1,
        "state": "available",
        "available": True,
        "platforms": ["windows", "linux", "darwin"],
        "current_platform_supported": True,
        "runtime_check_required": False,
    }
    assert capabilities["cli.convert"]["state"] == "runtime_check_required"
    gui_supported = platform.system().lower() == "windows"
    gui_control = capabilities["gui.control"]
    assert gui_control["platforms"] == ["windows"]
    assert gui_control["current_platform_supported"] is gui_supported
    assert gui_control["available"] is gui_supported
    assert gui_control["runtime_check_required"] is gui_supported
    assert gui_control["state"] == ("runtime_check_required" if gui_supported else "unavailable")
    assert gui_control["contract_version"] == 1
    gui_settings = capabilities["gui.settings"]
    assert gui_settings["contract_version"] == 1
    assert gui_settings["platforms"] == ["windows"]
    assert gui_settings["current_platform_supported"] is gui_supported
    assert gui_settings["available"] is gui_supported
    assert gui_settings["runtime_check_required"] is gui_supported
    assert gui_settings["state"] == ("runtime_check_required" if gui_supported else "unavailable")
    assert gui_settings["details"] == {
        "cold_start": True,
        "sections": ["proofread"],
    }


def test_schema_json_does_not_initialize_runtime(capsys: pytest.CaptureFixture[str]) -> None:
    from docwen_cli.main import main

    def forbidden_factory():
        raise AssertionError("schema must not initialize runtime or config")

    exit_code = main(
        ["schema", "convert", "--json"],
        runtime_port_factory=forbidden_factory,
        config_port_factory=forbidden_factory,
    )

    captured = capsys.readouterr()
    payload = json.loads(captured.out)
    assert exit_code == 0
    assert captured.err == ""
    assert payload["command"] == "schema"
    assert payload["data"]["command"] == "convert"


def test_capability_availability_is_effective_for_current_platform(monkeypatch: pytest.MonkeyPatch) -> None:
    from docwen_cli.commands import info

    monkeypatch.setattr(info.platform, "system", lambda: "Darwin")

    capabilities = {item["id"]: item for item in info.build_info_data()["capabilities"]}
    gui_control = capabilities["gui.control"]
    assert gui_control["current_platform_supported"] is False
    assert gui_control["available"] is False
    assert gui_control["state"] == "unavailable"
    assert gui_control["runtime_check_required"] is False
    assert "darwin" in gui_control["reason"].lower()

    gui_settings = capabilities["gui.settings"]
    assert gui_settings["current_platform_supported"] is False
    assert gui_settings["available"] is False
    assert gui_settings["runtime_check_required"] is False
    assert gui_settings["details"] == {"cold_start": True, "sections": ["proofread"]}

    inspect = capabilities["cli.inspect"]
    assert inspect["current_platform_supported"] is True
    assert inspect["available"] is True
    assert inspect["state"] == "available"

    for capability_id in (
        "cli.convert",
        "cli.validate",
        "cli.number.markdown",
        "cli.merge",
        "cli.split.pdf",
    ):
        operation = capabilities[capability_id]
        assert operation["platforms"] == ["windows", "linux"]
        assert operation["current_platform_supported"] is False
        assert operation["available"] is False
        assert operation["state"] == "unavailable"
        assert operation["runtime_check_required"] is False
        assert "darwin" in operation["reason"].lower()


@pytest.mark.parametrize("argv", [["--version"], ["info", "--version"]])
def test_version_is_lightweight(argv: list[str], capsys: pytest.CaptureFixture[str]) -> None:
    from docwen_cli.main import main

    def forbidden_factory():
        raise AssertionError("--version must not initialize runtime or config")

    exit_code = main(
        argv,
        runtime_port_factory=forbidden_factory,
        config_port_factory=forbidden_factory,
    )

    captured = capsys.readouterr()
    assert exit_code == 0
    assert captured.out.strip() == "DocWen 0.9.1 (CLI protocol 3)"
    assert captured.err == ""


def test_invalid_command_returns_usage_exit_code() -> None:
    from docwen_cli.main import main

    assert main(["does-not-exist"]) == 2


def test_invalid_command_json_is_one_typed_document(capsys: pytest.CaptureFixture[str]) -> None:
    from docwen_cli.main import main

    exit_code = main(["does-not-exist", "--json"])

    captured = capsys.readouterr()
    payload = json.loads(captured.out)
    assert exit_code == 2
    assert captured.err == ""
    assert payload["success"] is False
    assert payload["error"]["category"] == "invalid_input"
    assert payload["error"]["code"] == "invalid_arguments"

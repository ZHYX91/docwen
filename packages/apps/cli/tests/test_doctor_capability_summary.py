"""Doctor must report one canonical, path-free runtime health model."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

import pytest

pytestmark = pytest.mark.unit


def _projection(*, layout_available: bool = True) -> dict[str, Any]:
    route = {
        "id": "layout:pdf:md:convert",
        "operation": "conversion",
        "source": "pdf",
        "target": "md",
        "action": None,
        "plugin": "layout",
        "available": layout_available,
        "state": "available" if layout_available else "unavailable",
        "platforms": ["windows"],
        "platform_supported": True,
        "required_capabilities": ["python.pymupdf4llm"],
        "optional_capabilities": [],
        "missing_required_capabilities": [] if layout_available else ["python.pymupdf4llm"],
        "missing_optional_capabilities": [],
        "limitations": [] if layout_available else ["required_capability_unavailable:python.pymupdf4llm"],
        "options": [],
    }
    return {
        "contract": {"id": "docwen.runtime-capabilities", "version": 1},
        "runtime": {"state": "available", "platform": "windows"},
        "security": {
            "dependency_egress_guard": {
                "state": "enforced",
                "installed": True,
                "active": True,
                "scope": "docwen_python_process",
                "policy": "deny_dns_and_ip",
                "mechanism": "cpython_audit_hook",
                "bootstrap": "composition_root",
                "local_transports": ["windows_named_pipe", "unix_domain_socket"],
                "external_processes": "not_managed",
            }
        },
        "gates": [
            {
                "id": "python.pymupdf4llm",
                "kind": "python_module_with_resources",
                "label": "PyMuPDF4LLM",
                "available": layout_available,
                "reason": None if layout_available else "resource_hash_mismatch",
                "module": "pymupdf4llm",
            }
        ],
        "sources": [{"id": "pdf", "category": "layout", "available": layout_available, "routes": [route]}],
        "counts": {
            "sources": 1,
            "routes": 1,
            "available_routes": int(layout_available),
            "unavailable_routes": int(not layout_available),
            "actions": 0,
        },
    }


class _ConfigPort:
    def __init__(self, *, error: Exception | None = None) -> None:
        self._error = error

    def snapshot(self) -> dict[str, Any]:
        if self._error is not None:
            raise self._error
        return {"loaded": True}


class _Controller:
    def __init__(self, projection: dict[str, Any], *, config_error: Exception | None = None) -> None:
        self._projection = projection
        self.config_port = _ConfigPort(error=config_error)
        self.describe_calls = 0

    def describe_runtime_capabilities(self) -> dict[str, Any]:
        self.describe_calls += 1
        return self._projection


def _args(*, json_mode: bool) -> argparse.Namespace:
    return argparse.Namespace(json=json_mode, quiet=json_mode, verbose=False)


def test_doctor_json_uses_canonical_projection_once(capsys: pytest.CaptureFixture[str]) -> None:
    from docwen_cli.commands.doctor import execute_doctor

    projection = _projection()
    controller = _Controller(projection)

    exit_code = execute_doctor(_args(json_mode=True), controller)
    payload = json.loads(capsys.readouterr().out)

    assert exit_code == 0
    assert controller.describe_calls == 1
    assert payload["success"] is True
    assert payload["command"] == "doctor"
    assert payload["data"]["all_ok"] is True
    assert payload["data"]["capability_summary"] == projection
    assert [check["id"] for check in payload["data"]["checks"]] == [
        "path.temp_directory",
        "config.load",
        "security.dependency_egress_guard",
    ]
    assert all(check["status"] == "ok" for check in payload["data"]["checks"])
    assert all(check["kind"] in {"path", "config", "security"} for check in payload["data"]["checks"])


def test_doctor_required_gate_failure_stays_in_capability_summary(capsys: pytest.CaptureFixture[str]) -> None:
    from docwen_cli.commands.doctor import execute_doctor

    projection = _projection(layout_available=False)
    controller = _Controller(projection)

    exit_code = execute_doctor(_args(json_mode=True), controller)
    payload = json.loads(capsys.readouterr().out)

    assert exit_code == 0
    assert payload["success"] is True
    assert payload["data"]["all_ok"] is True
    assert payload["data"]["capability_summary"] == projection
    assert payload["data"]["capability_summary"]["gates"][0]["available"] is False
    assert payload["data"]["capability_summary"]["sources"][0]["routes"][0]["available"] is False
    assert {check["id"] for check in payload["data"]["checks"]} == {
        "path.temp_directory",
        "config.load",
        "security.dependency_egress_guard",
    }
    assert "\\" not in json.dumps(payload["data"]["checks"])


def test_doctor_optional_gate_failure_does_not_fail_base_health(
    capsys: pytest.CaptureFixture[str],
) -> None:
    from docwen_cli.commands.doctor import execute_doctor

    projection = _projection()
    projection["gates"].append(
        {
            "id": "python.pillow_heif",
            "kind": "python_module",
            "label": "pillow-heif",
            "available": False,
            "reason": "module_not_available",
            "module": "pillow_heif",
        }
    )
    route = projection["sources"][0]["routes"][0]
    route["state"] = "available_with_limits"
    route["optional_capabilities"] = ["python.pillow_heif"]
    route["missing_optional_capabilities"] = ["python.pillow_heif"]
    route["limitations"] = ["optional_capability_unavailable:python.pillow_heif"]

    exit_code = execute_doctor(_args(json_mode=True), _Controller(projection))
    payload = json.loads(capsys.readouterr().out)

    assert exit_code == 0
    assert payload["data"]["all_ok"] is True
    assert payload["data"]["capability_summary"] == projection
    assert payload["data"]["capability_summary"]["gates"][-1]["available"] is False
    assert payload["data"]["capability_summary"]["sources"][0]["routes"][0]["available"] is True
    assert [check["id"] for check in payload["data"]["checks"]] == [
        "path.temp_directory",
        "config.load",
        "security.dependency_egress_guard",
    ]


def test_doctor_external_office_failure_does_not_fail_base_health(
    capsys: pytest.CaptureFixture[str],
) -> None:
    from docwen_cli.commands.doctor import execute_doctor

    projection = _projection()
    projection["gates"].append(
        {
            "id": "external_office.word",
            "kind": "external_office",
            "label": "Word-compatible Office backend",
            "available": False,
            "reason": "no_compatible_backend",
            "providers": [],
        }
    )
    office_route = {
        "id": "markdown:md:docx:convert",
        "operation": "conversion",
        "source": "md",
        "target": "docx",
        "action": None,
        "plugin": "markdown",
        "available": False,
        "state": "unavailable",
        "platforms": ["windows"],
        "platform_supported": True,
        "required_capabilities": ["external_office.word"],
        "optional_capabilities": [],
        "missing_required_capabilities": ["external_office.word"],
        "missing_optional_capabilities": [],
        "limitations": ["required_capability_unavailable:external_office.word"],
        "options": [],
    }
    projection["sources"].append({"id": "md", "category": "markup", "available": False, "routes": [office_route]})
    projection["counts"].update({"sources": 2, "routes": 2, "available_routes": 1, "unavailable_routes": 1})

    exit_code = execute_doctor(_args(json_mode=True), _Controller(projection))
    payload = json.loads(capsys.readouterr().out)

    assert exit_code == 0
    assert payload["data"]["all_ok"] is True
    assert payload["data"]["capability_summary"] == projection
    assert payload["data"]["capability_summary"]["gates"][-1]["available"] is False
    assert payload["data"]["capability_summary"]["sources"][-1]["routes"][0]["available"] is False
    assert [check["id"] for check in payload["data"]["checks"]] == [
        "path.temp_directory",
        "config.load",
        "security.dependency_egress_guard",
    ]


def test_doctor_config_failure_uses_stable_reason_not_exception_text(capsys: pytest.CaptureFixture[str]) -> None:
    from docwen_cli.commands.doctor import execute_doctor
    from docwen_cli.exit_codes import ExitCode

    secret_path = r"C:\Users\private\configs\software.toml"
    controller = _Controller(_projection(), config_error=OSError(secret_path))

    exit_code = execute_doctor(_args(json_mode=True), controller)
    payload = json.loads(capsys.readouterr().out)

    assert exit_code == int(ExitCode.UNAVAILABLE)
    config = next(check for check in payload["data"]["checks"] if check["id"] == "config.load")
    assert config["reason"] == "config_load_failed"
    assert secret_path not in json.dumps(payload, ensure_ascii=False)


def test_doctor_fails_when_dependency_egress_guard_is_not_enforced(
    capsys: pytest.CaptureFixture[str],
) -> None:
    from docwen_cli.commands.doctor import execute_doctor
    from docwen_cli.exit_codes import ExitCode

    projection = _projection()
    projection["security"]["dependency_egress_guard"].update(
        {"state": "installed", "active": False},
    )

    exit_code = execute_doctor(_args(json_mode=True), _Controller(projection))
    payload = json.loads(capsys.readouterr().out)

    assert exit_code == int(ExitCode.UNAVAILABLE)
    assert payload["data"]["all_ok"] is False
    guard = next(check for check in payload["data"]["checks"] if check["id"] == "security.dependency_egress_guard")
    assert guard["status"] == "fail"
    assert guard["reason"] == "not_enforced"


def test_temp_probe_does_not_return_the_host_path(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    from docwen_cli.commands import doctor

    monkeypatch.setattr(doctor.tempfile, "gettempdir", lambda: str(tmp_path))

    check = doctor._check_temp_directory()

    assert check.ok is True
    assert check.id == "path.temp_directory"
    assert str(tmp_path) not in json.dumps(check.to_dict())


def test_doctor_temp_failure_is_typed_path_free_and_nonzero(
    capsys: pytest.CaptureFixture[str], monkeypatch: pytest.MonkeyPatch
) -> None:
    from docwen_cli.commands import doctor
    from docwen_cli.exit_codes import ExitCode

    secret_path = r"C:\Users\private\AppData\Local\Temp"

    def _fail_probe(*args: Any, **kwargs: Any) -> tuple[int, str]:
        del args, kwargs
        raise OSError(secret_path)

    monkeypatch.setattr(doctor.tempfile, "mkstemp", _fail_probe)

    exit_code = doctor.execute_doctor(_args(json_mode=True), _Controller(_projection()))
    payload = json.loads(capsys.readouterr().out)

    assert exit_code == int(ExitCode.UNAVAILABLE)
    assert payload["success"] is True
    assert payload["data"]["all_ok"] is False
    temp = next(check for check in payload["data"]["checks"] if check["id"] == "path.temp_directory")
    assert temp["kind"] == "path"
    assert temp["status"] == "fail"
    assert temp["reason"] == "not_writable"
    assert secret_path not in json.dumps(payload, ensure_ascii=False)


def test_doctor_text_reports_runtime_counts(capsys: pytest.CaptureFixture[str]) -> None:
    from docwen_cli.commands.doctor import execute_doctor

    exit_code = execute_doctor(_args(json_mode=False), _Controller(_projection()))
    output = capsys.readouterr().out

    assert exit_code == 0
    assert "DocWen 环境诊断" in output
    assert "── 基础检查 ──" in output
    assert "✓ 基础检查通过" in output
    assert "── 运行时能力 ──" in output
    assert "1/1 可用" in output


def test_main_doctor_without_runtime_is_typed_failure(capsys: pytest.CaptureFixture[str]) -> None:
    from docwen_cli.exit_codes import ExitCode
    from docwen_cli.main import main

    exit_code = main(["doctor", "--json", "--quiet"])
    payload = json.loads(capsys.readouterr().out)

    assert exit_code == int(ExitCode.UNAVAILABLE)
    assert payload["success"] is False
    assert payload["error"]["code"] == "capability_unavailable"


def test_main_doctor_uses_injected_controller(capsys: pytest.CaptureFixture[str]) -> None:
    from docwen_cli.main import main

    controller = _Controller(_projection())
    exit_code = main(["doctor", "--json", "--quiet"], controller=controller)
    payload = json.loads(capsys.readouterr().out)

    assert exit_code == 0
    assert payload["data"]["all_ok"] is True
    assert controller.describe_calls == 1

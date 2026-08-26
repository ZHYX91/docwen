from __future__ import annotations

import json
import os
import subprocess
import sys
from pathlib import Path

import pytest

pytestmark = pytest.mark.unit


def _fake_numbering_payload(args: tuple[str, ...]) -> dict[str, object]:
    output_file = Path(args[args.index("--output") + 1])
    operation = args[args.index("--operation") + 1]
    content = (
        "# 一、已有标题\n\n## （一）二级标题\n\n正文。\n"
        if operation == "add"
        else "# 已有标题\n\n## 二级标题\n\n正文。\n"
    )
    output_file.write_text(content, encoding="utf-8")
    return {
        "protocol_version": 3,
        "success": True,
        "command": "number",
        "data": {"output": str(output_file)},
        "error": None,
    }


def _fake_capability_payload() -> dict[str, object]:
    conversion_route = {
        "id": "probe:pdf:md:convert",
        "operation": "conversion",
        "source": "pdf",
        "target": "md",
        "action": None,
        "plugin": "probe",
        "available": True,
        "state": "available",
        "platforms": ["windows"],
        "platform_supported": True,
        "required_capabilities": ["python.pymupdf4llm"],
        "optional_capabilities": [],
        "limitations": [],
        "options": [],
    }
    action_route = {
        "id": "probe:pdf:pdf:split_pdf",
        "operation": "action",
        "source": "pdf",
        "target": "pdf",
        "action": "split_pdf",
        "plugin": "probe",
        "available": True,
        "state": "available",
        "platforms": ["windows"],
        "platform_supported": True,
        "required_capabilities": [],
        "optional_capabilities": [],
        "limitations": [],
        "options": [],
    }
    return {
        "protocol_version": 3,
        "success": True,
        "command": "resources list",
        "data": {
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
                    "bootstrap": "pyinstaller_runtime_hook",
                    "local_transports": ["windows_named_pipe", "unix_domain_socket"],
                    "external_processes": "not_managed",
                }
            },
            "gates": [{"id": "python.pymupdf4llm", "available": True}],
            "sources": [
                {
                    "id": "pdf",
                    "category": "layout",
                    "available": True,
                    "routes": [conversion_route, action_route],
                }
            ],
            "counts": {"sources": 1, "routes": 2, "available_routes": 2, "unavailable_routes": 0, "actions": 1},
        },
        "error": None,
    }


def _fake_doctor_payload() -> dict[str, object]:
    return {
        "protocol_version": 3,
        "success": True,
        "command": "doctor",
        "data": {
            "checks": [
                {"id": "path.temp_directory", "kind": "path", "status": "ok", "reason": None},
                {"id": "config.load", "kind": "config", "status": "ok", "reason": None},
                {
                    "id": "security.dependency_egress_guard",
                    "kind": "security",
                    "status": "ok",
                    "reason": None,
                },
            ],
            "all_ok": True,
            "capability_summary": {
                "security": {
                    "dependency_egress_guard": {
                        "state": "enforced",
                        "installed": True,
                        "active": True,
                        "scope": "docwen_python_process",
                        "policy": "deny_dns_and_ip",
                        "mechanism": "cpython_audit_hook",
                        "bootstrap": "pyinstaller_runtime_hook",
                        "local_transports": ["windows_named_pipe", "unix_domain_socket"],
                        "external_processes": "not_managed",
                    }
                },
                "gates": [{"id": "python.pymupdf4llm", "available": True}],
            },
        },
        "error": None,
    }


def _fake_optimization_payload() -> dict[str, object]:
    resources = [
        {
            "id": "gongwen",
            "name": "Gongwen",
            "action_name": "gongwen",
            "scopes": ["document_to_md"],
            "available": True,
            "state": "available",
            "bindings": [
                {
                    "scope": "document_to_md",
                    "route_id": "gongwen:docx:md:gongwen",
                    "source": "docx",
                    "source_category": "document",
                    "target": "md",
                    "available": True,
                    "state": "available",
                }
            ],
        },
        {
            "id": "invoice_cn",
            "name": "Invoice CN",
            "action_name": "invoice_cn",
            "scopes": ["layout_to_md", "image_to_md"],
            "available": True,
            "state": "available",
            "bindings": [
                {
                    "scope": "layout_to_md",
                    "route_id": "invoice:pdf:md:invoice_cn",
                    "source": "pdf",
                    "source_category": "layout",
                    "target": "md",
                    "available": True,
                    "state": "available",
                },
                {
                    "scope": "layout_to_md",
                    "route_id": "invoice:ofd:md:invoice_cn",
                    "source": "ofd",
                    "source_category": "layout",
                    "target": "md",
                    "available": True,
                    "state": "available",
                },
                {
                    "scope": "image_to_md",
                    "route_id": "invoice:image:md:invoice_cn",
                    "source": "image",
                    "source_category": "image",
                    "target": "md",
                    "available": True,
                    "state": "available",
                },
            ],
        },
    ]
    bindings = [binding for resource in resources for binding in resource["bindings"]]
    return {
        "protocol_version": 3,
        "success": True,
        "command": "resources list",
        "data": {
            "resource": "optimizations",
            "contract": {"id": "docwen.optimizations", "version": 1},
            "runtime": {"state": "available", "platform": "windows"},
            "resources": resources,
            "counts": {
                "resources": len(resources),
                "available_resources": len(resources),
                "unavailable_resources": 0,
                "bindings": len(bindings),
                "available_bindings": len(bindings),
                "unavailable_bindings": 0,
            },
        },
        "error": None,
    }


def test_packaged_cli_verifier_runs_optional_successful_warning_smoke(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from scripts.release import verify_packaged_cli

    binary_dir = tmp_path / "dist"
    binary_dir.mkdir()
    binary_name = "DocWenCLI.exe"
    (binary_dir / binary_name).write_text("placeholder", encoding="utf-8")
    warning_input = tmp_path / "rules.docx"
    warning_input.write_bytes(b"docx")
    expected_message = "缺少必需字段：成文日期"
    calls: list[tuple[str, ...]] = []

    def fake_write_xlsx(path: Path) -> None:
        path.write_bytes(b"xlsx")

    def fake_run(binary_path: Path, *args: str, cwd: Path) -> subprocess.CompletedProcess[str]:
        calls.append(args)
        (cwd / "log_home" / "logs").mkdir(parents=True, exist_ok=True)
        (cwd / "log_home" / "logs" / "docwen.log").write_text("ok", encoding="utf-8")
        if args[:3] == ("resources", "list", "formats"):
            payload = _fake_capability_payload()
            return subprocess.CompletedProcess([str(binary_path), *args], 0, stdout=json.dumps(payload), stderr="")
        if args[:3] == ("resources", "list", "optimizations"):
            payload = _fake_optimization_payload()
            return subprocess.CompletedProcess([str(binary_path), *args], 0, stdout=json.dumps(payload), stderr="")
        if args[:1] == ("doctor",):
            payload = _fake_doctor_payload()
            return subprocess.CompletedProcess([str(binary_path), *args], 0, stdout=json.dumps(payload), stderr="")
        if args[:2] == ("number", "markdown"):
            payload = _fake_numbering_payload(args)
            return subprocess.CompletedProcess([str(binary_path), *args], 0, stdout=json.dumps(payload), stderr="")
        if "--optimization" in args:
            output_dir = Path(args[args.index("--output") + 1])
            output_dir.mkdir(parents=True, exist_ok=True)
            root_name = "rules_20260824_120102_fromDocx"
            output_file = output_dir / root_name / f"{root_name}.md"
            os.makedirs(verify_packaged_cli._native_long_path(output_file.parent))
            with open(verify_packaged_cli._native_long_path(output_file), "w", encoding="utf-8") as stream:
                stream.write("# rules\n")
            if "--json" in args:
                payload = {
                    "protocol_version": 3,
                    "success": True,
                    "command": "convert",
                    "data": {"output": str(output_file)},
                    "error": None,
                    "warnings": [
                        {
                            "level": "warning",
                            "code": "GONGWEN-NEEDS-REVIEW",
                            "message": expected_message,
                            "location": "",
                        }
                    ],
                }
                return subprocess.CompletedProcess([str(binary_path), *args], 0, stdout=json.dumps(payload), stderr="")
            return subprocess.CompletedProcess(
                [str(binary_path), *args],
                0,
                stdout="转换成功\n",
                stderr=f"警告 [GONGWEN-NEEDS-REVIEW]: {expected_message}\n",
            )

        output_file = Path(args[args.index("--output") + 1])
        output_text = (
            f"# {verify_packaged_cli._PYMUPDF_LAYOUT_SMOKE_TEXT}\n"
            if output_file.name.startswith("PyMuPDF Layout")
            else "| name | value |\n| --- | --- |\n| alpha | 1 |\n"
        )
        output_file.write_text(output_text, encoding="utf-8")
        payload = {
            "protocol_version": 3,
            "success": True,
            "command": "convert",
            "data": {"output": str(output_file)},
            "error": None,
        }
        return subprocess.CompletedProcess([str(binary_path), *args], 0, stdout=json.dumps(payload), stderr="")

    monkeypatch.setattr(verify_packaged_cli, "_verify_resource_layout", lambda _binary_dir: None)
    monkeypatch.setattr(verify_packaged_cli, "_write_xlsx", fake_write_xlsx)
    monkeypatch.setattr(verify_packaged_cli, "_run", fake_run)
    monkeypatch.setattr(
        verify_packaged_cli,
        "_run_template_resource_smoke",
        lambda *_args, **_kwargs: tmp_path / "template-id.docx",
    )
    monkeypatch.setattr(
        verify_packaged_cli,
        "_run_multiprocessing_egress_boundary_smoke",
        lambda *_args, **_kwargs: tmp_path / "egress.json",
    )
    monkeypatch.setattr(
        verify_packaged_cli,
        "_run_content_first_contract_smoke",
        lambda _binary_path, *, work_dir: work_dir / "content-first.md",
    )
    monkeypatch.setattr(
        verify_packaged_cli,
        "_run_machine_protocol_smoke",
        lambda _binary_path, *, work_dir: work_dir / "machine-protocol.json",
    )

    exit_code = verify_packaged_cli.main(
        [
            "--binary-dir",
            str(binary_dir),
            "--binary-name",
            binary_name,
            "--successful-warning-smoke",
            str(warning_input),
            "--successful-warning-message",
            expected_message,
        ]
    )

    assert exit_code == 0
    warning_calls = [call for call in calls if "--optimization" in call]
    assert len(warning_calls) == 2
    assert "--json" in warning_calls[0]
    assert "--json" not in warning_calls[1]
    assert warning_calls[0][warning_calls[0].index("--optimization") + 1] == "gongwen"
    assert Path(warning_calls[0][warning_calls[0].index("--output") + 1]).name == "warning_json"
    assert Path(warning_calls[1][warning_calls[1].index("--output") + 1]).name == "warning_text"


def test_packaged_cli_successful_warning_smoke_fails_when_code_is_missing(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from scripts.release import verify_packaged_cli

    binary_path = tmp_path / "DocWenCLI.exe"
    binary_path.write_text("placeholder", encoding="utf-8")
    warning_input = tmp_path / "rules.docx"
    warning_input.write_bytes(b"docx")

    def fake_run(binary_path: Path, *args: str, cwd: Path) -> subprocess.CompletedProcess[str]:
        payload = {
            "protocol_version": 3,
            "success": True,
            "command": "convert",
            "data": {"output": str(cwd / "missing.md")},
            "error": None,
            "warnings": [],
        }
        return subprocess.CompletedProcess([str(binary_path), *args], 0, stdout=json.dumps(payload), stderr="")

    monkeypatch.setattr(verify_packaged_cli, "_run", fake_run)

    with pytest.raises(RuntimeError, match="packaged_successful_warning_code_mismatch"):
        verify_packaged_cli._run_optional_successful_warning_smoke(
            binary_path,
            work_dir=tmp_path,
            input_path=warning_input,
            action="gongwen",
            expected_code="GONGWEN-NEEDS-REVIEW",
        )


def test_packaged_cli_successful_warning_smoke_is_documented_and_in_help() -> None:
    proc = subprocess.run(
        [sys.executable, "scripts/release/verify_packaged_cli.py", "--help"],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
    )
    release_doc = Path("docs/packaging.md").read_text(encoding="utf-8")
    gate_doc = Path("docs/testing.md").read_text(encoding="utf-8")

    assert proc.returncode == 0, proc.stderr
    assert "--successful-warning-smoke" in proc.stdout
    assert "--successful-warning-smoke <input>" in release_doc
    assert "JSON structured warning" in gate_doc
    assert "text stderr" in gate_doc
    assert "final artifact 字节一致" in gate_doc

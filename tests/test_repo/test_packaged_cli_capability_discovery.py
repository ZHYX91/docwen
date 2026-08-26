"""Focused contracts for packaged CLI capability and PyMuPDF Layout gates."""

from __future__ import annotations

import json
import subprocess
from pathlib import Path

import pytest

pytestmark = pytest.mark.contract


def _conversion_route(*, available: bool = True) -> dict[str, object]:
    return {
        "id": "layout:pdf:md:",
        "operation": "conversion",
        "source": "pdf",
        "target": "md",
        "action": None,
        "plugin": "layout",
        "available": available,
        "state": "available" if available else "dependency_missing",
        "platforms": ["windows"],
        "platform_supported": True,
        "required_capabilities": ["python.pymupdf4llm"],
        "optional_capabilities": ["python.rapidocr"],
        "limitations": [],
        "options": [],
    }


def _action_route(*, available: bool = True) -> dict[str, object]:
    return {
        "id": "layout:pdf:pdf:split_pdf",
        "operation": "action",
        "source": "pdf",
        "target": "pdf",
        "action": "split_pdf",
        "plugin": "layout",
        "available": available,
        "state": "available" if available else "dependency_missing",
        "platforms": ["windows"],
        "platform_supported": True,
        "required_capabilities": ["python.fitz"],
        "optional_capabilities": [],
        "limitations": [],
        "options": [],
    }


def _capability_payload(
    *,
    routes: list[dict[str, object]] | None = None,
    layout_gate_available: bool = True,
) -> dict[str, object]:
    routes = list(routes) if routes is not None else [_conversion_route(), _action_route()]
    available_routes = sum(route.get("available") is True for route in routes)
    action_routes = sum(route.get("operation") == "action" for route in routes)
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
            "gates": [
                {
                    "id": "python.pymupdf4llm",
                    "kind": "python_module_with_resources",
                    "available": layout_gate_available,
                }
            ],
            "sources": [{"id": "pdf", "category": "layout", "available": True, "routes": routes}],
            "counts": {
                "sources": 1,
                "routes": len(routes),
                "available_routes": available_routes,
                "unavailable_routes": len(routes) - available_routes,
                "actions": action_routes,
            },
        },
        "error": None,
    }


def _doctor_payload(
    *,
    all_ok: bool = True,
    base_check_status: str = "ok",
    gate_available: bool = True,
    unrelated_gate_available: bool = True,
) -> dict[str, object]:
    return {
        "protocol_version": 3,
        "success": True,
        "command": "doctor",
        "data": {
            "checks": [
                {
                    "id": "path.temp_directory",
                    "kind": "path",
                    "label": "Temporary directory",
                    "status": base_check_status,
                    "reason": None if base_check_status == "ok" else "not_writable",
                },
                {
                    "id": "config.load",
                    "kind": "config",
                    "label": "Configuration",
                    "status": base_check_status,
                    "reason": None if base_check_status == "ok" else "config_load_failed",
                },
                {
                    "id": "security.dependency_egress_guard",
                    "kind": "security",
                    "label": "Dependency egress guard",
                    "status": base_check_status,
                    "reason": None if base_check_status == "ok" else "not_enforced",
                },
            ],
            "all_ok": all_ok,
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
                "gates": [
                    {
                        "id": "python.pymupdf4llm",
                        "kind": "python_module_with_resources",
                        "available": gate_available,
                    },
                    {
                        "id": "external_office.word",
                        "kind": "external_office",
                        "available": unrelated_gate_available,
                    },
                ],
            },
        },
        "error": None,
    }


def _optimization_resource(
    *,
    resource_id: str,
    action_name: str,
    scopes: list[str],
    bindings: list[tuple[str, str, str, str]],
) -> dict[str, object]:
    projected_bindings = [
        {
            "scope": scope,
            "route_id": f"optimizer:{source}:{target}:{action_name}",
            "source": source,
            "source_category": source_category,
            "target": target,
            "available": True,
            "state": "available",
        }
        for scope, source, source_category, target in bindings
    ]
    return {
        "id": resource_id,
        "name": resource_id,
        "action_name": action_name,
        "scopes": scopes,
        "available": True,
        "state": "available",
        "bindings": projected_bindings,
    }


def _optimization_payload() -> dict[str, object]:
    resources = [
        _optimization_resource(
            resource_id="gongwen",
            action_name="gongwen",
            scopes=["document_to_md"],
            bindings=[("document_to_md", "docx", "document", "md")],
        ),
        _optimization_resource(
            resource_id="invoice_cn",
            action_name="invoice_cn",
            scopes=["layout_to_md", "image_to_md"],
            bindings=[
                ("layout_to_md", "pdf", "layout", "md"),
                ("layout_to_md", "ofd", "layout", "md"),
                ("image_to_md", "image", "image", "md"),
            ],
        ),
    ]
    bindings: list[object] = []
    for resource in resources:
        resource_bindings = resource["bindings"]
        assert isinstance(resource_bindings, list)
        bindings.extend(resource_bindings)
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


def _completed(payload: dict[str, object]) -> subprocess.CompletedProcess[str]:
    return subprocess.CompletedProcess(
        ["DocWenCLI.exe", "resources", "list", "formats"],
        0,
        stdout=json.dumps(payload),
        stderr="",
    )


def _optimization_completed(payload: dict[str, object]) -> subprocess.CompletedProcess[str]:
    return subprocess.CompletedProcess(
        ["DocWenCLI.exe", "resources", "list", "optimizations"],
        0,
        stdout=json.dumps(payload),
        stderr="",
    )


def _doctor_completed(payload: dict[str, object], *, returncode: int = 0) -> subprocess.CompletedProcess[str]:
    return subprocess.CompletedProcess(
        ["DocWenCLI.exe", "doctor", "--json", "--quiet"],
        returncode,
        stdout=json.dumps(payload),
        stderr="",
    )


def test_packaged_cli_capability_discovery_accepts_complete_runtime_contract(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from scripts.release import verify_packaged_cli

    calls: list[tuple[tuple[str, ...], Path]] = []

    def fake_run(_binary_path: Path, *args: str, cwd: Path) -> subprocess.CompletedProcess[str]:
        calls.append((args, cwd))
        return _completed(_capability_payload())

    monkeypatch.setattr(verify_packaged_cli, "_run", fake_run)

    verify_packaged_cli._verify_capability_discovery(tmp_path / "DocWenCLI.exe", work_dir=tmp_path)

    assert calls == [(("resources", "list", "formats", "--json", "--quiet"), tmp_path)]


def test_packaged_machine_smoke_uses_resource_kind_for_pdf_merge_inputs() -> None:
    from scripts.release import verify_packaged_cli

    source = Path(verify_packaged_cli.__file__).read_text(encoding="utf-8")

    for input_id in ("input.merge.1", "input.merge.2"):
        marker = f'"input_id": "{input_id}"'
        start = source.index(marker)
        block = source[start : start + 420]
        assert '"kind": "resource"' in block
        assert '"role": "source"' in block
        assert '"media_type": "application/pdf"' in block


def test_packaged_cli_optimization_discovery_accepts_manifest_bound_resources(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from scripts.release import verify_packaged_cli

    calls: list[tuple[tuple[str, ...], Path]] = []

    def fake_run(_binary_path: Path, *args: str, cwd: Path) -> subprocess.CompletedProcess[str]:
        calls.append((args, cwd))
        return _optimization_completed(_optimization_payload())

    monkeypatch.setattr(verify_packaged_cli, "_run", fake_run)

    verify_packaged_cli._verify_optimization_discovery(tmp_path / "DocWenCLI.exe", work_dir=tmp_path)

    assert calls == [(("resources", "list", "optimizations", "--json", "--quiet"), tmp_path)]


def test_packaged_cli_optimization_discovery_rejects_incomplete_binding(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from scripts.release import verify_packaged_cli

    payload = _optimization_payload()
    data = payload["data"]
    assert isinstance(data, dict)
    resources = data["resources"]
    assert isinstance(resources, list)
    first = resources[0]
    assert isinstance(first, dict)
    bindings = first["bindings"]
    assert isinstance(bindings, list)
    binding = bindings[0]
    assert isinstance(binding, dict)
    binding.pop("route_id")
    monkeypatch.setattr(
        verify_packaged_cli,
        "_run",
        lambda _binary_path, *args, cwd: _optimization_completed(payload),
    )

    with pytest.raises(RuntimeError, match="binding contract is incomplete"):
        verify_packaged_cli._verify_optimization_discovery(tmp_path / "DocWenCLI.exe", work_dir=tmp_path)


def test_packaged_cli_capability_discovery_rejects_inactive_egress_guard(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from scripts.release import verify_packaged_cli

    payload = _capability_payload()
    data = payload["data"]
    assert isinstance(data, dict)
    security = data["security"]
    assert isinstance(security, dict)
    guard = security["dependency_egress_guard"]
    assert isinstance(guard, dict)
    guard.update({"state": "installed", "active": False})
    monkeypatch.setattr(
        verify_packaged_cli,
        "_run",
        lambda _binary_path, *args, cwd: _completed(payload),
    )

    with pytest.raises(RuntimeError, match="dependency egress guard is not enforced"):
        verify_packaged_cli._verify_capability_discovery(tmp_path / "DocWenCLI.exe", work_dir=tmp_path)


def test_packaged_cli_capability_discovery_rejects_incomplete_route_contract(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from scripts.release import verify_packaged_cli

    incomplete_route = _conversion_route()
    incomplete_route.pop("options")
    monkeypatch.setattr(
        verify_packaged_cli,
        "_run",
        lambda _binary_path, *args, cwd: _completed(_capability_payload(routes=[incomplete_route, _action_route()])),
    )

    with pytest.raises(RuntimeError, match="route contract is incomplete"):
        verify_packaged_cli._verify_capability_discovery(tmp_path / "DocWenCLI.exe", work_dir=tmp_path)


def test_packaged_cli_capability_discovery_rejects_when_all_routes_are_unavailable(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from scripts.release import verify_packaged_cli

    monkeypatch.setattr(
        verify_packaged_cli,
        "_run",
        lambda _binary_path, *args, cwd: _completed(
            _capability_payload(routes=[_conversion_route(available=False), _action_route(available=False)])
        ),
    )

    with pytest.raises(RuntimeError, match="any available packaged routes"):
        verify_packaged_cli._verify_capability_discovery(tmp_path / "DocWenCLI.exe", work_dir=tmp_path)


def test_packaged_cli_capability_discovery_rejects_unavailable_pdf_to_markdown_route(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from scripts.release import verify_packaged_cli

    monkeypatch.setattr(
        verify_packaged_cli,
        "_run",
        lambda _binary_path, *args, cwd: _completed(
            _capability_payload(routes=[_conversion_route(available=False), _action_route()])
        ),
    )

    with pytest.raises(RuntimeError, match="available PDF to Markdown route"):
        verify_packaged_cli._verify_capability_discovery(tmp_path / "DocWenCLI.exe", work_dir=tmp_path)


def test_packaged_cli_capability_discovery_rejects_unavailable_pymupdf_layout_gate(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from scripts.release import verify_packaged_cli

    monkeypatch.setattr(
        verify_packaged_cli,
        "_run",
        lambda _binary_path, *args, cwd: _completed(_capability_payload(layout_gate_available=False)),
    )

    with pytest.raises(RuntimeError, match=r"python\.pymupdf4llm unavailable"):
        verify_packaged_cli._verify_capability_discovery(tmp_path / "DocWenCLI.exe", work_dir=tmp_path)


def test_packaged_cli_doctor_accepts_healthy_base_checks_with_available_layout_gate() -> None:
    from scripts.release import verify_packaged_cli

    verify_packaged_cli._verify_doctor_payload(_doctor_payload())


def test_packaged_cli_doctor_does_not_require_unrelated_host_capabilities() -> None:
    from scripts.release import verify_packaged_cli

    verify_packaged_cli._verify_doctor_payload(_doctor_payload(unrelated_gate_available=False))


def test_packaged_cli_doctor_decodes_unhealthy_nonzero_payload_before_process_status() -> None:
    from scripts.release import verify_packaged_cli

    with pytest.raises(RuntimeError, match="all_ok=true"):
        verify_packaged_cli._load_verified_doctor_payload(
            _doctor_completed(_doctor_payload(all_ok=False), returncode=1)
        )


def test_packaged_cli_doctor_rejects_nonzero_exit_even_with_healthy_payload() -> None:
    from scripts.release import verify_packaged_cli

    with pytest.raises(RuntimeError, match="healthy payload but exited with 1"):
        verify_packaged_cli._load_verified_doctor_payload(_doctor_completed(_doctor_payload(), returncode=1))


@pytest.mark.parametrize(
    ("payload", "expected_error"),
    [
        (_doctor_payload(all_ok=False), "all_ok=true"),
        (_doctor_payload(base_check_status="fail"), "base check unavailable"),
        (_doctor_payload(gate_available=False), "summary reported python.pymupdf4llm unavailable"),
    ],
)
def test_packaged_cli_doctor_fails_closed_on_unhealthy_layout_runtime(
    payload: dict[str, object],
    expected_error: str,
) -> None:
    from scripts.release import verify_packaged_cli

    with pytest.raises(RuntimeError, match=expected_error):
        verify_packaged_cli._verify_doctor_payload(payload)


def test_packaged_cli_pymupdf_layout_smoke_runs_real_pdf_to_markdown_command(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import fitz
    from scripts.release import verify_packaged_cli

    calls: list[tuple[str, ...]] = []

    def fake_run(binary_path: Path, *args: str, cwd: Path) -> subprocess.CompletedProcess[str]:
        del binary_path, cwd
        calls.append(args)
        source = Path(args[1])
        with fitz.open(source) as document:
            assert verify_packaged_cli._PYMUPDF_LAYOUT_SMOKE_TEXT in "".join(page.get_text() for page in document)
        output = Path(args[args.index("--output") + 1])
        output.write_text(f"# {verify_packaged_cli._PYMUPDF_LAYOUT_SMOKE_TEXT}\n", encoding="utf-8")
        return _completed(
            {
                "protocol_version": 3,
                "success": True,
                "command": "convert",
                "data": {"output": str(output)},
                "error": None,
            }
        )

    monkeypatch.setattr(verify_packaged_cli, "_run", fake_run)

    output = verify_packaged_cli._run_pymupdf_layout_smoke(tmp_path / "DocWenCLI.exe", work_dir=tmp_path)

    assert output.read_text(encoding="utf-8").strip() == f"# {verify_packaged_cli._PYMUPDF_LAYOUT_SMOKE_TEXT}"
    assert calls == [
        (
            "convert",
            str(tmp_path / "PyMuPDF Layout 最小验证.pdf"),
            "--to",
            "md",
            "--output",
            str(tmp_path / "PyMuPDF Layout 最小验证.md"),
            "--json",
            "--quiet",
        )
    ]


def test_packaged_cli_pymupdf_layout_smoke_rejects_output_without_expected_text(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from scripts.release import verify_packaged_cli

    def fake_run(binary_path: Path, *args: str, cwd: Path) -> subprocess.CompletedProcess[str]:
        del binary_path, cwd
        output = Path(args[args.index("--output") + 1])
        output.write_text("# unrelated text\n", encoding="utf-8")
        return _completed(
            {
                "protocol_version": 3,
                "success": True,
                "command": "convert",
                "data": {"output": str(output)},
                "error": None,
            }
        )

    monkeypatch.setattr(verify_packaged_cli, "_run", fake_run)

    with pytest.raises(RuntimeError, match="pymupdf_layout_output_missing_expected_text"):
        verify_packaged_cli._run_pymupdf_layout_smoke(tmp_path / "DocWenCLI.exe", work_dir=tmp_path)


def test_packaged_cli_main_keeps_capability_doctor_and_layout_smoke_as_release_gates() -> None:
    source = Path("scripts/release/verify_packaged_cli.py").read_text(encoding="utf-8")
    main_body = source.split("def main(", 1)[1]

    assert "_verify_capability_discovery(binary_path, work_dir=work_dir)" in main_body
    assert "_verify_optimization_discovery(binary_path, work_dir=work_dir)" in main_body
    assert "_run_template_resource_smoke(binary_path, work_dir=work_dir)" in main_body
    assert "_load_verified_doctor_payload(doctor_process)" in main_body
    assert "_run_pymupdf_layout_smoke(binary_path, work_dir=work_dir)" in main_body

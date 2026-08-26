"""Repo checks for release packaging verification scripts."""

from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

import pytest
from tests.support.packaged_contracts import fake_dependency_egress_guard, fake_numbering_payload
from tests.support.release_packaging import (
    use_compact_pymupdf_layout_manifest,
)
from tests.support.release_packaging import (
    write_packaged_common_resources as _write_packaged_common_resources,
)

pytestmark = pytest.mark.unit


@pytest.fixture(autouse=True)
def _use_compact_pymupdf_layout_manifest(monkeypatch: pytest.MonkeyPatch) -> None:
    use_compact_pymupdf_layout_manifest(monkeypatch)


def _fake_doctor_payload() -> dict[str, object]:
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
                    "status": "ok",
                    "reason": None,
                },
                {
                    "id": "config.load",
                    "kind": "config",
                    "label": "Configuration",
                    "status": "ok",
                    "reason": None,
                },
                {
                    "id": "security.dependency_egress_guard",
                    "kind": "security",
                    "label": "Dependency egress guard",
                    "status": "ok",
                    "reason": None,
                },
            ],
            "all_ok": True,
            "capability_summary": {
                "gates": [{"id": "python.pymupdf4llm", "available": True}],
                "security": {"dependency_egress_guard": fake_dependency_egress_guard()},
            },
        },
        "error": None,
    }


def test_packaged_cli_multiprocessing_report_uses_short_basename_in_unicode_fixture(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from scripts.release import verify_packaged_cli

    work_dir = tmp_path / "资料 空格" / ("长路径验证目录" * 8)
    work_dir.mkdir(parents=True)
    observed_report: Path | None = None

    def fake_run(*_args: object, **kwargs: object) -> subprocess.CompletedProcess[str]:
        nonlocal observed_report
        env = kwargs["env"]
        assert isinstance(env, dict)
        observed_report = Path(env["DOCWEN_TEST_MULTIPROCESS_EGRESS_REPORT"])
        observed_report.write_text(
            json.dumps(
                {
                    "parent_guard": fake_dependency_egress_guard(),
                    "parent_audit_probe_blocked": True,
                    "child_exit_code": 0,
                    "child": {
                        "guard": {
                            "state": "not_installed",
                            "bootstrap": "none",
                        },
                        "audit_probe_allowed": True,
                    },
                }
            ),
            encoding="utf-8",
        )
        return subprocess.CompletedProcess([], 0, stdout="", stderr="")

    monkeypatch.setattr(verify_packaged_cli.subprocess, "run", fake_run)

    returned_report = verify_packaged_cli._run_multiprocessing_egress_boundary_smoke(
        tmp_path / "DocWenCLI.exe",
        work_dir=work_dir,
    )

    assert observed_report is not None
    assert observed_report.name == "egress.json"
    assert returned_report == observed_report
    assert returned_report.parent == work_dir


def test_packaged_cli_long_path_fixture_adapts_to_short_temp_roots() -> None:
    from scripts.release import verify_packaged_cli

    temp_root = Path("/tmp")
    work_dir = verify_packaged_cli._build_long_path_work_dir(temp_root)  # pyright: ignore[reportPrivateUsage]
    output = work_dir / verify_packaged_cli._LONG_PATH_OUTPUT_NAME  # pyright: ignore[reportPrivateUsage]

    assert len(str(output)) >= verify_packaged_cli._LONG_PATH_MINIMUM_LENGTH  # pyright: ignore[reportPrivateUsage]
    assert work_dir.is_relative_to(temp_root)
    assert work_dir.parts[len(temp_root.parts)] == "资料 空格"
    assert len(work_dir.parts) - len(temp_root.parts) <= (verify_packaged_cli._LONG_PATH_SEGMENT_LIMIT + 1)  # pyright: ignore[reportPrivateUsage]


def test_machine_protocol_smoke_cleans_process_and_temp_after_failure(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from scripts.release import verify_packaged_cli

    process: subprocess.Popen[bytes] | None = None
    temp_path: Path | None = None

    def failing_impl(
        _binary_path: Path | list[str],
        *,
        work_dir: Path,
        resources: verify_packaged_cli._MachineSmokeResources,  # pyright: ignore[reportPrivateUsage]
    ) -> Path:
        nonlocal process, temp_path
        assert work_dir == tmp_path
        resources.physical_temp = verify_packaged_cli.tempfile.TemporaryDirectory(prefix="dw-cleanup-test-")
        temp_path = Path(resources.physical_temp.name)
        process = subprocess.Popen(
            [sys.executable, "-c", "import time; time.sleep(60)"],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
        resources.process = process
        raise RuntimeError("injected_machine_failure")

    monkeypatch.setattr(verify_packaged_cli, "_run_machine_protocol_smoke_impl", failing_impl)

    with pytest.raises(RuntimeError, match="injected_machine_failure"):
        verify_packaged_cli._run_machine_protocol_smoke(Path("DocWenCLI.exe"), work_dir=tmp_path)

    assert process is not None
    assert process.poll() is not None
    assert process.stdin is not None and process.stdin.closed
    assert process.stdout is not None and process.stdout.closed
    assert process.stderr is not None and process.stderr.closed
    assert temp_path is not None and not temp_path.exists()


def test_packaged_cli_verifier_uses_protocol_3_convert_command(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    from scripts.release import verify_packaged_cli

    binary_dir = tmp_path / "dist"
    binary_dir.mkdir()
    binary_name = "DocWenCLI.exe"
    (binary_dir / binary_name).write_text("placeholder", encoding="utf-8")
    _write_packaged_common_resources(binary_dir)

    calls: list[tuple[str, ...]] = []

    def fake_write_xlsx(path: Path) -> None:
        path.write_bytes(b"xlsx")

    def fake_run(binary_path: Path, *args: str, cwd: Path) -> subprocess.CompletedProcess[str]:
        calls.append(args)
        if args[:1] == ("doctor",):
            payload = _fake_doctor_payload()
        elif args[:2] == ("number", "markdown"):
            payload = fake_numbering_payload(args)
        elif args[:1] == ("convert",):
            output_file = Path(args[args.index("--output") + 1])
            output_file.write_text("| name | value |\n| --- | --- |\n| alpha | 1 |\n", encoding="utf-8")
            payload = {
                "protocol_version": 3,
                "success": True,
                "command": "convert",
                "data": {"output": str(output_file)},
                "error": None,
            }
        else:
            payload = {"protocol_version": 3, "success": False, "command": args[0], "data": {}, "error": None}

        (cwd / "log_home" / "logs").mkdir(parents=True, exist_ok=True)
        (cwd / "log_home" / "logs" / "docwen.log").write_text("ok", encoding="utf-8")
        return subprocess.CompletedProcess(
            [str(binary_path), *args],
            0,
            stdout=json.dumps(payload),
            stderr="",
        )

    monkeypatch.setattr(verify_packaged_cli, "_write_xlsx", fake_write_xlsx)
    monkeypatch.setattr(verify_packaged_cli, "_run", fake_run)
    monkeypatch.setattr(verify_packaged_cli, "_run_machine_protocol_smoke", lambda *_args, **_kwargs: tmp_path)
    monkeypatch.setattr(verify_packaged_cli, "_verify_capability_discovery", lambda *_args, **_kwargs: None)
    monkeypatch.setattr(verify_packaged_cli, "_verify_optimization_discovery", lambda *_args, **_kwargs: None)
    monkeypatch.setattr(verify_packaged_cli, "_run_template_resource_smoke", lambda *_args, **_kwargs: tmp_path)
    monkeypatch.setattr(
        verify_packaged_cli,
        "_run_multiprocessing_egress_boundary_smoke",
        lambda *_args, **_kwargs: tmp_path / "egress.json",
    )
    monkeypatch.setattr(verify_packaged_cli, "_run_pymupdf_layout_smoke", lambda *_args, **_kwargs: tmp_path)
    monkeypatch.setattr(verify_packaged_cli, "_run_content_first_contract_smoke", lambda *_args, **_kwargs: tmp_path)

    exit_code = verify_packaged_cli.main(["--binary-dir", str(binary_dir), "--binary-name", binary_name])

    assert exit_code == 0
    assert any(call[:1] == ("doctor",) for call in calls)
    assert any(call[:1] == ("convert",) for call in calls)
    assert not any("--ocr" in call for call in calls)
    assert not any(call[:1] == ("run",) for call in calls)


def test_packaged_cli_verifier_runs_optional_ocr_smoke(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    from scripts.release import verify_packaged_cli

    binary_dir = tmp_path / "dist"
    binary_dir.mkdir()
    binary_name = "DocWenCLI.exe"
    (binary_dir / binary_name).write_text("placeholder", encoding="utf-8")
    _write_packaged_common_resources(binary_dir)

    calls: list[tuple[str, ...]] = []

    def fake_write_xlsx(path: Path) -> None:
        path.write_bytes(b"xlsx")

    def fake_write_ocr_png(path: Path) -> None:
        path.write_bytes(b"png")

    def fake_run(binary_path: Path, *args: str, cwd: Path) -> subprocess.CompletedProcess[str]:
        calls.append(args)
        if args[:1] == ("doctor",):
            payload = _fake_doctor_payload()
        elif args[:2] == ("number", "markdown"):
            payload = fake_numbering_payload(args)
        elif args[:1] == ("convert",) and "--ocr" in args:
            output_file = Path(args[args.index("--output") + 1])
            output_file.write_text("> HELLO DOCWEN OCR\n", encoding="utf-8")
            payload = {
                "protocol_version": 3,
                "success": True,
                "command": "convert",
                "data": {"output": str(output_file)},
                "error": None,
            }
        elif args[:1] == ("convert",):
            output_file = Path(args[args.index("--output") + 1])
            output_file.write_text("| name | value |\n| --- | --- |\n| alpha | 1 |\n", encoding="utf-8")
            payload = {
                "protocol_version": 3,
                "success": True,
                "command": "convert",
                "data": {"output": str(output_file)},
                "error": None,
            }
        else:
            payload = {"protocol_version": 3, "success": False, "command": args[0], "data": {}, "error": None}

        (cwd / "log_home" / "logs").mkdir(parents=True, exist_ok=True)
        (cwd / "log_home" / "logs" / "docwen.log").write_text("ok", encoding="utf-8")
        return subprocess.CompletedProcess(
            [str(binary_path), *args],
            0,
            stdout=json.dumps(payload),
            stderr="",
        )

    monkeypatch.setattr(verify_packaged_cli, "_write_xlsx", fake_write_xlsx)
    monkeypatch.setattr(verify_packaged_cli, "_write_ocr_png", fake_write_ocr_png)
    monkeypatch.setattr(verify_packaged_cli, "_run", fake_run)
    monkeypatch.setattr(verify_packaged_cli, "_run_machine_protocol_smoke", lambda *_args, **_kwargs: tmp_path)
    monkeypatch.setattr(verify_packaged_cli, "_verify_capability_discovery", lambda *_args, **_kwargs: None)
    monkeypatch.setattr(verify_packaged_cli, "_verify_optimization_discovery", lambda *_args, **_kwargs: None)
    monkeypatch.setattr(verify_packaged_cli, "_run_template_resource_smoke", lambda *_args, **_kwargs: tmp_path)
    monkeypatch.setattr(
        verify_packaged_cli,
        "_run_multiprocessing_egress_boundary_smoke",
        lambda *_args, **_kwargs: tmp_path / "egress.json",
    )
    monkeypatch.setattr(verify_packaged_cli, "_run_pymupdf_layout_smoke", lambda *_args, **_kwargs: tmp_path)
    monkeypatch.setattr(verify_packaged_cli, "_run_content_first_contract_smoke", lambda *_args, **_kwargs: tmp_path)

    exit_code = verify_packaged_cli.main(["--binary-dir", str(binary_dir), "--binary-name", binary_name, "--ocr-smoke"])

    assert exit_code == 0
    ocr_calls = [call for call in calls if "--ocr" in call]
    assert ocr_calls == [
        (
            "convert",
            str(Path(ocr_calls[0][1])),
            "--to",
            "md",
            "--output",
            str(Path(ocr_calls[0][5])),
            "--ocr",
            "--ocr-placement",
            "main_md",
            "--image-mode",
            "file",
            "--json",
            "--quiet",
        )
    ]


def test_packaged_cli_optional_ocr_smoke_fails_when_text_is_missing(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from scripts.release import verify_packaged_cli

    binary_path = tmp_path / "DocWenCLI.exe"
    binary_path.write_text("placeholder", encoding="utf-8")

    def fake_write_ocr_png(path: Path) -> None:
        path.write_bytes(b"png")

    def fake_run(binary_path: Path, *args: str, cwd: Path) -> subprocess.CompletedProcess[str]:
        output_file = cwd / "sample_ocr.md"
        output_file.write_text("> unreadable\n", encoding="utf-8")
        payload = {
            "protocol_version": 3,
            "success": True,
            "command": "convert",
            "data": {"output": str(output_file)},
            "error": None,
        }
        return subprocess.CompletedProcess([str(binary_path), *args], 0, stdout=json.dumps(payload), stderr="")

    monkeypatch.setattr(verify_packaged_cli, "_write_ocr_png", fake_write_ocr_png)
    monkeypatch.setattr(verify_packaged_cli, "_run", fake_run)

    with pytest.raises(RuntimeError, match="packaged_ocr_output_missing_expected_text"):
        verify_packaged_cli._run_optional_ocr_smoke(binary_path, work_dir=tmp_path)


def test_packaged_cli_verifier_source_does_not_call_old_run_command() -> None:
    source = Path("scripts/release/verify_packaged_cli.py").read_text(encoding="utf-8")

    assert '"run"' not in source
    assert '"convert"' in source
    assert "_run_pymupdf_layout_smoke(binary_path, work_dir=work_dir)" in source

"""Repo checks for the packaged GUI verification script."""

from __future__ import annotations

import json
import os
import subprocess
from pathlib import Path

import pytest
from tests.support.release_packaging import (
    use_compact_pymupdf_layout_manifest,
)
from tests.support.release_packaging import (
    write_packaged_common_resources as _write_packaged_common_resources,
)
from tests.support.release_packaging import (
    write_packaged_gui_assets as _write_packaged_gui_assets,
)

pytestmark = pytest.mark.unit


@pytest.fixture(autouse=True)
def _use_compact_pymupdf_layout_manifest(monkeypatch: pytest.MonkeyPatch) -> None:
    use_compact_pymupdf_layout_manifest(monkeypatch)


def test_packaged_gui_windows_process_snapshot_parses_tasklist(monkeypatch: pytest.MonkeyPatch) -> None:
    from scripts.release import verify_packaged_gui

    result = subprocess.CompletedProcess(
        [],
        0,
        stdout='"DocWen.exe","42"\n"notepad.exe","43"\n"soffice.bin","44"\n',
        stderr="",
    )
    monkeypatch.setattr(verify_packaged_gui.subprocess, "run", lambda *_args, **_kwargs: result)

    assert verify_packaged_gui._snapshot_windows_relevant_processes() == {
        42: "docwen.exe",
        44: "soffice.bin",
    }


def test_packaged_gui_windows_process_snapshot_falls_back_to_get_process(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from scripts.release import verify_packaged_gui

    calls: list[list[str]] = []

    def fake_run(args: list[str], **_kwargs: object) -> subprocess.CompletedProcess[str]:
        calls.append(args)
        if args[0] == "tasklist":
            return subprocess.CompletedProcess(args, 1, stdout="", stderr="ERROR: Access denied")
        assert args[0] == "powershell.exe"
        return subprocess.CompletedProcess(
            args,
            0,
            stdout="42\tDocWen\n43\tnotepad\n44\tsoffice.bin\n",
            stderr="",
        )

    monkeypatch.setattr(verify_packaged_gui.subprocess, "run", fake_run)

    assert verify_packaged_gui._snapshot_windows_relevant_processes() == {
        42: "docwen.exe",
        44: "soffice.bin",
    }
    assert [call[0] for call in calls] == ["tasklist", "powershell.exe"]


def test_packaged_gui_verifier_checks_resources_before_launch(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    from scripts.release import verify_packaged_gui

    binary_dir = tmp_path / "dist"
    binary_dir.mkdir()
    binary_name = "DocWen.exe"
    (binary_dir / binary_name).write_text("placeholder", encoding="utf-8")
    _write_packaged_common_resources(binary_dir)
    _write_packaged_gui_assets(binary_dir)

    calls: list[Path] = []

    def fake_run(binary_path: Path, *args: str, cwd: Path) -> subprocess.CompletedProcess[str]:
        calls.append(binary_path)
        (cwd / "log_home" / "logs").mkdir(parents=True, exist_ok=True)
        (cwd / "log_home" / "logs" / "docwen.log").write_text("ok", encoding="utf-8")
        return subprocess.CompletedProcess([str(binary_path), *args], 0, stdout="", stderr="")

    monkeypatch.setattr(verify_packaged_gui, "_run", fake_run)

    exit_code = verify_packaged_gui.main(["--binary-dir", str(binary_dir), "--binary-name", binary_name])

    assert exit_code == 0
    assert calls == [binary_dir / binary_name]


def test_packaged_gui_settings_archive_contract_tracks_dialog_pages() -> None:
    from scripts.release import verify_packaged_gui

    from docwen_gui.widgets.settings.dialog import _TAB_SPECS, TAB_KEYS

    expected = {f"docwen_gui.widgets.settings.{spec.module_name}" for spec in _TAB_SPECS.values()}

    assert expected == verify_packaged_gui._REQUIRED_SETTINGS_PAGE_MODULES
    assert tuple(TAB_KEYS) == verify_packaged_gui._REQUIRED_SETTINGS_TAB_KEYS


def test_packaged_gui_settings_archive_rejects_a_missing_page(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from scripts.release import verify_packaged_gui

    binary_path = tmp_path / "DocWen.exe"
    binary_path.write_bytes(b"placeholder")
    missing_module = "docwen_gui.widgets.settings.general_tab"
    monkeypatch.setattr(
        verify_packaged_gui,
        "_read_pyinstaller_module_names",
        lambda _binary: verify_packaged_gui._REQUIRED_SETTINGS_PAGE_MODULES - {missing_module},
    )

    with pytest.raises(RuntimeError, match="packaged_gui_settings_modules_missing"):
        verify_packaged_gui._verify_settings_page_archive(binary_path)


def test_packaged_gui_verifier_runs_settings_page_smoke(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from scripts.release import verify_packaged_gui

    from docwen_gui.widgets.settings.dialog import TAB_KEYS

    binary_dir = tmp_path / "dist"
    binary_dir.mkdir()
    binary_name = "DocWen.exe"
    (binary_dir / binary_name).write_text("placeholder", encoding="utf-8")
    _write_packaged_common_resources(binary_dir)
    _write_packaged_gui_assets(binary_dir)
    deploy_files_before = {path.relative_to(binary_dir).as_posix() for path in binary_dir.rglob("*") if path.is_file()}
    archive_checks: list[Path] = []
    run_directories: list[Path] = []
    monkeypatch.setattr(
        verify_packaged_gui,
        "_verify_settings_page_archive",
        lambda binary: archive_checks.append(binary),
    )

    def fake_run(binary_path: Path, *args: str, cwd: Path) -> subprocess.CompletedProcess[str]:
        run_directories.append(cwd)
        assert binary_dir not in (cwd, *cwd.parents)
        report_path = Path(os.environ["DOCWEN_GUI_TEST_SETTINGS_REPORT"])
        report_path.write_text(
            json.dumps(
                {
                    "success": True,
                    "expectedTabs": TAB_KEYS,
                    "loadedTabs": TAB_KEYS,
                    "failedTabs": [],
                    "missingTabs": [],
                    "unexpectedTabs": [],
                    "pageObjectNames": dict.fromkeys(TAB_KEYS, "settingsTabRoot"),
                    "error": None,
                }
            ),
            encoding="utf-8",
        )
        (cwd / "log_home" / "logs").mkdir(parents=True, exist_ok=True)
        (cwd / "log_home" / "logs" / "docwen.log").write_text("ok", encoding="utf-8")
        return subprocess.CompletedProcess([str(binary_path), *args], 0, stdout="", stderr="")

    monkeypatch.setattr(verify_packaged_gui, "_run", fake_run)

    exit_code = verify_packaged_gui.main(
        ["--binary-dir", str(binary_dir), "--binary-name", binary_name, "--settings-smoke"]
    )

    assert exit_code == 0
    assert archive_checks == [binary_dir / binary_name]
    assert "DOCWEN_GUI_TEST_SETTINGS_REPORT" not in os.environ
    assert len(run_directories) == 1
    assert run_directories[0].exists() is False
    assert {
        path.relative_to(binary_dir).as_posix() for path in binary_dir.rglob("*") if path.is_file()
    } == deploy_files_before


def test_packaged_gui_verifier_runs_optional_notification_smoke(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from scripts.release import verify_packaged_gui

    binary_dir = tmp_path / "dist"
    binary_dir.mkdir()
    binary_name = "DocWen.exe"
    (binary_dir / binary_name).write_text("placeholder", encoding="utf-8")
    _write_packaged_common_resources(binary_dir)
    _write_packaged_gui_assets(binary_dir)
    monkeypatch.delenv("DOCWEN_GUI_TEST_NOTIFICATION_REPORT", raising=False)
    captured_report: dict[str, object] = {}

    def fake_run(binary_path: Path, *args: str, cwd: Path) -> subprocess.CompletedProcess[str]:
        report_path = Path(os.environ["DOCWEN_GUI_TEST_NOTIFICATION_REPORT"])
        captured_report.update(
            {
                "isSystemTrayAvailable": True,
                "supportsMessages": True,
                "defaultTrayIconPresent": False,
                "probeCreatedTrayIcon": True,
                "hasTrayIcon": True,
                "showMessageCalled": True,
                "error": None,
            }
        )
        report_path.write_text(json.dumps(captured_report), encoding="utf-8")
        (cwd / "log_home" / "logs").mkdir(parents=True, exist_ok=True)
        (cwd / "log_home" / "logs" / "docwen.log").write_text("ok", encoding="utf-8")
        return subprocess.CompletedProcess([str(binary_path), *args], 0, stdout="", stderr="")

    monkeypatch.setattr(verify_packaged_gui, "_run", fake_run)
    exit_code = verify_packaged_gui.main(
        ["--binary-dir", str(binary_dir), "--binary-name", binary_name, "--notification-smoke"]
    )
    assert exit_code == 0
    assert "DOCWEN_GUI_TEST_NOTIFICATION_REPORT" not in os.environ
    assert captured_report["showMessageCalled"] is True
    assert not (binary_dir / "notification_smoke.json").exists()


def test_packaged_gui_verifier_runs_optional_ocr_smoke(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    from scripts.release import verify_packaged_gui

    binary_dir = tmp_path / "dist"
    binary_dir.mkdir()
    binary_name = "DocWen.exe"
    (binary_dir / binary_name).write_text("placeholder", encoding="utf-8")
    _write_packaged_common_resources(binary_dir)
    _write_packaged_gui_assets(binary_dir)
    monkeypatch.delenv("DOCWEN_GUI_TEST_OCR_REPORT", raising=False)
    monkeypatch.delenv("DOCWEN_GUI_TEST_OCR_INPUT", raising=False)
    captured_outputs: list[str] = []

    def fake_write_ocr_png(path: Path) -> None:
        path.write_bytes(b"png")

    def fake_run(binary_path: Path, *args: str, cwd: Path) -> subprocess.CompletedProcess[str]:
        report_path = Path(os.environ["DOCWEN_GUI_TEST_OCR_REPORT"])
        output_dir = Path(os.environ["DOCWEN_GUI_TEST_OCR_OUTPUT_DIR"])
        output_dir.mkdir(parents=True, exist_ok=True)
        output_path = output_dir / "sample_gui_ocr.md"
        sidecar_path = output_dir / "sample_gui_ocr_ocr.md"
        output_path.write_text("> HELLO\n> DOCWEN OCR\n", encoding="utf-8")
        captured_outputs.append(output_path.read_text(encoding="utf-8"))
        report_path.write_text(
            json.dumps(
                {
                    "success": False,
                    "status": "completed",
                    "inputPath": os.environ["DOCWEN_GUI_TEST_OCR_INPUT"],
                    "outputPath": str(output_path),
                    "sidecarPath": str(sidecar_path),
                    "primaryOutputExists": True,
                    "sidecarOutputExists": False,
                    "primaryReferencesSidecar": False,
                    "primaryContainsExpectedText": False,
                    "sidecarContainsExpectedText": False,
                    "expectedText": "HELLO DOCWEN OCR",
                    "error": None,
                }
            ),
            encoding="utf-8",
        )
        (cwd / "log_home" / "logs").mkdir(parents=True, exist_ok=True)
        (cwd / "log_home" / "logs" / "docwen.log").write_text("ok", encoding="utf-8")
        return subprocess.CompletedProcess([str(binary_path), *args], 0, stdout="", stderr="")

    monkeypatch.setattr(verify_packaged_gui, "_write_ocr_png", fake_write_ocr_png)
    monkeypatch.setattr(verify_packaged_gui, "_run", fake_run)

    exit_code = verify_packaged_gui.main(["--binary-dir", str(binary_dir), "--binary-name", binary_name, "--ocr-smoke"])

    assert exit_code == 0
    assert "DOCWEN_GUI_TEST_OCR_REPORT" not in os.environ
    assert "DOCWEN_GUI_TEST_OCR_INPUT" not in os.environ
    assert captured_outputs == ["> HELLO\n> DOCWEN OCR\n"]
    assert not (binary_dir / "gui_ocr_smoke.json").exists()


def test_packaged_gui_verifier_runs_optional_ipc_smoke(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    from scripts.release import verify_packaged_gui

    binary_dir = tmp_path / "dist"
    binary_dir.mkdir()
    binary_name = "DocWen.exe"
    (binary_dir / binary_name).write_text("placeholder", encoding="utf-8")
    _write_packaged_common_resources(binary_dir)
    _write_packaged_gui_assets(binary_dir)

    calls: list[Path] = []

    def fake_run_ipc_smoke(
        binary_path: Path,
        *,
        cwd: Path,
        binary_dir: Path | None = None,
    ) -> subprocess.CompletedProcess[str]:
        calls.append(binary_path)
        assert binary_dir == binary_path.parent
        input_path = cwd / "ipc_smoke_input.md"
        input_path.write_text("# IPC smoke\n", encoding="utf-8")
        report_path = cwd / "ipc_smoke.json"
        report_path.write_text(
            json.dumps(
                {
                    "success": True,
                    "status": "completed",
                    "error": None,
                    "expectedFile": str(input_path),
                    "receivedFiles": [str(input_path)],
                    "expectedReceived": True,
                    "files": [str(input_path)],
                    "expectedInFiles": True,
                    "selectedFile": str(input_path),
                    "activationCount": 1,
                    "elapsedMs": 123,
                }
            ),
            encoding="utf-8",
        )
        (cwd / "log_home" / "logs").mkdir(parents=True, exist_ok=True)
        (cwd / "log_home" / "logs" / "docwen.log").write_text("ok", encoding="utf-8")
        return subprocess.CompletedProcess([str(binary_path), "--ipc-smoke-primary"], 0, stdout="", stderr="")

    monkeypatch.setattr(verify_packaged_gui, "_run_ipc_smoke", fake_run_ipc_smoke)

    exit_code = verify_packaged_gui.main(["--binary-dir", str(binary_dir), "--binary-name", binary_name, "--ipc-smoke"])

    assert exit_code == 0
    assert calls == [binary_dir / binary_name]


def test_packaged_gui_ipc_smoke_must_run_without_autoclose_smokes(tmp_path: Path) -> None:
    from scripts.release import verify_packaged_gui

    binary_dir = tmp_path / "dist"
    binary_dir.mkdir()
    binary_name = "DocWen.exe"
    (binary_dir / binary_name).write_text("placeholder", encoding="utf-8")
    _write_packaged_common_resources(binary_dir)
    _write_packaged_gui_assets(binary_dir)

    with pytest.raises(SystemExit):
        verify_packaged_gui.main(
            [
                "--binary-dir",
                str(binary_dir),
                "--binary-name",
                binary_name,
                "--ipc-smoke",
                "--ocr-smoke",
            ]
        )


def test_packaged_gui_notification_smoke_fails_without_report(tmp_path: Path) -> None:
    from scripts.release import verify_packaged_gui

    with pytest.raises(RuntimeError, match="packaged_gui_notification_report_missing"):
        verify_packaged_gui._verify_notification_smoke_report(tmp_path / "missing.json")


def test_packaged_gui_ocr_smoke_fails_without_report(tmp_path: Path) -> None:
    from scripts.release import verify_packaged_gui

    with pytest.raises(RuntimeError, match="packaged_gui_ocr_report_missing"):
        verify_packaged_gui._verify_ocr_smoke_report(tmp_path / "missing.json")


def test_packaged_gui_ipc_smoke_fails_without_report(tmp_path: Path) -> None:
    from scripts.release import verify_packaged_gui

    with pytest.raises(RuntimeError, match="packaged_gui_ipc_report_missing"):
        verify_packaged_gui._verify_ipc_smoke_report(tmp_path / "missing.json")


@pytest.mark.parametrize(
    ("field", "replacement"),
    [
        ("success", 1),
        ("success", "true"),
        ("expectedReceived", 1),
        ("expectedReceived", "true"),
        ("expectedInFiles", 1),
        ("expectedInFiles", "true"),
        ("activationCount", True),
        ("activationCount", "1"),
        ("activationCount", 1.0),
        ("activationCount", 0),
    ],
)
def test_packaged_gui_ipc_smoke_rejects_truthy_schema_drift(
    tmp_path: Path,
    field: str,
    replacement: object,
) -> None:
    from scripts.release import verify_packaged_gui

    input_path = tmp_path / "ipc_smoke_input.md"
    input_path.write_text("# IPC smoke\n", encoding="utf-8")
    report: dict[str, object] = {
        "success": True,
        "status": "completed",
        "error": None,
        "expectedFile": str(input_path),
        "receivedFiles": [str(input_path)],
        "expectedReceived": True,
        "files": [str(input_path)],
        "expectedInFiles": True,
        "selectedFile": str(input_path),
        "activationCount": 1,
        "elapsedMs": 123,
    }
    report[field] = replacement
    report_path = tmp_path / "ipc_smoke.json"
    report_path.write_text(json.dumps(report), encoding="utf-8")

    with pytest.raises(RuntimeError, match="packaged_gui_ipc_"):
        verify_packaged_gui._verify_ipc_smoke_report(report_path)


def test_packaged_gui_ipc_smoke_requires_settings_handshake() -> None:
    from scripts.release import verify_packaged_gui

    valid = {
        "running": True,
        "control_ready": True,
        "supported_actions": ["status", "activate", "open", "open_settings"],
        "settings_sections": ["proofread"],
    }
    verify_packaged_gui._verify_gui_settings_handshake(valid)

    with pytest.raises(RuntimeError, match="packaged_gui_settings_capability_missing"):
        verify_packaged_gui._verify_gui_settings_handshake({**valid, "supported_actions": ["status", "open"]})
    with pytest.raises(RuntimeError, match="packaged_gui_proofread_settings_section_missing"):
        verify_packaged_gui._verify_gui_settings_handshake({**valid, "settings_sections": []})


@pytest.mark.parametrize("reused", (False, True))
def test_packaged_gui_ipc_smoke_verifies_open_settings_contract(reused: bool) -> None:
    from scripts.release import verify_packaged_gui

    completed = subprocess.CompletedProcess(
        ["DocWenCLI.exe", "gui", "open-settings"],
        0,
        stdout=json.dumps(
            {
                "success": True,
                "command": "gui open-settings",
                "data": {
                    "accepted": True,
                    "running": True,
                    "action": "open_settings",
                    "section": "proofread",
                    "reused": reused,
                },
            }
        ),
        stderr="",
    )

    assert (
        verify_packaged_gui._verify_open_settings_response(
            completed,
            expected_reused=reused,
        )["reused"]
        is reused
    )


def test_packaged_gui_ipc_smoke_rejects_open_settings_contract_drift() -> None:
    from scripts.release import verify_packaged_gui

    completed = subprocess.CompletedProcess(
        ["DocWenCLI.exe", "gui", "open-settings"],
        0,
        stdout=json.dumps(
            {
                "success": True,
                "command": "gui open-settings",
                "data": {
                    "accepted": True,
                    "running": True,
                    "action": "open_settings",
                    "section": "proofread",
                    "reused": False,
                },
            }
        ),
        stderr="",
    )

    with pytest.raises(RuntimeError, match="packaged_gui_open_settings_contract_mismatch"):
        verify_packaged_gui._verify_open_settings_response(completed, expected_reused=True)


def test_packaged_gui_smoke_documents_ipc_boundary() -> None:
    script_source = Path("scripts/release/verify_packaged_gui.py").read_text(encoding="utf-8")
    gui_app_source = Path("packages/apps/gui/src/docwen_gui/app.py").read_text(encoding="utf-8")
    bundle_entry_source = Path("packages/bundle/src/docwen_bundle/gui_entry.py").read_text(encoding="utf-8")
    release_doc = Path("docs/packaging.md").read_text(encoding="utf-8")
    gate_doc = Path("docs/testing.md").read_text(encoding="utf-8")

    assert "DOCWEN_GUI_TEST_AUTOCLOSE_MS" in script_source
    assert "DOCWEN_GUI_TEST_NOTIFICATION_REPORT" in script_source
    assert "DOCWEN_GUI_TEST_OCR_REPORT" in script_source
    assert "DOCWEN_GUI_TEST_IPC_REPORT" in script_source
    assert "--ipc-smoke" in script_source
    assert "DOCWEN_GUI_TEST_NOTIFICATION_REPORT" in gui_app_source
    assert "DOCWEN_GUI_TEST_OCR_REPORT" in gui_app_source
    assert "DOCWEN_GUI_TEST_IPC_REPORT" in gui_app_source
    assert "_schedule_test_notification_report" in bundle_entry_source
    assert "_schedule_test_ocr_report" in bundle_entry_source
    assert "_schedule_test_ipc_report" in bundle_entry_source
    assert "--notification-smoke" in release_doc
    assert "--ocr-smoke" in release_doc
    assert "--ipc-smoke" in release_doc
    assert "--notification-smoke" in gate_doc
    assert "--ocr-smoke" in gate_doc
    assert "--ipc-smoke" in gate_doc
    assert "不证明 Windows 通知中心可见性" in release_doc
    assert "禁用 IPC" in release_doc
    assert "不证明单实例锁、二次启动文件投递或窗口激活" in release_doc
    assert "可选 `--ipc-smoke`" in release_doc
    assert "The default packaged GUI smoke uses `DOCWEN_GUI_TEST_AUTOCLOSE_MS`" in gate_doc

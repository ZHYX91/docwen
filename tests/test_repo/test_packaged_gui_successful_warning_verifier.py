"""Direct guards for the packaged GUI successful-warning release gate."""

from __future__ import annotations

import json
import subprocess
from pathlib import Path

import pytest

pytestmark = pytest.mark.unit

WARNING = "缺少必需字段：成文日期、发文机关署名；识别提示：存在低置信度识别"


def _document_node_output(parent: Path, *, source_stem: str = "rules") -> Path:
    root_name = f"{source_stem}_20260824_120102_fromDocx"
    return parent / root_name / f"{root_name}.md"


def _write_valid_report(
    report_path: Path,
    *,
    input_path: Path,
    output_path: Path,
    screenshot_path: Path,
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text("---\ntitle: test\n---\n", encoding="utf-8")
    screenshot_path.parent.mkdir(parents=True, exist_ok=True)
    screenshot_path.write_bytes(b"\x89PNG\r\n\x1a\npackaged-gui-warning")
    row_screenshot_path = screenshot_path.with_name(f"{screenshot_path.stem}_warning_row{screenshot_path.suffix}")
    row_screenshot_path.write_bytes(b"\x89PNG\r\n\x1a\npackaged-gui-warning-row")
    report_path.write_text(
        json.dumps(
            {
                "success": True,
                "status": "completed",
                "inputPath": str(input_path),
                "outputPath": str(output_path),
                "outputExists": True,
                "outputBytes": output_path.stat().st_size,
                "surface": "action",
                "actionName": "gongwen",
                "expectedWarningMessage": WARNING,
                "historyRows": [
                    {"message": "Completed", "messageType": "warning", "filePath": str(output_path)},
                    {"message": WARNING, "messageType": "warning", "filePath": str(output_path)},
                ],
                "warningHistoryIndex": 1,
                "warningRowTone": "warning",
                "warningRowTooltip": WARNING,
                "warningRowVisible": True,
                "warningRowScreenshotPath": str(row_screenshot_path),
                "warningRowScreenshotSaved": True,
                "warningRowScreenshotBytes": row_screenshot_path.stat().st_size,
                "warningRowScreenshotWidth": 600,
                "warningRowScreenshotHeight": 90,
                "taskSummary": {
                    "state": "success",
                    "tone": "warning",
                    "completedCount": 1,
                    "totalCount": 1,
                    "failedCount": 0,
                    "navigatePath": str(output_path),
                },
                "statusSource": "task",
                "statusTone": "warning",
                "infoAreaVisible": True,
                "screenshotPath": str(screenshot_path),
                "screenshotSaved": True,
                "screenshotBytes": screenshot_path.stat().st_size,
                "screenshotWidth": 640,
                "screenshotHeight": 320,
                "error": "",
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )


def test_successful_warning_runner_drives_gongwen_action_and_retains_outputs(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from scripts.release import verify_packaged_gui

    binary = tmp_path / "DocWen.exe"
    binary.write_bytes(b"exe")
    source = tmp_path / "rules.docx"
    source.write_bytes(b"PK\x03\x04")
    observed: dict[str, object] = {}

    def fake_run(
        binary_path: Path,
        *args: str,
        cwd: Path,
        env: dict[str, str],
        timeout: int,
    ) -> subprocess.CompletedProcess[str]:
        observed.update(binary=binary_path, args=args, cwd=cwd, env=env.copy(), timeout=timeout)
        _write_valid_report(
            Path(env["DOCWEN_GUI_TEST_CONVERSION_REPORT"]),
            input_path=source.resolve(),
            output_path=_document_node_output(Path(env["DOCWEN_GUI_TEST_CONVERSION_OUTPUT_DIR"])),
            screenshot_path=Path(env["DOCWEN_GUI_TEST_CONVERSION_SCREENSHOT"]),
        )
        return subprocess.CompletedProcess([str(binary_path)], 0, stdout="", stderr="")

    monkeypatch.setattr(verify_packaged_gui, "_run_with_env", fake_run)
    monkeypatch.setattr(verify_packaged_gui, "_snapshot_relevant_processes", dict)
    monkeypatch.setattr(verify_packaged_gui, "_wait_for_no_new_relevant_processes", lambda _before: None)

    output, screenshot, report, proc = verify_packaged_gui._run_successful_warning_smoke(
        binary,
        cwd=tmp_path,
        input_path=source,
        expected_message=WARNING,
    )

    env = observed["env"]
    assert isinstance(env, dict)
    assert env["DOCWEN_GUI_TEST_CONVERSION_SURFACE"] == "action"
    assert env["DOCWEN_GUI_TEST_CONVERSION_ACTION"] == "gongwen"
    assert env["DOCWEN_GUI_TEST_CONVERSION_TARGET"] == "md"
    assert env["DOCWEN_GUI_TEST_CONVERSION_EXPECT_WARNING"] == WARNING
    assert observed["timeout"] == 150
    assert output.name == "rules_20260824_120102_fromDocx.md"
    assert screenshot.name == "successful_warning_info_area.png"
    assert report.name == "successful_warning_smoke.json"
    assert proc.returncode == 0


def test_successful_warning_report_fails_closed_when_warning_row_is_missing(tmp_path: Path) -> None:
    from scripts.release import verify_packaged_gui

    source = tmp_path / "rules.docx"
    output = _document_node_output(tmp_path)
    screenshot = tmp_path / "info.png"
    report = tmp_path / "report.json"
    source.write_bytes(b"PK\x03\x04")
    _write_valid_report(report, input_path=source, output_path=output, screenshot_path=screenshot)
    payload = json.loads(report.read_text(encoding="utf-8"))
    payload["historyRows"] = []
    report.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")

    with pytest.raises(RuntimeError, match="packaged_gui_successful_warning_history_missing"):
        verify_packaged_gui._verify_successful_warning_smoke_report(
            report,
            screenshot_path=screenshot,
            input_path=source,
            expected_message=WARNING,
        )


def test_successful_warning_report_rejects_empty_gui_output(tmp_path: Path) -> None:
    from scripts.release import verify_packaged_gui

    source = tmp_path / "rules.docx"
    output = _document_node_output(tmp_path)
    screenshot = tmp_path / "info.png"
    report = tmp_path / "report.json"
    source.write_bytes(b"PK\x03\x04")
    _write_valid_report(report, input_path=source, output_path=output, screenshot_path=screenshot)
    output.write_bytes(b"")
    payload = json.loads(report.read_text(encoding="utf-8"))
    payload["outputBytes"] = 0
    report.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")

    with pytest.raises(RuntimeError, match="output_empty_or_untracked"):
        verify_packaged_gui._verify_successful_warning_smoke_report(
            report,
            screenshot_path=screenshot,
            input_path=source,
            expected_message=WARNING,
        )


def test_successful_warning_preflight_uses_packaged_cli_as_canonical_message(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from scripts.release import verify_packaged_gui

    cli = tmp_path / verify_packaged_gui._default_cli_binary_name()
    cli.write_bytes(b"exe")
    source = tmp_path / "rules.docx"
    source.write_bytes(b"PK\x03\x04fixture")
    observed: dict[str, object] = {}

    def fake_run(
        binary_path: Path,
        *args: str,
        cwd: Path,
        env: dict[str, str],
        timeout: int,
    ) -> subprocess.CompletedProcess[str]:
        observed.update(binary=binary_path, args=args, cwd=cwd, env=env, timeout=timeout)
        output_dir = Path(args[args.index("--output") + 1])
        output_dir.mkdir(parents=True, exist_ok=True)
        root_name = "rules_20260824_120102_fromDocx"
        output = output_dir / root_name / f"{root_name}.md"
        output.parent.mkdir()
        output.write_text("# converted\n", encoding="utf-8")
        payload = {
            "protocol_version": 3,
            "success": True,
            "data": {"output": str(output)},
            "warnings": [{"level": "warning", "code": "GONGWEN-NEEDS-REVIEW", "message": WARNING}],
        }
        return subprocess.CompletedProcess([str(binary_path), *args], 0, stdout=json.dumps(payload), stderr="")

    monkeypatch.setattr(verify_packaged_gui, "_run_with_env", fake_run)

    message, output = verify_packaged_gui._probe_successful_warning_contract(
        cli,
        cwd=tmp_path,
        input_path=source,
        action="gongwen",
        expected_code="GONGWEN-NEEDS-REVIEW",
    )

    assert message == WARNING
    assert output.read_text(encoding="utf-8") == "# converted\n"
    assert observed["binary"] == cli
    args = observed["args"]
    assert isinstance(args, tuple)
    assert args[:3] == ("convert", str(source.resolve()), "--to")
    assert (args[4], args[5]) == ("--optimization", "gongwen")
    assert (args[8], args[9]) == ("--lang", "zh_CN")


@pytest.mark.parametrize(
    ("warnings", "output_text", "expected_error"),
    [
        ([], "# converted\n", "warning_count_unexpected"),
        (
            [
                {"level": "warning", "code": "GONGWEN-NEEDS-REVIEW", "message": WARNING},
                {"level": "warning", "code": "GONGWEN-NEEDS-REVIEW", "message": WARNING},
            ],
            "# converted\n",
            "warning_count_unexpected",
        ),
        ([{"level": "warning", "code": "GONGWEN-NEEDS-REVIEW", "message": ""}], "# converted\n", "message_missing"),
        ([{"level": "warning", "code": "GONGWEN-NEEDS-REVIEW", "message": WARNING}], "", "output_empty"),
    ],
)
def test_successful_warning_preflight_fails_closed(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    warnings: list[dict[str, str]],
    output_text: str,
    expected_error: str,
) -> None:
    from scripts.release import verify_packaged_gui

    cli = tmp_path / verify_packaged_gui._default_cli_binary_name()
    cli.write_bytes(b"exe")
    source = tmp_path / "rules.docx"
    source.write_bytes(b"PK\x03\x04fixture")

    def fake_run(
        _binary_path: Path,
        *args: str,
        cwd: Path,
        env: dict[str, str],
        timeout: int,
    ) -> subprocess.CompletedProcess[str]:
        del cwd, env, timeout
        output_dir = Path(args[args.index("--output") + 1])
        output_dir.mkdir(parents=True, exist_ok=True)
        root_name = "rules_20260824_120102_fromDocx"
        output = output_dir / root_name / f"{root_name}.md"
        output.parent.mkdir()
        output.write_text(output_text, encoding="utf-8")
        payload = {
            "protocol_version": 3,
            "success": True,
            "data": {"output": str(output)},
            "warnings": warnings,
        }
        return subprocess.CompletedProcess([str(cli), *args], 0, stdout=json.dumps(payload), stderr="")

    monkeypatch.setattr(verify_packaged_gui, "_run_with_env", fake_run)

    with pytest.raises(RuntimeError, match=expected_error):
        verify_packaged_gui._probe_successful_warning_contract(
            cli,
            cwd=tmp_path,
            input_path=source,
            action="gongwen",
            expected_code="GONGWEN-NEEDS-REVIEW",
        )


def test_successful_warning_preflight_rejects_golden_mismatch_before_gui(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from scripts.release import verify_packaged_gui

    cli = tmp_path / verify_packaged_gui._default_cli_binary_name()
    cli.write_bytes(b"exe")
    source = tmp_path / "rules.docx"
    source.write_bytes(b"PK\x03\x04fixture")

    def fake_run(
        _binary_path: Path,
        *args: str,
        cwd: Path,
        env: dict[str, str],
        timeout: int,
    ) -> subprocess.CompletedProcess[str]:
        del args, cwd, env, timeout
        payload = {
            "protocol_version": 3,
            "success": True,
            "warnings": [{"level": "warning", "code": "GONGWEN-NEEDS-REVIEW", "message": "actual"}],
        }
        return subprocess.CompletedProcess([str(cli)], 0, stdout=json.dumps(payload), stderr="")

    monkeypatch.setattr(verify_packaged_gui, "_run_with_env", fake_run)

    with pytest.raises(RuntimeError, match="golden_mismatch"):
        verify_packaged_gui._probe_successful_warning_contract(
            cli,
            cwd=tmp_path,
            input_path=source,
            action="gongwen",
            expected_code="GONGWEN-NEEDS-REVIEW",
            expected_message="golden",
        )


def test_main_wires_successful_warning_smoke_without_notification_overclaim(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from scripts.release import verify_packaged_gui

    binary_dir = tmp_path / "dist"
    binary_dir.mkdir()
    binary = binary_dir / verify_packaged_gui._default_binary_name()
    binary.write_bytes(b"exe")
    cli = binary_dir / verify_packaged_gui._default_cli_binary_name()
    cli.write_bytes(b"exe")
    source = tmp_path / "rules.docx"
    source.write_bytes(b"PK\x03\x04")
    observed: dict[str, object] = {}

    def fake_warning_run(
        binary_path: Path,
        *,
        cwd: Path,
        input_path: Path,
        expected_message: str,
    ) -> tuple[Path, Path, Path, subprocess.CompletedProcess[str]]:
        observed.update(
            binary=binary_path,
            cwd=cwd,
            input=input_path,
            expected_message=expected_message,
        )
        output = cwd / "gui_successful_warning_smoke" / "outputs" / "rules.md"
        screenshot = cwd / "gui_successful_warning_smoke" / "successful_warning_info_area.png"
        report = cwd / "gui_successful_warning_smoke" / "successful_warning_smoke.json"
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text("ok", encoding="utf-8")
        screenshot.write_bytes(b"\x89PNG\r\n\x1a\n")
        report.write_text("{}", encoding="utf-8")
        log_dir = cwd / "log_home" / "logs"
        log_dir.mkdir(parents=True, exist_ok=True)
        (log_dir / "docwen.log").write_text("ok", encoding="utf-8")
        proc = subprocess.CompletedProcess([str(binary_path)], 0, stdout="", stderr="")
        return output, screenshot, report, proc

    def fake_warning_probe(
        cli_path: Path,
        *,
        cwd: Path,
        input_path: Path,
        action: str,
        expected_code: str,
        expected_message: str,
    ) -> tuple[str, Path]:
        observed.update(
            cli=cli_path,
            probe_input=input_path,
            probe_action=action,
            probe_code=expected_code,
            golden_message=expected_message,
        )
        output = cwd / "gui_successful_warning_preflight" / "rules.md"
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text("ok", encoding="utf-8")
        return WARNING, output

    monkeypatch.setattr(verify_packaged_gui, "_verify_resource_layout", lambda _path: None)
    monkeypatch.setattr(verify_packaged_gui, "_probe_successful_warning_contract", fake_warning_probe)
    monkeypatch.setattr(verify_packaged_gui, "_run_successful_warning_smoke", fake_warning_run)

    exit_code = verify_packaged_gui.main(
        [
            "--binary-dir",
            str(binary_dir),
            "--successful-warning-smoke",
            str(source),
            "--successful-warning-message",
            WARNING,
        ]
    )

    assert exit_code == 0
    assert observed["binary"] == binary
    assert observed["cli"] == cli
    assert observed["input"] == source
    assert observed["expected_message"] == WARNING
    assert observed["probe_action"] == "gongwen"
    assert observed["probe_code"] == "GONGWEN-NEEDS-REVIEW"
    assert observed["golden_message"] == WARNING

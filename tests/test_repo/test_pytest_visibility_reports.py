from __future__ import annotations

import json
import os
import subprocess
import sys
from pathlib import Path
from types import SimpleNamespace

import pytest
from tests._pytest_hooks.dependencies import _NOT_COLLECTED_BY_REASON, _REPORT_DIR_ENV
from tests._pytest_hooks.reporting import (
    _MISSING_PRIMARY_MARKER_ITEMS,
    _WORKER_MISSING_MARKER_KEY,
    _WORKER_MISSING_MARKER_STASH_KEY,
    _missing_marker_report_payload,
    _skip_report_payload,
    _slow_report_payload,
    _subprocess_report_payload,
    _write_visibility_reports,
    _xdist_payload,
    pytest_collection_finish,
    pytest_sessionfinish,
    pytest_testnodedown,
)

pytestmark = pytest.mark.unit


@pytest.fixture(autouse=True)
def _preserve_not_collected_report_state() -> None:
    original = {reason: set(paths) for reason, paths in _NOT_COLLECTED_BY_REASON.items()}
    original_missing_markers = list(_MISSING_PRIMARY_MARKER_ITEMS)
    try:
        yield
    finally:
        _NOT_COLLECTED_BY_REASON.clear()
        for reason, paths in original.items():
            _NOT_COLLECTED_BY_REASON[reason].update(paths)
        _MISSING_PRIMARY_MARKER_ITEMS.clear()
        _MISSING_PRIMARY_MARKER_ITEMS.extend(original_missing_markers)


def test_skip_report_payload_extracts_reason_and_location() -> None:
    config = SimpleNamespace(
        option=SimpleNamespace(markexpr="not slow", numprocesses="auto", dist="loadfile"),
        getini=lambda name: "-v --tb=short --strict-markers -ra" if name == "addopts" else "",
    )
    report = SimpleNamespace(
        nodeid="tests/test_gui/test_gui_headless_smoke.py::test_linux_only",
        when="setup",
        location=("tests/test_gui/test_gui_headless_smoke.py", 9, "test_linux_only"),
        longrepr=("tests/test_gui/test_gui_headless_smoke.py", 9, "Skipped: linux only"),
        longreprtext="Skipped: linux only",
    )

    payload = _skip_report_payload([report], config)

    assert payload["summary"]["count"] == 1
    assert payload["summary"]["skipped_count"] == 1
    assert payload["pytest"]["selection"]["markexpr"] == "not slow"
    assert payload["pytest"]["xdist"] == {
        "enabled": True,
        "numprocesses": "auto",
        "dist": "loadfile",
    }
    assert payload["items"] == [
        {
            "nodeid": "tests/test_gui/test_gui_headless_smoke.py::test_linux_only",
            "phase": "setup",
            "reason": "Skipped: linux only",
            "location": {
                "path": "tests/test_gui/test_gui_headless_smoke.py",
                "line": 10,
            },
        }
    ]


def test_write_visibility_reports_emits_expected_json(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv(_REPORT_DIR_ENV, str(tmp_path))
    _NOT_COLLECTED_BY_REASON.clear()
    _NOT_COLLECTED_BY_REASON["missing pandas dependency"].update(
        {
            "tests/test_services/test_spreadsheet_pipeline_smoke.py",
            "tests/test_converter/test_xlsx2md_blocks.py",
        }
    )

    report = SimpleNamespace(
        nodeid="tests/test_gui/test_gui_headless_smoke.py::test_linux_only",
        when="setup",
        location=("tests/test_gui/test_gui_headless_smoke.py", 9, "test_linux_only"),
        longrepr=("tests/test_gui/test_gui_headless_smoke.py", 9, "Skipped: linux only"),
        longreprtext="Skipped: linux only",
    )
    config = SimpleNamespace(
        option=SimpleNamespace(markexpr="not integration", numprocesses=0, dist="no"),
        getini=lambda name: "-v --tb=short --strict-markers -ra" if name == "addopts" else "",
    )

    slow_items = [
        {
            "nodeid": "tests/test_cli/test_cli_e2e.py::test_cli_convert_xlsx_to_md_end_to_end",
            "outcome": "passed",
            "duration_seconds": 1.25,
        }
    ]
    subprocess_records = [
        {
            "argv": ["DocWenCLI.exe", "info", "--json"],
            "command": "DocWenCLI.exe info --json",
            "cwd": str(tmp_path),
            "returncode": 0,
            "timed_out": False,
            "timeout_seconds": 30,
            "duration_seconds": 0.42,
            "text": True,
            "capture_output": True,
        }
    ]
    missing_marker_items = [
        {
            "nodeid": "tests/test_misc/test_sample.py::test_example",
            "path": "tests/test_misc/test_sample.py",
            "primary_markers": [],
        }
    ]

    report_paths = _write_visibility_reports(
        [report],
        config,
        slow_items=slow_items,
        subprocess_records=subprocess_records,
        missing_marker_items=missing_marker_items,
    )
    skip_payload = json.loads(report_paths["skip_report"].read_text(encoding="utf-8"))
    not_collected_payload = json.loads(report_paths["not_collected_report"].read_text(encoding="utf-8"))
    slow_payload = json.loads(report_paths["slow_report"].read_text(encoding="utf-8"))
    subprocess_payload = json.loads(report_paths["subprocess_report"].read_text(encoding="utf-8"))
    missing_marker_payload = json.loads(report_paths["missing_marker_report"].read_text(encoding="utf-8"))

    assert report_paths["skip_report"] == tmp_path / "skip_report.json"
    assert report_paths["not_collected_report"] == tmp_path / "not_collected_report.json"
    assert report_paths["slow_report"] == tmp_path / "slow_report.json"
    assert report_paths["subprocess_report"] == tmp_path / "subprocess_report.json"
    assert report_paths["missing_marker_report"] == tmp_path / "missing_marker_report.json"
    assert skip_payload["summary"]["count"] == 1
    assert skip_payload["items"][0]["reason"] == "Skipped: linux only"
    assert skip_payload["pytest"]["xdist"]["enabled"] is False
    assert not_collected_payload["summary"]["dependency_gated_not_collected_count"] == 2
    assert not_collected_payload["pytest"]["selection"]["markexpr"] == "not integration"
    assert not_collected_payload["summary"]["count"] == 2
    assert not_collected_payload["by_reason"] == [
        {
            "reason": "missing pandas dependency",
            "count": 2,
            "paths": [
                "tests/test_converter/test_xlsx2md_blocks.py",
                "tests/test_services/test_spreadsheet_pipeline_smoke.py",
            ],
        }
    ]
    assert slow_payload["summary"]["count"] == 1
    assert slow_payload["items"][0]["duration_seconds"] == 1.25
    assert subprocess_payload["summary"]["count"] == 1
    assert subprocess_payload["items"][0]["command"] == "DocWenCLI.exe info --json"
    assert missing_marker_payload["summary"]["missing_primary_marker_count"] == 1
    assert missing_marker_payload["items"][0]["path"] == "tests/test_misc/test_sample.py"


def test_additional_visibility_payloads_include_expected_summary(tmp_path: Path) -> None:
    config = SimpleNamespace()
    config.option = SimpleNamespace(markexpr="", numprocesses=0, dist="no")
    config.getini = lambda name: "-v --tb=short --strict-markers -ra" if name == "addopts" else ""

    slow_payload = _slow_report_payload(
        [{"nodeid": "tests/test_a.py::test_x", "outcome": "passed", "duration_seconds": 1.5}],
        config,
    )
    subprocess_payload = _subprocess_report_payload(
        [
            {
                "argv": ["DocWenCLI.exe", "info", "--json"],
                "command": "DocWenCLI.exe info --json",
                "cwd": str(tmp_path),
                "returncode": 0,
                "timed_out": False,
                "timeout_seconds": 30,
                "duration_seconds": 0.3,
                "text": True,
                "capture_output": True,
            }
        ],
        config,
    )
    missing_marker_payload = _missing_marker_report_payload(
        [{"nodeid": "tests/test_a.py::test_x", "path": "tests/test_a.py", "primary_markers": []}],
        config,
    )

    assert slow_payload["summary"]["slow_threshold_seconds"] >= 1.0
    assert subprocess_payload["summary"]["timeout_count"] == 0
    assert missing_marker_payload["required_primary_markers"] == [
        "contract",
        "e2e",
        "gui",
        "gui_smoke",
        "host",
        "integration",
        "packaged",
        "unit",
    ]


@pytest.mark.parametrize("raw_limit", ["not-an-integer", "-1"])
def test_primary_marker_ratchet_rejects_invalid_limit(
    monkeypatch: pytest.MonkeyPatch,
    raw_limit: str,
) -> None:
    monkeypatch.setenv("DOCWEN_PYTEST_MAX_MISSING_PRIMARY_MARKERS", raw_limit)
    session = SimpleNamespace(
        exitstatus=pytest.ExitCode.OK,
        config=SimpleNamespace(pluginmanager=SimpleNamespace(get_plugin=lambda _name: None)),
    )

    pytest_sessionfinish(session, int(pytest.ExitCode.OK))

    assert session.exitstatus == pytest.ExitCode.USAGE_ERROR


def test_primary_marker_ratchet_fails_only_when_green_debt_grows(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("DOCWEN_PYTEST_MAX_MISSING_PRIMARY_MARKERS", "1")
    _MISSING_PRIMARY_MARKER_ITEMS[:] = [
        {"nodeid": "tests/test_a.py::test_a"},
        {"nodeid": "tests/test_b.py::test_b"},
    ]
    config = SimpleNamespace(pluginmanager=SimpleNamespace(get_plugin=lambda _name: None))
    green_session = SimpleNamespace(exitstatus=pytest.ExitCode.OK, config=config)

    pytest_sessionfinish(green_session, int(pytest.ExitCode.OK))

    assert green_session.exitstatus == pytest.ExitCode.TESTS_FAILED
    already_failed_session = SimpleNamespace(exitstatus=pytest.ExitCode.TESTS_FAILED, config=config)
    pytest_sessionfinish(already_failed_session, int(pytest.ExitCode.TESTS_FAILED))
    assert already_failed_session.exitstatus == pytest.ExitCode.TESTS_FAILED


def test_xdist_worker_missing_markers_are_merged_once() -> None:
    first = {"nodeid": "tests/test_a.py::test_a", "path": "tests/test_a.py", "primary_markers": []}
    second = {"nodeid": "tests/test_b.py::test_b", "path": "tests/test_b.py", "primary_markers": []}
    _MISSING_PRIMARY_MARKER_ITEMS[:] = [first, second]
    worker_output: dict[str, object] = {}
    stash = {_WORKER_MISSING_MARKER_STASH_KEY: [first, second]}

    pytest_collection_finish(SimpleNamespace(config=SimpleNamespace(workeroutput=worker_output, stash=stash)))
    _MISSING_PRIMARY_MARKER_ITEMS.clear()
    node = SimpleNamespace(workeroutput=worker_output)
    pytest_testnodedown(node, None)
    pytest_testnodedown(node, None)

    assert worker_output[_WORKER_MISSING_MARKER_KEY] == [first, second]
    assert [first, second] == _MISSING_PRIMARY_MARKER_ITEMS


def test_xdist_payload_reports_the_effective_process_cap() -> None:
    config = SimpleNamespace(option=SimpleNamespace(numprocesses=12, maxprocesses=6, dist="loadfile"))

    assert _xdist_payload(config) == {
        "enabled": True,
        "numprocesses": 6,
        "dist": "loadfile",
    }


@pytest.mark.slow
@pytest.mark.parametrize(("marker_debt_limit", "expected_returncode"), [("1", 0), ("0", 1)])
def test_xdist_visibility_report_preserves_unmarked_worker_items(
    tmp_path: Path,
    marker_debt_limit: str,
    expected_returncode: int,
) -> None:
    project_root = Path(__file__).resolve().parents[2]
    probe = tmp_path / "test_unmarked_xdist_probe.py"
    probe.write_text("def test_unmarked_worker_item():\n    assert True\n", encoding="utf-8")
    report_dir = tmp_path / "reports"
    system_temp = tmp_path / "system-temp"
    system_temp.mkdir()
    environment = os.environ.copy()
    environment.update(
        {
            "DOCWEN_PYTEST_REPORT_DIR": str(report_dir),
            "DOCWEN_PYTEST_MAX_MISSING_PRIMARY_MARKERS": marker_debt_limit,
            "PYTHONDONTWRITEBYTECODE": "1",
            "TEMP": str(system_temp),
            "TMP": str(system_temp),
        }
    )

    completed = subprocess.run(
        [
            sys.executable,
            "-m",
            "pytest",
            str(probe),
            "-n",
            "2",
            "--dist",
            "loadfile",
            f"--rootdir={project_root}",
            f"--confcutdir={tmp_path}",
            "-c",
            str(project_root / "pyproject.toml"),
            "-p",
            "tests._pytest_hooks.reporting",
            "-o",
            "addopts=-q --tb=short --strict-markers --import-mode=importlib -ra",
            "-o",
            f"cache_dir={tmp_path / 'cache'}",
            "--basetemp",
            str(tmp_path / "basetemp"),
        ],
        cwd=project_root,
        env=environment,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=60,
        check=False,
    )

    assert completed.returncode == expected_returncode, completed.stdout + completed.stderr
    if expected_returncode:
        assert "primary marker debt exceeded: actual=1, limit=0" in completed.stdout + completed.stderr
    payload = json.loads((report_dir / "missing_marker_report.json").read_text(encoding="utf-8"))
    assert payload["pytest"]["xdist"] == {"enabled": True, "numprocesses": 2, "dist": "loadfile"}
    assert payload["summary"]["missing_primary_marker_count"] == 1
    assert payload["items"][0]["nodeid"].endswith("test_unmarked_worker_item")

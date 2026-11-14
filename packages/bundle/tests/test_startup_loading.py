"""Fresh-process loading and first-paint evidence for desktop startup."""

from __future__ import annotations

import json
import os
import sys
from pathlib import Path

import pytest
from tests.support.subprocess_runner import run_subprocess

pytestmark = pytest.mark.integration


def test_markdown_discovery_does_not_load_conversion_engines() -> None:
    completed = run_subprocess(
        [
            sys.executable,
            "-c",
            "import sys; from docwen_plugin_markdown import MarkdownPlugin; "
            "p=MarkdownPlugin(); assert p.can_handle('markdown', 'docx'); "
            "assert p.can_handle('markdown', 'xlsx'); "
            "assert not {'docx', 'openpyxl', 'numpy'} & sys.modules.keys(); "
            "assert not any(n.startswith('docwen_plugin_markdown.') and n.endswith('.converter') "
            "for n in sys.modules)",
        ]
    )
    assert completed.returncode == 0, completed.stderr


def test_direct_gui_reports_paint_and_event_loop_without_cli(tmp_path: Path) -> None:
    report = tmp_path / "startup.json"
    environment = dict(os.environ)
    environment.update(
        QT_QPA_PLATFORM="offscreen",
        DOCWEN_DATA_DIR=str(tmp_path / "profile"),
        DOCWEN_CONFIG_DIR=str(tmp_path / "profile" / "configs"),
        DOCWEN_LOG_DIR=str(tmp_path / "profile" / "logs"),
        DOCWEN_LOG_TO_TEMP="",
        DOCWEN_GUI_TEST_STARTUP_REPORT=str(report),
        DOCWEN_GUI_TEST_AUTOCLOSE_MS="500",
    )
    completed = run_subprocess([sys.executable, "-m", "docwen_gui"], env=environment)
    assert completed.returncode == 0, completed.stderr
    observed = json.loads(report.read_text(encoding="utf-8"))
    assert observed["window_visible"] and observed["window_enabled"]
    elapsed = observed["elapsed_ms"]
    assert 0 == elapsed["python_entry"] < elapsed["runtime_ready"] < elapsed["window_created"]
    assert elapsed["window_created"] <= elapsed["first_paint"] <= elapsed["event_loop_ready"]


def test_startup_diagnostics_preserve_existing_evidence(tmp_path: Path) -> None:
    from docwen_bundle.startup_diagnostics import StartupTrace

    report = tmp_path / "startup.json"
    report.write_text("previous evidence", encoding="utf-8")
    trace = StartupTrace(str(report))
    trace.mark("first_paint")
    trace.write(window_visible=True, window_enabled=True)
    assert report.read_text(encoding="utf-8") == "previous evidence"

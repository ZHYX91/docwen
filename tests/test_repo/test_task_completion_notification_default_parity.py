"""Guards VIS-098 task-completion notification default parity evidence."""

from __future__ import annotations

import json
import tomllib
from pathlib import Path

import pytest

pytestmark = pytest.mark.unit

PROJECT_ROOT = Path(__file__).resolve().parents[2]
REPORT_NAME = "task-completion-notification-default-parity-2026-07-16.md"


def _read(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def test_notification_default_config_and_production_gate_are_aligned() -> None:
    config = tomllib.loads(_read(PROJECT_ROOT / "configs" / "gui.toml"))["notifications"]
    main_window = _read(PROJECT_ROOT / "packages" / "apps" / "gui" / "src" / "docwen_gui" / "main_window.py")
    app = _read(PROJECT_ROOT / "packages" / "apps" / "gui" / "src" / "docwen_gui" / "app.py")

    assert config == {
        "task_completion": True,
        "min_elapsed_seconds": 8,
        "system_tray": False,
        "system_tray_timeout_ms": 6000,
    }
    assert 'cfg_port.get("gui.notifications.system_tray", False)' in main_window
    assert "if not tray_enabled or not QSystemTrayIcon.isSystemTrayAvailable():" in main_window
    assert '"defaultTrayIconPresent": False' in app
    assert '"probeCreatedTrayIcon": False' in app
    assert "if probe_created_tray and tray is not None:" in app


@pytest.mark.parametrize(
    ("changes", "error"),
    (
        ({"defaultTrayIconPresent": True}, "packaged_gui_notification_default_tray_icon_present"),
        ({"probeCreatedTrayIcon": False}, "packaged_gui_notification_probe_tray_not_created"),
    ),
)
def test_packaged_notification_verifier_rejects_default_or_probe_drift(
    tmp_path: Path,
    changes: dict[str, bool],
    error: str,
) -> None:
    from scripts.release import verify_packaged_gui

    report = {
        "isSystemTrayAvailable": True,
        "supportsMessages": True,
        "defaultTrayIconPresent": False,
        "probeCreatedTrayIcon": True,
        "hasTrayIcon": True,
        "showMessageCalled": True,
        "error": None,
    }
    report.update(changes)
    report_path = tmp_path / "notification_smoke.json"
    report_path.write_text(json.dumps(report), encoding="utf-8")

    with pytest.raises(RuntimeError, match=error):
        verify_packaged_gui._verify_notification_smoke_report(report_path)

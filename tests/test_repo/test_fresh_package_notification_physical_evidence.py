"""Fail-closed contracts for VIS-2026-07-17-114 notification evidence."""

from __future__ import annotations

from pathlib import Path

import pytest

pytestmark = pytest.mark.contract

ROOT = Path(__file__).resolve().parents[2]
REPORT_NAME = "fresh-package-held-notification-boundary-2026-07-17.md"


def test_held_notification_probe_is_explicit_and_default_safe() -> None:
    app_source = (ROOT / "packages/apps/gui/src/docwen_gui/app.py").read_text(encoding="utf-8")
    regression = (ROOT / "packages/apps/gui/tests/test_app_notification_probe.py").read_text(encoding="utf-8")

    assert "DOCWEN_GUI_TEST_NOTIFICATION_HOLD_MS" in app_source
    assert "message_timeout_ms = max(500, probe_hold_ms)" in app_source
    assert "QTimer.singleShot(probe_hold_ms, _cleanup_probe_tray)" in app_source
    assert 'os.environ.get("DOCWEN_GUI_TEST_NOTIFICATION_HOLD_MS", "0")' in app_source
    assert "test_notification_probe_can_hold_temporary_tray_for_physical_observation" in regression
    assert '"probeHoldMs": 6000' in regression
    assert '"messageTimeoutMs": 6000' in regression

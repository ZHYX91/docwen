from __future__ import annotations

import json
from pathlib import Path
from types import SimpleNamespace
from typing import Any, cast

import pytest

pytestmark = pytest.mark.gui


def test_notification_probe_can_hold_temporary_tray_for_physical_observation(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from PySide6 import QtWidgets
    from PySide6.QtCore import QTimer

    from docwen_gui.app import _schedule_test_notification_report

    report_path = tmp_path / "notification_smoke.json"
    monkeypatch.setenv("DOCWEN_GUI_TEST_NOTIFICATION_REPORT", str(report_path))
    monkeypatch.setenv("DOCWEN_GUI_TEST_NOTIFICATION_HOLD_MS", "6000")

    scheduled: list[tuple[int, object]] = []
    events: list[object] = []

    class _FakeTray:
        class MessageIcon:
            Information = "information"

        @staticmethod
        def isSystemTrayAvailable() -> bool:
            return True

        @staticmethod
        def supportsMessages() -> bool:
            return True

        def __init__(self, icon: object, parent: object) -> None:
            events.append(("created", icon, parent))

        def setVisible(self, visible: bool) -> None:
            events.append(("visible", visible))

        def showMessage(self, title: str, body: str, icon: object, timeout_ms: int) -> None:
            events.append(("message", title, body, icon, timeout_ms))

        def hide(self) -> None:
            events.append("hidden")

        def deleteLater(self) -> None:
            events.append("deleted")

    monkeypatch.setattr(QtWidgets, "QSystemTrayIcon", _FakeTray)
    monkeypatch.setattr(
        QTimer,
        "singleShot",
        lambda delay, callback: scheduled.append((delay, callback)),
    )

    app = SimpleNamespace(processEvents=lambda: events.append("processed"))
    window = SimpleNamespace(
        _system_tray_icon=None,
        windowIcon=lambda: "window-icon",
    )

    _schedule_test_notification_report(cast(Any, app), cast(Any, window))

    assert scheduled[0][0] == 250
    cast(Any, scheduled.pop(0)[1])()

    payload = json.loads(report_path.read_text(encoding="utf-8"))
    assert payload == {
        "isSystemTrayAvailable": True,
        "supportsMessages": True,
        "defaultTrayIconPresent": False,
        "probeCreatedTrayIcon": True,
        "hasTrayIcon": True,
        "showMessageCalled": True,
        "probeHoldMs": 6000,
        "messageTimeoutMs": 6000,
        "error": None,
    }
    assert events[-1] == "processed"
    assert "hidden" not in events
    assert "deleted" not in events
    assert scheduled[0][0] == 6000

    cast(Any, scheduled.pop(0)[1])()
    assert events[-2:] == ["hidden", "deleted"]

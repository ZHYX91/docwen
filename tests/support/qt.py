from __future__ import annotations

from collections.abc import Callable

import pytest
from PySide6.QtWidgets import QWidget
from pytestqt.qtbot import QtBot


def register_widget[WidgetT: QWidget](qtbot: QtBot, widget: WidgetT) -> WidgetT:
    """Register a widget with qtbot and return it for inline construction."""

    qtbot.addWidget(widget)
    return widget


def wait_for_ui(qtbot: QtBot, ms: int = 0) -> None:
    """Pump the Qt event loop through qtbot instead of direct processEvents()."""

    qtbot.wait(ms)


def wait_until(qtbot: QtBot, predicate: Callable[[], bool], *, timeout: int = 1000) -> None:
    """Wait for a UI condition through pytest-qt's stable polling helper."""

    qtbot.waitUntil(predicate, timeout=timeout)


@pytest.fixture
def qt_widget(qtbot: QtBot) -> Callable[[QWidget], QWidget]:
    """Small factory for tests that want qtbot-managed widget cleanup."""

    def _register[WidgetT: QWidget](widget: WidgetT) -> WidgetT:
        return register_widget(qtbot, widget)

    return _register

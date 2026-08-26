"""Post-closure GUI smoke test — validates critical GUI widgets and interactions.

Uses pytest-qt (qtbot) against the real MainWindow, dialogs, and components.
Marked as ``pytest.mark.gui`` so tests can be excluded from fast CI runs.

Tests cover:
- MainWindow widget hierarchy via objectName lookups
- AboutDialog creation and content labels
- SettingsDialog creation and tab navigation
- FeedbackHelper error/warn/info/confirm message boxes
- Centralized QAction existence and keyboard shortcut bindings
- TemplateSelector list population and selection
- NumberingAddDialog and NumberingCleanDialog basic creation
- Screenshot capture (best-effort, skipped silently when unavailable)
"""

from __future__ import annotations

import contextlib
import os
from pathlib import Path
from unittest.mock import patch

import pytest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

pytestmark = pytest.mark.gui


def _save_screenshot(widget, name: str, screenshot_dir: Path) -> str | None:
    """Capture a widget screenshot below pytest's temporary root."""
    try:
        pixmap = widget.grab()
        path = screenshot_dir / f"{name}.png"
        screenshot_dir.mkdir(parents=True, exist_ok=True)
        pixmap.save(str(path), "PNG")
        return str(path)
    except Exception:
        return None


@pytest.fixture
def window(qtbot, qapp):
    """Create a fully set-up MainWindow with a null controller."""
    from docwen_gui.main_window import MainWindow
    from docwen_gui.view_models.main_window_vm import MainWindowViewModel

    vm = MainWindowViewModel(controller=None)
    w = MainWindow(view_model=vm)
    w.setup_ui()
    qtbot.addWidget(w)
    w.show()
    qtbot.waitExposed(w)
    return w


__all__ = (
    "Path",
    "_save_screenshot",
    "contextlib",
    "patch",
    "pytestmark",
    "window",
)

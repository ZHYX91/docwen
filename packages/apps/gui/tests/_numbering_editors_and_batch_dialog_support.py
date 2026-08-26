"""Tests for numbering editors and batch dialog.

Covers:
- NumberingAddDialog: creation, scheme CRUD, level editing, preview,
  dirty tracking, save/reject, validation.
- NumberingCleanDialog: creation, rule CRUD, regex editing, live test,
  dirty tracking, save/reject.
- TomlEditorWidget: construction, load, save validation, config switching.
- show_batch_add_failed_dialog: coverage of the dialog function (real QMessageBox).

Requires QApplication (provided by conftest.py ``qapp`` fixture).
"""

from __future__ import annotations

import os
from pathlib import Path

import pytest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication, QMessageBox, QWidget

pytestmark = pytest.mark.gui


def _write_minimal_base_config_tree(base_dir: Path) -> None:
    """Create an empty TOML file for every spec in the registry under base_dir."""
    from docwen_runtime.config.registry import CONFIG_FILES

    for spec in CONFIG_FILES:
        path = base_dir / spec.rel_path
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text("\n", encoding="utf-8")


def _patch_dialog_confirm(dlg: QWidget, return_value: bool = True) -> None:
    """Monkey-patch the ``_confirm`` method on a dialog to return fixed value."""
    dlg._confirm = lambda title="", message="": return_value  # type: ignore[method-assign]


def _patch_dialog_notify(dlg: QWidget) -> None:
    """Monkey-patch ``_notify_info`` and ``_notify_warning`` on a dialog."""
    dlg._notify_info = lambda message, title="": None  # type: ignore[method-assign]
    dlg._notify_warning = lambda message: None  # type: ignore[method-assign]
    dlg._notify_error = lambda title, message: None  # type: ignore[method-assign]


def _patch_all_modals(dlg: QWidget, confirm_value: bool = True) -> None:
    """Suppress all modal dialogs on a dialog instance."""
    _patch_dialog_confirm(dlg, confirm_value)
    _patch_dialog_notify(dlg)


__all__ = (
    "Path",
    "QApplication",
    "QMessageBox",
    "_patch_all_modals",
    "_patch_dialog_notify",
    "_write_minimal_base_config_tree",
    "pytest",
    "pytestmark",
)

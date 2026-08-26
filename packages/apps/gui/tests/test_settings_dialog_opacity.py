"""SettingsDialog opacity preview behavior.

These tests cover the user-visible Settings UX parity contract from the
PySide6 reference project: changing transparency in General settings previews
the main-window opacity immediately, and Cancel rolls it back.
"""

from __future__ import annotations

import pytest
from PySide6.QtWidgets import QWidget

from docwen_gui.view_models.settings_vm import SettingsViewModel
from docwen_gui.widgets.settings.dialog import SettingsDialog
from docwen_gui.widgets.settings.general_tab import GeneralTab

pytestmark = pytest.mark.gui


def _general_tab(dialog: SettingsDialog) -> tuple[SettingsViewModel, GeneralTab]:
    general = dialog._tabs["general"]  # pyright: ignore[reportPrivateUsage]
    assert isinstance(general, GeneralTab)
    return dialog.view_model, general


def test_transparency_value_previews_parent_opacity(qapp) -> None:
    parent = QWidget()
    parent.setWindowOpacity(1.0)
    dialog = SettingsDialog(parent=parent, view_model=SettingsViewModel())
    try:
        _vm, general = _general_tab(dialog)
        general._transparency_enabled.setChecked(True)  # pyright: ignore[reportPrivateUsage]
        general._transparency_value.setValue(0.55)  # pyright: ignore[reportPrivateUsage]
        qapp.processEvents()

        assert abs(parent.windowOpacity() - 0.55) < 0.01
    finally:
        dialog.close()
        parent.close()


def test_transparency_toggle_off_previews_opaque_parent(qapp) -> None:
    parent = QWidget()
    parent.setWindowOpacity(1.0)
    dialog = SettingsDialog(parent=parent, view_model=SettingsViewModel())
    try:
        _vm, general = _general_tab(dialog)
        general._transparency_enabled.setChecked(True)  # pyright: ignore[reportPrivateUsage]
        general._transparency_value.setValue(0.55)  # pyright: ignore[reportPrivateUsage]
        general._transparency_enabled.setChecked(False)  # pyright: ignore[reportPrivateUsage]
        qapp.processEvents()

        assert abs(parent.windowOpacity() - 1.0) < 0.01
    finally:
        dialog.close()
        parent.close()


def test_cancel_restores_parent_opacity(monkeypatch: pytest.MonkeyPatch, qapp) -> None:
    from docwen_gui.widgets.settings import dialog as dialog_module

    parent = QWidget()
    parent.setWindowOpacity(0.8)
    dialog = SettingsDialog(parent=parent, view_model=SettingsViewModel())
    monkeypatch.setattr(dialog_module, "_show_confirm", lambda *args, **kwargs: True)
    cancelled = False
    try:
        _vm, general = _general_tab(dialog)
        general._transparency_enabled.setChecked(True)  # pyright: ignore[reportPrivateUsage]
        general._transparency_value.setValue(0.5)  # pyright: ignore[reportPrivateUsage]
        qapp.processEvents()
        assert abs(parent.windowOpacity() - 0.5) < 0.01

        dialog._on_cancel()  # pyright: ignore[reportPrivateUsage]
        cancelled = True
        qapp.processEvents()

        assert abs(parent.windowOpacity() - 0.8) < 0.01
    finally:
        if not cancelled:
            dialog.close()
        parent.close()


def test_transparency_preview_without_parent_is_safe_noop_and_updates_draft(qapp) -> None:
    dialog = SettingsDialog(view_model=SettingsViewModel())
    try:
        vm, general = _general_tab(dialog)
        general._transparency_enabled.setChecked(True)  # pyright: ignore[reportPrivateUsage]
        general._transparency_value.setValue(0.66)  # pyright: ignore[reportPrivateUsage]
        qapp.processEvents()

        assert vm.config.gui.transparency_enabled is True
        assert abs(vm.config.gui.transparency_value - 0.66) < 0.01
    finally:
        dialog.close()

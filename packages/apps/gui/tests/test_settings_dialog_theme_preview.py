"""SettingsDialog theme preview behavior."""

from __future__ import annotations

import pytest
from PySide6.QtCore import QCoreApplication, QEvent

from docwen_gui.view_models.settings_vm import SettingsViewModel
from docwen_gui.widgets.settings.dialog import SettingsDialog
from docwen_gui.widgets.settings.general_tab import GeneralTab

pytestmark = pytest.mark.gui


def _flush_deferred_deletes(qapp) -> None:
    QCoreApplication.sendPostedEvents(None, QEvent.Type.DeferredDelete)
    qapp.processEvents()


def test_theme_change_previews_and_cancel_restores_initial_theme(monkeypatch: pytest.MonkeyPatch, qapp) -> None:
    from docwen_gui.styles.theme_manager import ThemeManager
    from docwen_gui.widgets.settings import dialog as dialog_module

    ThemeManager.reset_instance()
    manager = ThemeManager.get_instance()
    manager.initialize(qapp, "light")

    dialog = SettingsDialog(view_model=SettingsViewModel())
    monkeypatch.setattr(dialog_module, "_show_confirm", lambda *args, **kwargs: True)
    cancelled = False
    try:
        general = dialog._tabs["general"]  # pyright: ignore[reportPrivateUsage]
        assert isinstance(general, GeneralTab)
        combo = general._theme_combo  # pyright: ignore[reportPrivateUsage]
        dark_index = combo.findData("dark")
        assert dark_index >= 0

        combo.setCurrentIndex(dark_index)
        qapp.processEvents()

        assert manager.get_current_theme() == "dark"

        dialog._on_cancel()  # pyright: ignore[reportPrivateUsage]
        cancelled = True
        qapp.processEvents()

        assert manager.get_current_theme() == "light"
    finally:
        if not cancelled:
            dialog.close()
        _flush_deferred_deletes(qapp)
        manager.apply_theme("light")


def test_apply_updates_theme_cancel_rollback_baseline(monkeypatch: pytest.MonkeyPatch, qapp) -> None:
    from docwen_gui.styles.theme_manager import ThemeManager
    from docwen_gui.widgets.settings import dialog as dialog_module

    ThemeManager.reset_instance()
    manager = ThemeManager.get_instance()
    manager.initialize(qapp, "light")

    vm = SettingsViewModel()
    dialog = SettingsDialog(view_model=vm)
    monkeypatch.setattr(dialog_module, "_show_confirm", lambda *args, **kwargs: True)
    cancelled = False
    try:
        general = dialog._tabs["general"]  # pyright: ignore[reportPrivateUsage]
        assert isinstance(general, GeneralTab)
        combo = general._theme_combo  # pyright: ignore[reportPrivateUsage]
        dark_index = combo.findData("dark")
        light_index = combo.findData("light")
        assert dark_index >= 0
        assert light_index >= 0

        combo.setCurrentIndex(dark_index)
        qapp.processEvents()
        assert manager.get_current_theme() == "dark"

        dialog._on_apply()  # pyright: ignore[reportPrivateUsage]

        combo.setCurrentIndex(light_index)
        qapp.processEvents()
        assert manager.get_current_theme() == "light"

        dialog._on_cancel()  # pyright: ignore[reportPrivateUsage]
        cancelled = True
        qapp.processEvents()

        assert manager.get_current_theme() == "dark"
    finally:
        if not cancelled:
            dialog.close()
        _flush_deferred_deletes(qapp)
        manager.apply_theme("light")


def test_settings_dialog_is_deleted_on_close(qapp) -> None:
    from shiboken6 import isValid

    dialog = SettingsDialog(view_model=SettingsViewModel())
    assert dialog in qapp.topLevelWidgets()

    dialog.close()
    _flush_deferred_deletes(qapp)

    assert not isValid(dialog)

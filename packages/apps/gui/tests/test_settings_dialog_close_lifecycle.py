"""Real-Qt coverage for the SettingsDialog cancel/close lifecycle."""

from __future__ import annotations

from collections.abc import Callable

import pytest
from PySide6.QtCore import Qt
from PySide6.QtTest import QSignalSpy, QTest
from PySide6.QtWidgets import QPushButton, QWidget

from docwen_gui.models.settings_config import GUIConfig, SettingsConfig
from docwen_gui.view_models.settings_vm import SettingsViewModel
from docwen_gui.widgets.settings import dialog as dialog_module
from docwen_gui.widgets.settings.dialog import SettingsDialog
from docwen_gui.widgets.settings.general_tab import GeneralTab

pytestmark = pytest.mark.gui


def _dirty_opacity_preview(qapp) -> tuple[QWidget, SettingsViewModel, SettingsDialog]:
    parent = QWidget()
    parent.setWindowOpacity(0.8)
    vm = SettingsViewModel(
        config=SettingsConfig(
            gui=GUIConfig(
                transparency_enabled=True,
                transparency_value=0.8,
            )
        )
    )
    dialog = SettingsDialog(parent=parent, view_model=vm)
    dialog.show()
    qapp.processEvents()

    general = dialog._tabs["general"]  # pyright: ignore[reportPrivateUsage]
    assert isinstance(general, GeneralTab)
    general._transparency_value.setValue(0.4)  # pyright: ignore[reportPrivateUsage]
    qapp.processEvents()

    assert vm.is_dirty is True
    assert abs(parent.windowOpacity() - 0.4) < 0.01
    return parent, vm, dialog


def _track_cancel_changes(vm: SettingsViewModel) -> list[bool]:
    calls: list[bool] = []
    original = vm.cancel_changes

    def cancel_changes() -> None:
        calls.append(True)
        original()

    vm.cancel_changes = cancel_changes  # type: ignore[method-assign]
    return calls


def test_window_manager_close_event_for_title_bar_or_alt_f4_refuses_then_rolls_back(
    qapp,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    parent, vm, dialog = _dirty_opacity_preview(qapp)
    cancel_calls = _track_cancel_changes(vm)
    answers = iter((False, True))
    confirmations: list[bool] = []

    def confirm(*_args, **_kwargs) -> bool:
        confirmations.append(True)
        return next(answers)

    monkeypatch.setattr(dialog_module, "_show_confirm", confirm)
    rejected = QSignalSpy(dialog.rejected)
    try:
        assert dialog.close() is False
        assert dialog.isVisible() is True
        assert vm.is_dirty is True
        assert abs(parent.windowOpacity() - 0.4) < 0.01
        assert cancel_calls == []
        assert rejected.count() == 0

        assert dialog.close() is True
        assert dialog.isVisible() is False
        assert vm.is_dirty is False
        assert abs(vm.config.gui.transparency_value - 0.8) < 0.01
        assert abs(parent.windowOpacity() - 0.8) < 0.01
        assert cancel_calls == [True]
        assert rejected.count() == 1
        assert confirmations == [True, True]
    finally:
        parent.close()


@pytest.mark.parametrize(
    "close_action",
    ("programmatic-reject", "cancel-button", "escape"),
)
def test_every_cancel_entry_point_confirms_and_rolls_back_once(
    qapp,
    monkeypatch: pytest.MonkeyPatch,
    close_action: str,
) -> None:
    parent, vm, dialog = _dirty_opacity_preview(qapp)
    cancel_calls = _track_cancel_changes(vm)
    confirmations: list[bool] = []
    monkeypatch.setattr(
        dialog_module,
        "_show_confirm",
        lambda *_args, **_kwargs: confirmations.append(True) or True,
    )
    rejected = QSignalSpy(dialog.rejected)

    if close_action == "programmatic-reject":
        action: Callable[[], object] = dialog.reject
    elif close_action == "cancel-button":
        button = dialog.findChild(QPushButton, "settingsCancelButton")
        assert button is not None

        def action() -> object:
            return QTest.mouseClick(button, Qt.MouseButton.LeftButton)

    else:

        def action() -> object:
            return QTest.keyClick(dialog, Qt.Key.Key_Escape)

    try:
        action()

        assert dialog.isVisible() is False
        assert vm.is_dirty is False
        assert abs(vm.config.gui.transparency_value - 0.8) < 0.01
        assert abs(parent.windowOpacity() - 0.8) < 0.01
        assert confirmations == [True]
        assert cancel_calls == [True]
        assert rejected.count() == 1
    finally:
        if dialog.isVisible():
            vm.cancel_changes()
            dialog.close()
        parent.close()


def test_reentrant_reject_during_preview_restore_does_not_repeat_cleanup(
    qapp,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    parent, vm, dialog = _dirty_opacity_preview(qapp)
    cancel_calls = _track_cancel_changes(vm)
    restore_calls: list[bool] = []
    original_restore = dialog._restore_preview_state  # pyright: ignore[reportPrivateUsage]

    def restore_with_reentrant_reject() -> None:
        restore_calls.append(True)
        dialog.reject()
        original_restore()

    monkeypatch.setattr(dialog_module, "_show_confirm", lambda *_args, **_kwargs: True)
    monkeypatch.setattr(dialog, "_restore_preview_state", restore_with_reentrant_reject)
    try:
        dialog.reject()

        assert dialog.isVisible() is False
        assert restore_calls == [True]
        assert cancel_calls == [True]
        assert vm.is_dirty is False
        assert abs(parent.windowOpacity() - 0.8) < 0.01
    finally:
        if dialog.isVisible():
            vm.cancel_changes()
            dialog.close()
        parent.close()

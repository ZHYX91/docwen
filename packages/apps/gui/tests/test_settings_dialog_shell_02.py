"""Focused tests split from test_settings_dialog_shell.py."""

from __future__ import annotations

from ._settings_dialog_shell_support import (
    pytest,
)

pytestmark = pytest.mark.gui


def test_settings_dialog_apply_refreshes_source_after_success_or_partial_failure(qapp, monkeypatch) -> None:
    from PySide6.QtTest import QSignalSpy

    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings.dialog import SettingsDialog

    vm = SettingsViewModel()
    dialog = SettingsDialog(view_model=vm)
    spy = QSignalSpy(dialog.settings_source_changed)

    monkeypatch.setattr(vm, "apply_changes", lambda: False)
    dialog._on_apply()  # pyright: ignore[reportPrivateUsage]
    assert spy.count() == 1

    monkeypatch.setattr(vm, "apply_changes", lambda: True)
    dialog._on_apply()  # pyright: ignore[reportPrivateUsage]
    assert spy.count() == 2

    def raise_after_partial_apply() -> bool:
        raise OSError("simulated partial apply")

    monkeypatch.setattr(vm, "apply_changes", raise_after_partial_apply)
    dialog._on_apply()  # pyright: ignore[reportPrivateUsage]
    assert spy.count() == 3

    dialog.close()


@pytest.mark.parametrize("apply_action", ("apply", "ok"))
def test_settings_dialog_partial_apply_uses_persisted_visual_cancel_baseline(
    qapp,
    monkeypatch: pytest.MonkeyPatch,
    apply_action: str,
) -> None:
    from copy import deepcopy

    from PySide6.QtWidgets import QWidget

    from docwen_application.controller import ApplicationController
    from docwen_gui.styles.theme_manager import ThemeManager
    from docwen_gui.view_models.settings_vm import SECTION_GUI, SECTION_OUTPUT, SettingsViewModel
    from docwen_gui.widgets.settings import dialog as dialog_module
    from docwen_gui.widgets.settings.dialog import SettingsDialog

    class PartialPort:
        def __init__(self) -> None:
            self.raw: dict[str, object] = {
                "gui": {
                    "theme": {"default_theme": "dark"},
                    "transparency": {"enabled": True, "default_value": 0.55},
                },
                "output": {"directory": {"mode": "source"}},
            }

        def snapshot(self) -> dict[str, object]:
            return deepcopy(self.raw)

        def get(self, key: str, default: object = None) -> object:
            current: object = self.raw
            for part in key.split("."):
                if not isinstance(current, dict) or part not in current:
                    return default
                current = current[part]
            return current

        def reload(self) -> None:
            return None

    ThemeManager.reset_instance()
    manager = ThemeManager.get_instance()
    manager.initialize(qapp, "dark")
    parent = QWidget()
    parent.setWindowOpacity(0.55)
    port = PartialPort()
    controller = ApplicationController(config_port=port)  # type: ignore[arg-type]
    vm = SettingsViewModel(controller=controller)
    dialog = SettingsDialog(parent=parent, view_model=vm)
    monkeypatch.setattr(dialog_module, "_show_confirm", lambda *_args, **_kwargs: True)

    def persist_gui_then_fail(config, _baseline=None) -> bool:
        port.raw["gui"] = {
            "theme": {"default_theme": config.gui.theme},
            "transparency": {
                "enabled": config.gui.transparency_enabled,
                "default_value": config.gui.transparency_value,
            },
        }
        return False

    monkeypatch.setattr(vm, "_persist_to_controller_config", persist_gui_then_fail)
    rejected = False
    try:
        vm.set_field(SECTION_GUI, "theme", "light")
        vm.set_field(SECTION_GUI, "transparency_enabled", False)
        vm.set_field(SECTION_GUI, "transparency_value", 1.0)
        vm.set_field(SECTION_OUTPUT, "output_mode", "custom")
        manager.apply_theme("light")
        parent.setWindowOpacity(1.0)

        if apply_action == "apply":
            dialog._on_apply()  # pyright: ignore[reportPrivateUsage]
        else:
            dialog._on_ok()  # pyright: ignore[reportPrivateUsage]

        assert vm.is_dirty is True
        assert vm.persisted_config.gui.theme == "light"
        assert vm.persisted_config.output.output_mode == "source"
        manager.apply_theme("dark")
        parent.setWindowOpacity(0.55)

        dialog._on_cancel()  # pyright: ignore[reportPrivateUsage]
        rejected = True

        assert manager.get_current_theme() == "light"
        assert abs(parent.windowOpacity() - 1.0) < 0.01
        assert vm.config.output.output_mode == "source"
    finally:
        if not rejected:
            dialog.close()
        parent.close()
        manager.apply_theme("light")


@pytest.mark.parametrize("reset_action", ("tab", "all"))
def test_settings_dialog_failed_noop_reset_keeps_persisted_visual_cancel_baseline(
    qapp,
    monkeypatch: pytest.MonkeyPatch,
    reset_action: str,
) -> None:
    from PySide6.QtWidgets import QWidget

    from docwen_gui.models.settings_config import GUIConfig, SettingsConfig
    from docwen_gui.styles.theme_manager import ThemeManager
    from docwen_gui.view_models.settings_vm import SECTION_GUI, SettingsViewModel
    from docwen_gui.widgets.settings import dialog as dialog_module
    from docwen_gui.widgets.settings.dialog import TAB_KEYS, SettingsDialog

    ThemeManager.reset_instance()
    manager = ThemeManager.get_instance()
    manager.initialize(qapp, "dark")
    parent = QWidget()
    parent.setWindowOpacity(0.55)
    vm = SettingsViewModel(
        config=SettingsConfig(gui=GUIConfig(theme="dark", transparency_enabled=True, transparency_value=0.55))
    )
    dialog = SettingsDialog(parent=parent, view_model=vm)
    monkeypatch.setattr(dialog_module, "_show_confirm", lambda *_args, **_kwargs: True)
    monkeypatch.setattr(vm, "reset_group", lambda _group: False)
    monkeypatch.setattr(vm, "reset_all", lambda: False)
    rejected = False
    try:
        vm.set_field(SECTION_GUI, "theme", "light")
        vm.set_field(SECTION_GUI, "transparency_enabled", False)
        vm.set_field(SECTION_GUI, "transparency_value", 1.0)
        manager.apply_theme("light")
        parent.setWindowOpacity(1.0)

        if reset_action == "tab":
            dialog._tab_widget.setCurrentIndex(TAB_KEYS.index("general"))  # pyright: ignore[reportPrivateUsage]
            dialog._on_reset_tab()  # pyright: ignore[reportPrivateUsage]
        else:
            dialog._on_reset_all()  # pyright: ignore[reportPrivateUsage]

        assert manager.get_current_theme() == "light"
        assert abs(parent.windowOpacity() - 1.0) < 0.01
        dialog._on_cancel()  # pyright: ignore[reportPrivateUsage]
        rejected = True

        assert manager.get_current_theme() == "dark"
        assert abs(parent.windowOpacity() - 0.55) < 0.01
    finally:
        if not rejected:
            dialog.close()
        parent.close()
        manager.apply_theme("light")


def test_settings_dialog_reset_reload_failures_are_contained_and_other_tabs_continue(
    qapp,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from PySide6.QtTest import QSignalSpy

    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings import dialog as dialog_module
    from docwen_gui.widgets.settings.dialog import TAB_KEYS, SettingsDialog

    vm = SettingsViewModel()
    dialog = SettingsDialog(view_model=vm)
    spy = QSignalSpy(dialog.settings_source_changed)
    reloads: list[str] = []
    visual_syncs: list[bool] = []

    def fail_general_reload() -> None:
        reloads.append("general")
        raise RuntimeError("simulated tab reload failure")

    def reload_text() -> None:
        reloads.append("text")

    monkeypatch.setattr(dialog_module, "_show_confirm", lambda *_args, **_kwargs: True)
    monkeypatch.setattr(vm, "reset_group", lambda _group: True)
    monkeypatch.setattr(vm, "reset_all", lambda: True)
    monkeypatch.setattr(
        dialog,
        "_apply_visual_config_as_preview_baseline",
        lambda: visual_syncs.append(True),
    )
    monkeypatch.setattr(dialog._tabs["general"], "reload_from_config", fail_general_reload)  # pyright: ignore[reportPrivateUsage]
    monkeypatch.setattr(dialog._tabs["text"], "reload_from_config", reload_text)  # pyright: ignore[reportPrivateUsage]
    dialog._tab_widget.setCurrentIndex(TAB_KEYS.index("general"))  # pyright: ignore[reportPrivateUsage]

    dialog._on_reset_tab()  # pyright: ignore[reportPrivateUsage]
    assert reloads == ["general"]
    assert spy.count() == 1
    assert len(visual_syncs) == 1

    reloads.clear()
    dialog._on_reset_all()  # pyright: ignore[reportPrivateUsage]
    assert reloads[:2] == ["general", "text"]
    assert spy.count() == 2
    assert len(visual_syncs) == 2
    dialog.close()


def test_settings_dialog_ok_refreshes_every_attempt_and_closes_only_after_success(qapp, monkeypatch) -> None:
    from PySide6.QtTest import QSignalSpy

    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings.dialog import SettingsDialog

    vm = SettingsViewModel()
    dialog = SettingsDialog(view_model=vm)
    spy = QSignalSpy(dialog.settings_source_changed)
    accept_calls: list[bool] = []
    monkeypatch.setattr(dialog, "accept", lambda: accept_calls.append(True))

    monkeypatch.setattr(vm, "ok_changes", lambda: False)
    dialog._on_ok()  # pyright: ignore[reportPrivateUsage]
    assert spy.count() == 1
    assert accept_calls == []

    def raise_after_partial_ok() -> bool:
        raise OSError("simulated partial OK")

    monkeypatch.setattr(vm, "ok_changes", raise_after_partial_ok)
    dialog._on_ok()  # pyright: ignore[reportPrivateUsage]
    assert spy.count() == 2
    assert accept_calls == []

    monkeypatch.setattr(vm, "ok_changes", lambda: True)
    dialog._on_ok()  # pyright: ignore[reportPrivateUsage]
    assert spy.count() == 3
    assert accept_calls == [True]
    dialog.close()


def test_settings_dialog_reset_attempts_refresh_possible_partial_source(
    qapp,
    monkeypatch,
) -> None:
    from PySide6.QtTest import QSignalSpy

    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings import dialog as dialog_module
    from docwen_gui.widgets.settings.dialog import TAB_KEYS, SettingsDialog

    vm = SettingsViewModel()
    dialog = SettingsDialog(view_model=vm)
    spy = QSignalSpy(dialog.settings_source_changed)
    monkeypatch.setattr(dialog_module, "_show_confirm", lambda *_args, **_kwargs: True)
    monkeypatch.setattr(vm, "reset_group", lambda _group: False)
    dialog._tab_widget.setCurrentIndex(TAB_KEYS.index("general"))  # pyright: ignore[reportPrivateUsage]

    dialog._on_reset_tab()  # pyright: ignore[reportPrivateUsage]

    assert spy.count() == 1

    monkeypatch.setattr(vm, "reset_group", lambda _group: True)
    dialog._on_reset_tab()  # pyright: ignore[reportPrivateUsage]

    assert spy.count() == 2

    monkeypatch.setattr(vm, "reset_all", lambda: False)
    dialog._on_reset_all()  # pyright: ignore[reportPrivateUsage]

    assert spy.count() == 3

    monkeypatch.setattr(vm, "reset_all", lambda: True)
    dialog._on_reset_all()  # pyright: ignore[reportPrivateUsage]

    assert spy.count() == 4

    def raise_after_partial_reset(*_args) -> bool:
        raise OSError("simulated partial reset")

    monkeypatch.setattr(vm, "reset_group", raise_after_partial_reset)
    dialog._on_reset_tab()  # pyright: ignore[reportPrivateUsage]
    assert spy.count() == 5

    monkeypatch.setattr(vm, "reset_all", raise_after_partial_reset)
    dialog._on_reset_all()  # pyright: ignore[reportPrivateUsage]
    assert spy.count() == 6
    dialog.close()


def test_settings_dialog_cancel_does_not_emit_source_change(qapp, monkeypatch) -> None:
    from PySide6.QtTest import QSignalSpy

    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings import dialog as dialog_module
    from docwen_gui.widgets.settings.dialog import SettingsDialog

    vm = SettingsViewModel()
    dialog = SettingsDialog(view_model=vm)
    spy = QSignalSpy(dialog.settings_source_changed)
    monkeypatch.setattr(vm, "cancel_changes", lambda: None)
    monkeypatch.setattr(dialog_module, "_show_confirm", lambda *_args, **_kwargs: True)

    dialog._on_cancel()  # pyright: ignore[reportPrivateUsage]

    assert spy.count() == 0


@pytest.mark.parametrize("finish_action", ("cancel", "ok"))
def test_settings_dialog_reset_tab_preserves_other_tab_draft_and_widget(
    qapp,
    monkeypatch,
    finish_action: str,
) -> None:
    from copy import deepcopy

    from docwen_application.controller import ApplicationController
    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings import dialog as dialog_module
    from docwen_gui.widgets.settings.dialog import TAB_KEYS, SettingsDialog

    class SuccessfulGeneralResetPort:
        def __init__(self) -> None:
            self.raw: dict[str, object] = {
                "gui": {
                    "window": {
                        "remember_gui_state": False,
                        "auto_center": True,
                        "expand_side_panels": True,
                    }
                },
                "output": {"directory": {"mode": "source"}},
            }

        def snapshot(self) -> dict[str, object]:
            return deepcopy(self.raw)

        def get(self, _key: str, default: object = None) -> object:
            return default

        def set_many(self, values: dict[str, object]) -> bool:
            for dotted_key, value in values.items():
                current = self.raw
                parts = dotted_key.split(".")
                for part in parts[:-1]:
                    child = current.setdefault(part, {})
                    assert isinstance(child, dict)
                    current = child
                current[parts[-1]] = deepcopy(value)
            return True

        def reset_group(self, group: str) -> bool:
            assert group == "general"
            self.raw["gui"] = {
                "window": {
                    "remember_gui_state": True,
                    "auto_center": False,
                    "expand_side_panels": False,
                }
            }
            return True

        def reload(self) -> None:
            return None

    port = SuccessfulGeneralResetPort()
    vm = SettingsViewModel(controller=ApplicationController(config_port=port))  # type: ignore[arg-type]
    dialog = SettingsDialog(view_model=vm)
    monkeypatch.setattr(dialog_module, "_show_confirm", lambda *_args, **_kwargs: True)

    output_tab = dialog._tabs["output"]  # pyright: ignore[reportPrivateUsage]
    output_mode = output_tab._output_mode  # type: ignore[attr-defined]  # pyright: ignore[reportPrivateUsage]
    output_mode.setCurrentIndex(output_mode.findData("custom"))
    assert vm.config.output.output_mode == "custom"

    dialog._tab_widget.setCurrentIndex(TAB_KEYS.index("general"))  # pyright: ignore[reportPrivateUsage]
    dialog._on_reset_tab()  # pyright: ignore[reportPrivateUsage]

    assert vm.config.gui.auto_center is False
    assert vm.config.output.output_mode == "custom"
    assert output_mode.currentData() == "custom"
    assert vm.persisted_config.output.output_mode == "source"
    assert vm.is_dirty is True
    assert {change["field"] for change in vm.get_change_summary()} == {"output.output_mode"}
    general_tab = dialog._tabs["general"]  # pyright: ignore[reportPrivateUsage]
    assert general_tab._auto_center.isChecked() is False  # type: ignore[attr-defined]  # pyright: ignore[reportPrivateUsage]
    assert "1" in dialog._changes_label.text()  # pyright: ignore[reportPrivateUsage]

    if finish_action == "cancel":
        dialog._on_cancel()  # pyright: ignore[reportPrivateUsage]
        assert vm.config.gui.auto_center is False
        assert vm.config.output.output_mode == "source"
        assert port.raw["output"]["directory"]["mode"] == "source"  # type: ignore[index]
    else:
        dialog._on_ok()  # pyright: ignore[reportPrivateUsage]
        assert vm.persisted_config.gui.auto_center is False
        assert vm.persisted_config.output.output_mode == "custom"
        assert port.raw["output"]["directory"]["mode"] == "custom"  # type: ignore[index]
        assert vm.is_dirty is False


def test_settings_text_template_restore_does_not_create_unsaved_draft(qapp) -> None:
    from docwen_gui.models.settings_config import GUIConfig, SettingsConfig
    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings.text_tab import TextTab

    vm = SettingsViewModel(config=SettingsConfig(gui=GUIConfig(md_default_template="docx")))
    vm.begin_session()
    tab = TextTab(vm)
    consumed_feedback: list[object] = []
    tab._template_selector.template_selected.connect(  # pyright: ignore[reportPrivateUsage]
        lambda *_args: consumed_feedback.append(
            tab._template_selector.consume_last_selection_feedback()  # pyright: ignore[reportPrivateUsage]
        )
    )

    vm.set_templates({"docx": ["Document Template"], "xlsx": ["Workbook Template"]})
    qapp.processEvents()

    assert vm.config.gui.md_default_template == "docx"
    assert vm.selected_templates == {
        "docx": "Document Template",
        "xlsx": "Workbook Template",
    }
    assert vm.is_dirty is False
    assert vm.get_change_summary() == []
    assert len(consumed_feedback) == 2
    assert all(feedback is not None for feedback in consumed_feedback)

    tab.close()

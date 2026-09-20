from __future__ import annotations

import pytest

pytestmark = pytest.mark.gui


def _combo_values(combo) -> list[object]:
    return [combo.itemData(i) for i in range(combo.count())]


def test_logging_tab_updates_directory_and_runtime_path(qapp, tmp_path, monkeypatch: pytest.MonkeyPatch) -> None:
    from types import SimpleNamespace

    from docwen_gui.models.settings_config import SettingsConfig
    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings.logging_tab import LoggingTab

    monkeypatch.delenv("DOCWEN_LOG_DIR", raising=False)
    monkeypatch.delenv("DOCWEN_LOG_TO_TEMP", raising=False)
    monkeypatch.setattr(
        "docwen_runtime.logging.get_logging_runtime_state",
        lambda: SimpleNamespace(active_log_file=None, overridden_by_env=None, fallback_reason=None),
    )

    vm = SettingsViewModel(config=SettingsConfig())
    tab = LoggingTab(vm)

    assert _combo_values(tab._level_combo) == ["debug", "info", "warning", "error", "critical"]  # pyright: ignore[reportPrivateUsage]
    assert _combo_values(tab._dir_mode) == ["user", "temp", "custom"]  # pyright: ignore[reportPrivateUsage]
    assert _combo_values(tab._console_colorize) == ["auto", "always", "never"]  # pyright: ignore[reportPrivateUsage]

    tab.set_combo_data(tab._dir_mode, "custom")  # pyright: ignore[reportPrivateUsage]
    assert tab._dir_edit.isEnabled() is True  # pyright: ignore[reportPrivateUsage]

    custom_dir = str(tmp_path / "logs")
    tab._dir_edit.setText(custom_dir)  # pyright: ignore[reportPrivateUsage]
    tab._file_prefix_edit.setText("audit")  # pyright: ignore[reportPrivateUsage]
    tab.set_combo_data(tab._console_colorize, "never")  # pyright: ignore[reportPrivateUsage]

    assert vm.config.logging.directory_mode == "custom"
    assert vm.config.logging.directory == custom_dir
    assert vm.config.logging.file_prefix == "audit"
    assert vm.config.logging.console_colorize == "never"
    assert tab._resolved_path.text().endswith("audit.log")  # pyright: ignore[reportPrivateUsage]
    assert custom_dir in tab._resolved_path.text()  # pyright: ignore[reportPrivateUsage]

    tab._file_prefix_edit.setText("")  # pyright: ignore[reportPrivateUsage]
    assert tab.validate()


def test_logging_tab_disables_directory_controls_when_env_override_active(
    qapp, monkeypatch: pytest.MonkeyPatch
) -> None:
    from docwen_gui.models.settings_config import SettingsConfig
    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings.logging_tab import LoggingTab

    monkeypatch.setenv("DOCWEN_LOG_DIR", "C:/override/logs")
    vm = SettingsViewModel(config=SettingsConfig())
    tab = LoggingTab(vm)

    tab.set_combo_data(tab._dir_mode, "custom")  # pyright: ignore[reportPrivateUsage]

    assert tab._override_source.text() == "DOCWEN_LOG_DIR"  # pyright: ignore[reportPrivateUsage]
    assert tab._dir_mode.isEnabled() is False  # pyright: ignore[reportPrivateUsage]
    assert tab._dir_edit.isEnabled() is False  # pyright: ignore[reportPrivateUsage]
    assert tab._browse_btn.isEnabled() is False  # pyright: ignore[reportPrivateUsage]
    assert "环境变量" in tab._dir_notice.text()  # pyright: ignore[reportPrivateUsage]


@pytest.mark.parametrize("override", [None, "DOCWEN_LOG_TO_TEMP"])
def test_logging_tab_uses_runtime_state_for_actual_path(qapp, monkeypatch: pytest.MonkeyPatch, override) -> None:
    from types import SimpleNamespace

    from docwen_gui.models.settings_config import SettingsConfig
    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings.logging_tab import LoggingTab

    runtime_state = SimpleNamespace(
        active_log_file="C:/runtime/logs/live.log",
        overridden_by_env=override,
        fallback_reason="",
    )
    monkeypatch.setattr("docwen_runtime.logging.get_logging_runtime_state", lambda: runtime_state)
    monkeypatch.setattr("docwen_runtime.logging.resolve_log_file_path", lambda _config: "C:/guessed/docwen.log")
    monkeypatch.delenv("DOCWEN_LOG_DIR", raising=False)
    monkeypatch.delenv("DOCWEN_LOG_TO_TEMP", raising=False)

    vm = SettingsViewModel(config=SettingsConfig())
    tab = LoggingTab(vm)

    assert tab._resolved_path.text() == "C:/runtime/logs/live.log"  # pyright: ignore[reportPrivateUsage]
    assert tab._override_source.text() == (override or "")  # pyright: ignore[reportPrivateUsage]
    assert tab._dir_mode.isEnabled() is (override is None)  # pyright: ignore[reportPrivateUsage]


def test_logging_tab_path_buttons_use_settings_secondary_style(qapp) -> None:
    from PySide6.QtWidgets import QPushButton

    from docwen_gui.models.settings_config import SettingsConfig
    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings.logging_tab import LoggingTab

    tab = LoggingTab(SettingsViewModel(config=SettingsConfig()))

    copy_button = tab.findChild(QPushButton, "settingsLoggingCopyPathButton")
    open_button = tab.findChild(QPushButton, "settingsLoggingOpenDirectoryButton")

    assert copy_button is not None
    assert open_button is not None
    assert copy_button.property("class") == "secondary"
    assert open_button.property("class") == "secondary"


def test_logging_tab_uses_fluent_settings_checkbox(qapp) -> None:
    from qfluentwidgets import CheckBox as FluentCheckBox

    from docwen_gui.models.settings_config import SettingsConfig
    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings.logging_tab import LoggingTab

    tab = LoggingTab(SettingsViewModel(config=SettingsConfig()))

    assert isinstance(tab._enable_checkbox, FluentCheckBox)  # pyright: ignore[reportPrivateUsage]
    assert isinstance(tab._console_enable, FluentCheckBox)  # pyright: ignore[reportPrivateUsage]


@pytest.mark.parametrize("failed_query", [False, True])
def test_logging_tab_does_not_invent_an_actual_path_when_logging_failed(qapp, monkeypatch, failed_query):
    from types import SimpleNamespace

    from docwen_gui.models.settings_config import SettingsConfig
    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings.logging_tab import LoggingTab

    def runtime_state():
        if failed_query:
            raise OSError("log unavailable")
        return SimpleNamespace(active_log_file=None, overridden_by_env=None, fallback_reason="log unavailable")

    monkeypatch.setattr("docwen_runtime.logging.get_logging_runtime_state", runtime_state)
    tab = LoggingTab(SettingsViewModel(config=SettingsConfig()))
    assert tab._resolved_path.text() == ""  # pyright: ignore[reportPrivateUsage]
    assert tab._fallback_reason.text() == "log unavailable"  # pyright: ignore[reportPrivateUsage]


def test_profile_selection_keeps_logging_controls_available(qapp, tmp_path, monkeypatch):
    from docwen_gui.models.settings_config import SettingsConfig
    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings.logging_tab import LoggingTab
    from docwen_runtime.logging import init_logging
    from docwen_runtime.profile_paths import bind_process_profile

    monkeypatch.setenv("DOCWEN_DATA_DIR", str(tmp_path / "profile"))
    monkeypatch.delenv("DOCWEN_LOG_DIR", raising=False)
    monkeypatch.delenv("DOCWEN_LOG_TO_TEMP", raising=False)
    with bind_process_profile():
        init_logging({"logger": {"enable": False, "console_enable": False}})
        monkeypatch.setenv("DOCWEN_LOG_DIR", str(tmp_path / "later-edit"))
        tab = LoggingTab(SettingsViewModel(config=SettingsConfig()))
        assert tab._override_source.text() == ""  # pyright: ignore[reportPrivateUsage]
        assert tab._dir_mode.isEnabled()  # pyright: ignore[reportPrivateUsage]

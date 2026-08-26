from __future__ import annotations

import pytest

pytestmark = pytest.mark.gui


def _combo_values(combo) -> list[object]:
    return [combo.itemData(i) for i in range(combo.count())]


def _set_combo_data(combo, value: object) -> None:
    index = combo.findData(value)
    assert index >= 0
    combo.setCurrentIndex(index)


def test_output_tab_uses_runtime_config_values(qapp) -> None:
    from docwen_gui.models.settings_config import SettingsConfig
    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings.output_tab import OutputTab

    vm = SettingsViewModel(config=SettingsConfig())
    tab = OutputTab(vm)

    assert _combo_values(tab._output_mode) == ["source", "custom"]  # pyright: ignore[reportPrivateUsage]
    assert _combo_values(tab._date_format) == ["%Y-%m-%d", "%Y%m%d", "%Y年%m月%d日"]  # pyright: ignore[reportPrivateUsage]
    assert tab._custom_path.text() == vm.config.output.custom_path  # pyright: ignore[reportPrivateUsage]
    assert tab._save_intermediate.isChecked() is vm.config.output.save_intermediate_files  # pyright: ignore[reportPrivateUsage]


def test_output_tab_user_edits_update_view_model(qapp, monkeypatch: pytest.MonkeyPatch) -> None:
    import docwen_gui.widgets.settings.output_tab as output_tab_module
    from docwen_gui.models.settings_config import SettingsConfig
    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings.output_tab import OutputTab

    vm = SettingsViewModel(config=SettingsConfig())
    tab = OutputTab(vm)

    tab._save_intermediate.setChecked(True)  # pyright: ignore[reportPrivateUsage]
    _set_combo_data(tab._output_mode, "custom")  # pyright: ignore[reportPrivateUsage]
    tab._custom_path.setText("C:/Exports")  # pyright: ignore[reportPrivateUsage]
    tab._create_date_subfolder.setChecked(True)  # pyright: ignore[reportPrivateUsage]
    _set_combo_data(tab._date_format, "%Y%m%d")  # pyright: ignore[reportPrivateUsage]
    tab._auto_open_folder.setChecked(True)  # pyright: ignore[reportPrivateUsage]

    output = vm.config.output
    assert output.save_intermediate_files is True
    assert output.output_mode == "custom"
    assert output.custom_path == "C:/Exports"
    assert output.create_date_subfolder is True
    assert output.date_folder_format == "%Y%m%d"
    assert output.auto_open_folder is True

    monkeypatch.setattr(output_tab_module.QFileDialog, "getExistingDirectory", lambda *_args, **_kwargs: "D:/Done")
    tab._browse_path()  # pyright: ignore[reportPrivateUsage]

    assert vm.config.output.custom_path == "D:/Done"

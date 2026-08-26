from __future__ import annotations

import pytest

pytestmark = pytest.mark.gui


def _combo_values(combo) -> list[object]:
    return [combo.itemData(i) for i in range(combo.count())]


def test_export_tab_base64_mode_locks_ocr_placement(qapp) -> None:
    from docwen_gui.models.settings_config import SettingsConfig
    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings.export_tab import ExportTab

    vm = SettingsViewModel(config=SettingsConfig())
    tab = ExportTab(vm)

    assert _combo_values(tab._image_mode) == ["file", "base64"]  # pyright: ignore[reportPrivateUsage]
    assert _combo_values(tab._ocr_mode) == ["image_md", "main_md"]  # pyright: ignore[reportPrivateUsage]
    assert tab.get_combo_data(tab._ocr_mode) == "image_md"  # pyright: ignore[reportPrivateUsage]

    tab.set_combo_data(tab._image_mode, "base64")  # pyright: ignore[reportPrivateUsage]

    assert vm.config.export.image_mode == "base64"
    assert vm.config.export.ocr_mode == "main_md"
    assert tab.get_combo_data(tab._ocr_mode) == "main_md"  # pyright: ignore[reportPrivateUsage]
    assert tab._ocr_mode.isEnabled() is False  # pyright: ignore[reportPrivateUsage]

    tab.set_combo_data(tab._image_mode, "file")  # pyright: ignore[reportPrivateUsage]

    assert vm.config.export.image_mode == "file"
    assert vm.config.export.ocr_mode == "image_md"
    assert tab.get_combo_data(tab._ocr_mode) == "image_md"  # pyright: ignore[reportPrivateUsage]
    assert tab._ocr_mode.isEnabled() is True  # pyright: ignore[reportPrivateUsage]


def test_export_tab_writes_all_controls_to_view_model(qapp) -> None:
    from docwen_gui.i18n import t
    from docwen_gui.models.settings_config import SettingsConfig
    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings.export_tab import ExportTab

    vm = SettingsViewModel(config=SettingsConfig())
    tab = ExportTab(vm)

    tab.set_combo_data(tab._ocr_mode, "main_md")  # pyright: ignore[reportPrivateUsage]
    tab._ocr_title_enabled.setChecked(False)  # pyright: ignore[reportPrivateUsage]
    tab._ocr_title_text.setText("Custom OCR")  # pyright: ignore[reportPrivateUsage]
    tab._compress_enabled.setChecked(False)  # pyright: ignore[reportPrivateUsage]
    tab._compress_threshold.setValue(512)  # pyright: ignore[reportPrivateUsage]

    assert vm.config.export.ocr_mode == "main_md"
    assert vm.config.export.ocr_title_enabled is False
    assert vm.config.export.ocr_title_text == "Custom OCR"
    assert vm.config.export.base64_compress_enabled is False
    assert vm.config.export.base64_compress_threshold_kb == 512

    tab._reset_ocr_title()  # pyright: ignore[reportPrivateUsage]

    assert vm.config.export.ocr_title_text == t("conversion.ocr_output.blockquote_prefix")


def test_export_tab_uses_fluent_settings_checkboxes(qapp) -> None:
    from qfluentwidgets import CheckBox as FluentCheckBox

    from docwen_gui.models.settings_config import SettingsConfig
    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings.export_tab import ExportTab

    tab = ExportTab(SettingsViewModel(config=SettingsConfig()))

    assert isinstance(tab._ocr_title_enabled, FluentCheckBox)  # pyright: ignore[reportPrivateUsage]
    assert isinstance(tab._compress_enabled, FluentCheckBox)  # pyright: ignore[reportPrivateUsage]


class _FakeConfigPort:
    def __init__(self) -> None:
        self.set_calls: list[tuple[str, object]] = []

    def snapshot(self) -> dict[str, object]:
        return {
            "gui": {"language": {"locale": "zh_CN"}},
            "export": {
                "to_md_image_extraction_mode": "base64",
                "to_md_ocr_placement_mode": "main_md",
            },
            "conversion": {
                "export": {
                    "base64_compress_enabled": False,
                    "base64_compress_threshold_kb": 512,
                },
                "ocr_output": {
                    "show_blockquote_title": True,
                    "blockquote_title_override_by_locale": {"zh_CN": "旧标题"},
                },
            },
        }

    def set(self, key: str, value: object) -> bool:
        self.set_calls.append((key, value))
        return True

    def set_many(self, values: dict[str, object]) -> bool:
        self.set_calls.extend(values.items())
        return True


class _FakeController:
    def __init__(self) -> None:
        self.config_port = _FakeConfigPort()


def test_export_settings_round_trip_locale_title_override(monkeypatch: pytest.MonkeyPatch) -> None:
    from docwen_gui import i18n
    from docwen_gui.view_models.settings_vm import SECTION_EXPORT, SettingsViewModel

    monkeypatch.setattr(i18n, "get_locale", lambda: "zh_CN")
    controller = _FakeController()
    vm = SettingsViewModel(controller=controller)  # type: ignore[arg-type]

    assert vm.config.export.image_mode == "base64"
    assert vm.config.export.ocr_mode == "main_md"
    assert vm.config.export.ocr_title_text == "旧标题"
    assert vm.config.export.base64_compress_enabled is False
    assert vm.config.export.base64_compress_threshold_kb == 512

    vm.set_field(SECTION_EXPORT, "ocr_title_text", "新标题")
    assert vm.apply_settings() is True

    assert (
        "conversion.ocr_output.blockquote_title_override_by_locale.zh_CN",
        "新标题",
    ) in controller.config_port.set_calls

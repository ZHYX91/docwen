from __future__ import annotations

import pytest

pytestmark = pytest.mark.gui


def _combo_values(combo) -> list[object]:
    return [combo.itemData(i) for i in range(combo.count())]


def test_link_tab_updates_combo_fields_and_max_depth(qapp) -> None:
    from docwen_gui.models.settings_config import SettingsConfig
    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings.link_tab import LinkTab

    vm = SettingsViewModel(config=SettingsConfig())
    tab = LinkTab(vm)

    assert _combo_values(tab._image_link_style) == [  # pyright: ignore[reportPrivateUsage]
        "wiki_embed",
        "wiki_link",
        "markdown_embed",
        "markdown_link",
    ]
    assert _combo_values(tab._wiki_link_mode) == ["keep", "extract_text", "extract_address", "remove", "hyperlink"]  # pyright: ignore[reportPrivateUsage]
    assert _combo_values(tab._embed_md_file_mode) == ["keep", "extract_text", "remove", "embed"]  # pyright: ignore[reportPrivateUsage]

    tab.set_combo_data(tab._image_link_style, "markdown_link")  # pyright: ignore[reportPrivateUsage]
    tab.set_combo_data(tab._wiki_link_mode, "extract_text")  # pyright: ignore[reportPrivateUsage]
    tab.set_combo_data(tab._embed_md_file_mode, "remove")  # pyright: ignore[reportPrivateUsage]
    tab._max_depth.setValue(7)  # pyright: ignore[reportPrivateUsage]

    assert vm.config.link.image_link_style == "markdown_link"
    assert vm.config.link.wiki_link_mode == "extract_text"
    assert vm.config.link.embed_md_file_mode == "remove"
    assert vm.config.link.max_depth == 7


def test_link_tab_user_edits_update_all_view_model_fields(qapp) -> None:
    from docwen_gui.models.settings_config import SettingsConfig
    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings.link_tab import LinkTab

    vm = SettingsViewModel(config=SettingsConfig())
    tab = LinkTab(vm)

    cases = [
        ("_image_link_style", "image_link_style", "markdown_link"),
        ("_md_file_link_style", "md_file_link_style", "wiki_link"),
        ("_wiki_link_mode", "wiki_link_mode", "extract_text"),
        ("_md_link_mode", "markdown_link_mode", "remove"),
        ("_wiki_hyperlink_fallback", "wiki_hyperlink_fallback", "extract_address"),
        ("_markdown_hyperlink_fallback", "markdown_hyperlink_fallback", "extract_text"),
        ("_wiki_embed_image_mode", "wiki_embed_image_mode", "keep"),
        ("_md_embed_image_mode", "markdown_embed_image_mode", "extract_text"),
        ("_embed_md_file_mode", "embed_md_file_mode", "remove"),
    ]

    for widget_name, field_name, value in cases:
        combo = getattr(tab, widget_name)
        tab.set_combo_data(combo, value)
        assert getattr(vm.config.link, field_name) == value

    tab._max_depth.setValue(5)  # pyright: ignore[reportPrivateUsage]

    assert vm.config.link.max_depth == 5


@pytest.mark.parametrize("mode", ["keep", "extract_text", "extract_address", "remove", "hyperlink"])
def test_hyperlink_fallback_controls_are_independent_and_preserve_inactive_choices(qapp, mode: str) -> None:
    from docwen_gui.models.settings_config import SettingsConfig
    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings.link_tab import LinkTab

    vm = SettingsViewModel(config=SettingsConfig())
    tab = LinkTab(vm)
    assert vm.is_dirty is False
    assert vm.config.link.wiki_link_mode == vm.config.link.markdown_link_mode == "hyperlink"
    for combo in (tab._wiki_hyperlink_fallback, tab._markdown_hyperlink_fallback):
        assert _combo_values(combo) == ["keep", "extract_text", "extract_address"]
        assert combo.currentData() == "keep"
        assert combo.isEnabled()

    tab.set_combo_data(tab._wiki_hyperlink_fallback, "extract_address")
    tab.set_combo_data(tab._markdown_hyperlink_fallback, "extract_text")
    tab.set_combo_data(tab._wiki_link_mode, mode)
    assert tab._wiki_hyperlink_fallback.isEnabled() is (mode == "hyperlink")
    assert tab._markdown_hyperlink_fallback.isEnabled()
    tab.set_combo_data(tab._md_link_mode, mode)
    assert tab._markdown_hyperlink_fallback.isEnabled() is (mode == "hyperlink")
    assert vm.config.link.wiki_hyperlink_fallback == "extract_address"
    assert vm.config.link.markdown_hyperlink_fallback == "extract_text"

    tab.set_combo_data(tab._wiki_link_mode, "hyperlink")
    tab.set_combo_data(tab._md_link_mode, "hyperlink")
    assert tab._wiki_hyperlink_fallback.isEnabled()
    assert tab._markdown_hyperlink_fallback.isEnabled()
    assert tab._wiki_hyperlink_fallback.currentData() == "extract_address"
    assert tab._markdown_hyperlink_fallback.currentData() == "extract_text"


def test_link_reload_does_not_replay_changes_and_updates_fallback_availability(qapp) -> None:
    from PySide6.QtTest import QSignalSpy

    from docwen_gui.models.settings_config import SettingsConfig
    from docwen_gui.view_models.settings_vm import SECTION_LINK, SettingsViewModel
    from docwen_gui.widgets.settings.link_tab import LinkTab

    vm = SettingsViewModel(config=SettingsConfig())
    tab = LinkTab(vm)
    fields = {
        "wiki_link_mode": (tab._wiki_link_mode, "extract_address"),
        "markdown_link_mode": (tab._md_link_mode, "extract_text"),
        "wiki_hyperlink_fallback": (tab._wiki_hyperlink_fallback, "extract_text"),
        "markdown_hyperlink_fallback": (tab._markdown_hyperlink_fallback, "extract_address"),
    }
    spies = [QSignalSpy(combo.currentIndexChanged) for combo, _value in fields.values()]
    for field, (_combo, value) in fields.items():
        vm.set_field(SECTION_LINK, field, value)
    changes = vm.get_change_summary()
    tab.reload_from_config()
    tab.reload_from_config()

    assert vm.get_change_summary() == changes
    assert all(spy.count() == 0 for spy in spies)
    for combo, value in fields.values():
        assert combo.currentData() == value
    assert not tab._wiki_hyperlink_fallback.isEnabled()
    assert not tab._markdown_hyperlink_fallback.isEnabled()


def test_link_mode_and_fallback_help_have_info_buttons_and_line_breaks(qapp) -> None:
    from PySide6.QtWidgets import QToolButton

    from docwen_gui.models.settings_config import SettingsConfig
    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings.base_tab import SETTINGS_INFO_BUTTON_OBJECT_NAME
    from docwen_gui.widgets.settings.link_tab import LinkTab

    tab = LinkTab(SettingsViewModel(config=SettingsConfig()))
    info_buttons = tab.findChildren(QToolButton, SETTINGS_INFO_BUTTON_OBJECT_NAME)
    for combo in (
        tab._wiki_link_mode,
        tab._md_link_mode,
        tab._wiki_hyperlink_fallback,
        tab._markdown_hyperlink_fallback,
    ):
        assert "\n" in combo.toolTip()
        assert any(button.toolTip() == combo.toolTip() for button in info_buttons)

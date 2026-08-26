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
    assert _combo_values(tab._wiki_link_mode) == ["keep", "extract_text", "remove", "hyperlink"]  # pyright: ignore[reportPrivateUsage]
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

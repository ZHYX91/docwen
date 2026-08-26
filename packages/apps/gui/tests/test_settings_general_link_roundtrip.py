from __future__ import annotations

from pathlib import Path

import pytest

pytestmark = pytest.mark.gui

PROJECT_CONFIGS = Path(__file__).resolve().parents[4] / "configs"


def _new_view_model(user_dir: Path):
    from docwen_application.controller import ApplicationController
    from docwen_bundle.config_port import ConfigPortAdapter
    from docwen_gui.view_models.settings_vm import SettingsViewModel

    config_port = ConfigPortAdapter(base_dir=PROJECT_CONFIGS, user_dir=user_dir)
    view_model = SettingsViewModel(controller=ApplicationController(config_port=config_port))
    return view_model, config_port


def test_general_window_toggles_round_trip_through_apply(
    qapp,
    tmp_path: Path,
) -> None:
    from PySide6.QtTest import QSignalSpy

    from docwen_gui.view_models.settings_vm import SECTION_GUI
    from docwen_gui.widgets.settings.general_tab import GeneralTab

    user_dir = tmp_path / "user"
    view_model, config_port = _new_view_model(user_dir)
    tab = GeneralTab(view_model)
    assert view_model.is_dirty is False

    tab._remember_state.setChecked(False)  # pyright: ignore[reportPrivateUsage]
    tab._auto_center.setChecked(True)  # pyright: ignore[reportPrivateUsage]
    tab._expand_side_panels.setChecked(True)  # pyright: ignore[reportPrivateUsage]

    draft = view_model.config.gui
    assert draft.remember_gui_state is False
    assert draft.auto_center is True
    assert draft.expand_side_panels is True
    assert view_model.is_dirty is True
    assert {change["field"] for change in view_model.get_change_summary()} == {
        "gui.remember_gui_state",
        "gui.auto_center",
        "gui.expand_side_panels",
    }
    assert view_model.apply_settings() is True
    assert view_model.is_dirty is False
    assert view_model.get_change_summary() == []

    window = config_port.snapshot()["gui"]["window"]
    assert window["remember_gui_state"] is False
    assert window["auto_center"] is True
    assert window["expand_side_panels"] is True

    fresh_view_model, _fresh_port = _new_view_model(user_dir)
    fresh = fresh_view_model.config.gui
    assert fresh.remember_gui_state is False
    assert fresh.auto_center is True
    assert fresh.expand_side_panels is True

    # Reload must update all three widgets without replaying user-action
    # signals into the draft. Spying on the widget signals proves the
    # blockSignals boundary directly.
    view_model.set_field(SECTION_GUI, "remember_gui_state", True)
    view_model.set_field(SECTION_GUI, "auto_center", False)
    view_model.set_field(SECTION_GUI, "expand_side_panels", False)
    remember_spy = QSignalSpy(tab._remember_state.toggled)  # pyright: ignore[reportPrivateUsage]
    center_spy = QSignalSpy(tab._auto_center.toggled)  # pyright: ignore[reportPrivateUsage]
    expand_spy = QSignalSpy(tab._expand_side_panels.toggled)  # pyright: ignore[reportPrivateUsage]
    tab.reload_from_config()

    assert tab._remember_state.isChecked() is True  # pyright: ignore[reportPrivateUsage]
    assert tab._auto_center.isChecked() is False  # pyright: ignore[reportPrivateUsage]
    assert tab._expand_side_panels.isChecked() is False  # pyright: ignore[reportPrivateUsage]
    assert remember_spy.count() == 0
    assert center_spy.count() == 0
    assert expand_spy.count() == 0


def test_all_visible_link_fields_round_trip_through_apply(
    qapp,
    tmp_path: Path,
) -> None:
    from docwen_gui.widgets.settings.link_tab import LinkTab

    user_dir = tmp_path / "user"
    view_model, config_port = _new_view_model(user_dir)
    assert config_port.set_many(
        {
            "link.non_embed_links.auto_link_bare_url": False,
            "link.path_resolution.search_dirs": [".", "custom-assets"],
            "link.error_handling.file_not_found": "keep",
            "link.error_handling.detect_circular": False,
        }
    )
    tab = LinkTab(view_model)
    assert view_model.is_dirty is False

    combo_updates = (
        ("_image_link_style", "image_link_style", "markdown_link"),
        ("_md_file_link_style", "md_file_link_style", "wiki_link"),
        ("_wiki_link_mode", "wiki_link_mode", "extract_text"),
        ("_md_link_mode", "markdown_link_mode", "remove"),
        ("_wiki_embed_image_mode", "wiki_embed_image_mode", "keep"),
        ("_md_embed_image_mode", "markdown_embed_image_mode", "extract_text"),
        ("_embed_md_file_mode", "embed_md_file_mode", "remove"),
    )
    for widget_name, field_name, value in combo_updates:
        tab.set_combo_data(getattr(tab, widget_name), value)
        assert getattr(view_model.config.link, field_name) == value
    tab._max_depth.setValue(7)  # pyright: ignore[reportPrivateUsage]
    assert view_model.config.link.max_depth == 7
    assert {change["field"] for change in view_model.get_change_summary()} == {
        "link.image_link_style",
        "link.md_file_link_style",
        "link.wiki_link_mode",
        "link.markdown_link_mode",
        "link.wiki_embed_image_mode",
        "link.markdown_embed_image_mode",
        "link.embed_md_file_mode",
        "link.max_depth",
    }

    assert view_model.apply_settings() is True
    assert view_model.is_dirty is False
    assert view_model.get_change_summary() == []

    link = config_port.snapshot()["link"]
    assert link["format"] == {
        "image_link_style": "markdown_link",
        "md_file_link_style": "wiki_link",
    }
    assert link["non_embed_links"]["wiki_mode"] == "extract_text"
    assert link["non_embed_links"]["markdown_mode"] == "remove"
    assert link["embed_links"] == {
        "wiki_image_mode": "keep",
        "markdown_image_mode": "extract_text",
        "md_file_mode": "remove",
    }
    assert link["embedding"]["max_depth"] == 7
    assert link["non_embed_links"]["auto_link_bare_url"] is False
    assert link["path_resolution"]["search_dirs"] == [".", "custom-assets"]
    assert link["error_handling"]["file_not_found"] == "keep"
    assert link["error_handling"]["detect_circular"] is False

    fresh_view_model, _fresh_port = _new_view_model(user_dir)
    fresh = fresh_view_model.config.link
    assert fresh.image_link_style == "markdown_link"
    assert fresh.md_file_link_style == "wiki_link"
    assert fresh.wiki_link_mode == "extract_text"
    assert fresh.markdown_link_mode == "remove"
    assert fresh.wiki_embed_image_mode == "keep"
    assert fresh.markdown_embed_image_mode == "extract_text"
    assert fresh.embed_md_file_mode == "remove"
    assert fresh.max_depth == 7

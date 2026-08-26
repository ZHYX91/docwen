"""Focused tests split from test_settings_dialog_shell.py."""

from __future__ import annotations

import pytest

pytestmark = pytest.mark.gui


def test_settings_text_reset_restores_template_type_widget_without_dirty(qapp) -> None:
    from copy import deepcopy

    from docwen_application.controller import ApplicationController
    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings.text_tab import TextTab

    class SuccessfulTextResetPort:
        def __init__(self) -> None:
            self.raw: dict[str, object] = {
                "gui": {"template": {"md_default_template": "docx"}},
                "text": {
                    "remove_numbering": False,
                    "add_numbering": True,
                    "numbering_scheme": "chinese_chapters",
                    "heading_numbering_render_mode": "text",
                },
                "numbering": {"add": {"settings": {"default_scheme": "chinese_chapters"}}},
            }

        def snapshot(self) -> dict[str, object]:
            return deepcopy(self.raw)

        def reset_group(self, group: str) -> bool:
            assert group == "text"
            self.raw["gui"] = {"template": {"md_default_template": "docx"}}
            self.raw["text"] = {
                "remove_numbering": True,
                "add_numbering": False,
                "numbering_scheme": "hierarchical_standard",
                "heading_numbering_render_mode": "text",
            }
            self.raw["numbering"] = {
                "add": {"settings": {"default_scheme": "hierarchical_standard"}},
            }
            return True

        def reload(self) -> None:
            return None

    port = SuccessfulTextResetPort()
    vm = SettingsViewModel(
        controller=ApplicationController(config_port=port),  # type: ignore[arg-type]
    )
    vm.begin_session()
    tab = TextTab(vm)
    template_data = {"docx": ["Document Template"], "xlsx": ["Workbook Template"]}
    vm.set_templates(template_data)
    qapp.processEvents()
    assert tab._template_selector.current_tab == "docx"  # pyright: ignore[reportPrivateUsage]

    docx_selector = tab._template_selector.get_selector("docx")  # pyright: ignore[reportPrivateUsage]
    assert docx_selector is not None
    docx_selector.select_template("Document Template", selection_source="user")
    tab._template_selector._set_current_tab(  # pyright: ignore[reportPrivateUsage]
        "xlsx",
        emit_signal=True,
    )
    assert tab._template_selector.current_tab == "xlsx"  # pyright: ignore[reportPrivateUsage]
    assert vm.config.gui.md_default_template == "xlsx"
    assert vm.is_dirty is True

    xlsx_selector = tab._template_selector.get_selector("xlsx")  # pyright: ignore[reportPrivateUsage]
    assert xlsx_selector is not None
    xlsx_selector.select_template("Workbook Template", selection_source="user")
    assert tab._template_selector.current_tab == "xlsx"  # pyright: ignore[reportPrivateUsage]
    assert vm.config.gui.md_default_template == "xlsx"
    assert vm.is_dirty is True

    assert vm.reset_group("text") is True
    tab.reload_from_config()

    assert vm.config.gui.md_default_template == "docx"
    assert tab._template_selector.current_tab == "docx"  # pyright: ignore[reportPrivateUsage]
    assert vm.is_dirty is False
    assert vm.get_change_summary() == []

    vm.set_templates(template_data)
    qapp.processEvents()
    assert vm.config.gui.md_default_template == "docx"
    assert tab._template_selector.current_tab == "docx"  # pyright: ignore[reportPrivateUsage]
    assert vm.is_dirty is False

    tab._template_selector._set_current_tab(  # pyright: ignore[reportPrivateUsage]
        "xlsx",
        emit_signal=True,
    )
    assert vm.config.gui.md_default_template == "xlsx"
    assert tab._template_selector.current_tab == "xlsx"  # pyright: ignore[reportPrivateUsage]
    assert vm.is_dirty is True
    assert {change["field"] for change in vm.get_change_summary()} == {"gui.md_default_template"}

    tab.close()

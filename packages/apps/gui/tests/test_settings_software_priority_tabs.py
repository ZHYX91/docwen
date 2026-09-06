from __future__ import annotations

import pytest

pytestmark = pytest.mark.gui


def test_priority_lists_align_and_fit_their_rows_without_unused_space(qapp) -> None:
    from PySide6.QtCore import QPoint

    from docwen_gui.models.settings_config import SettingsConfig
    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings.priority_editor import SoftwarePriorityEditor
    from docwen_gui.widgets.settings.spreadsheet_tab import SpreadsheetTab

    tab = SpreadsheetTab(SettingsViewModel(config=SettingsConfig()))
    tab.resize(600, 900)
    tab.show()
    try:
        for _ in range(12):
            qapp.processEvents()
        editors = tab.findChildren(SoftwarePriorityEditor)
        assert len(editors) == 3
        lists = [editor.list_widget for editor in editors]
        assert len({widget.mapTo(tab, QPoint()).x() for widget in lists}) == 1
        assert len({widget.width() for widget in lists}) == 1
        assert lists[1].count() == 2 and lists[0].count() == lists[2].count() == 3
        assert lists[1].height() < lists[0].height() == lists[2].height()
        button_sizes = set()
        for editor in editors:
            widget = editor.list_widget
            last_row = widget.visualItemRect(widget.item(widget.count() - 1))
            assert 0 <= widget.viewport().height() - last_row.bottom() <= 4
            assert widget.verticalScrollBar().maximum() == 0
            assert editor.title_label.mapTo(editor, QPoint()).y() < widget.mapTo(editor, QPoint()).y()
            for button in (editor.move_up_button, editor.move_down_button):
                button_sizes.add((button.width(), button.height()))
        assert len(button_sizes) == 1
    finally:
        tab.close()


def test_document_software_priority_buttons_write_back(qapp) -> None:
    from docwen_gui.models.settings_config import SettingsConfig
    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings.document_tab import DocumentTab

    vm = SettingsViewModel(config=SettingsConfig())
    tab = DocumentTab(vm)

    priority_list = tab._priority_lists["word_processors"]  # pyright: ignore[reportPrivateUsage]
    move_up = tab._move_up_btns["word_processors"]  # pyright: ignore[reportPrivateUsage]
    move_down = tab._move_down_btns["word_processors"]  # pyright: ignore[reportPrivateUsage]

    assert priority_list.currentRow() == 0
    assert move_up.isEnabled() is False
    assert move_down.isEnabled() is True

    move_down.click()

    assert vm.config.software_priority.word_processors == ["msoffice_word", "wps_writer", "libreoffice"]
    assert priority_list.currentRow() == 1


def test_document_to_pdf_priority_buttons_write_back(qapp) -> None:
    from docwen_gui.models.settings_config import SettingsConfig
    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings.document_tab import DocumentTab

    vm = SettingsViewModel(config=SettingsConfig())
    tab = DocumentTab(vm)

    priority_list = tab._priority_lists["document_to_pdf"]  # pyright: ignore[reportPrivateUsage]
    move_down = tab._move_down_btns["document_to_pdf"]  # pyright: ignore[reportPrivateUsage]

    move_down.click()

    assert vm.config.software_priority.document_to_pdf == ["msoffice_word", "wps_writer", "libreoffice"]
    assert priority_list.currentRow() == 1


def test_layout_software_priority_buttons_write_back(qapp) -> None:
    from docwen_gui.models.settings_config import SettingsConfig
    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings.layout_tab import LayoutTab

    vm = SettingsViewModel(config=SettingsConfig())
    tab = LayoutTab(vm)

    priority_list = tab._priority_list  # pyright: ignore[reportPrivateUsage]
    move_up = tab._move_up_btn  # pyright: ignore[reportPrivateUsage]
    move_down = tab._move_down_btn  # pyright: ignore[reportPrivateUsage]

    assert priority_list.currentRow() == 0
    assert move_up.isEnabled() is False
    assert move_down.isEnabled() is True

    move_down.click()

    assert vm.config.software_priority.pdf_to_office == ["libreoffice", "msoffice_word"]
    assert priority_list.currentRow() == 1


def test_layout_pdf_to_office_priority_filters_wps_and_unknown_backends(qapp) -> None:
    from docwen_gui.models.settings_config import SettingsConfig
    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings.layout_tab import LayoutTab

    config = SettingsConfig()
    config.software_priority.pdf_to_office = [
        "wps_writer",
        "libreoffice",
        "unknown_backend",
        "msoffice_word",
        "libreoffice",
    ]
    vm = SettingsViewModel(config=config)

    assert vm.config.software_priority.pdf_to_office == ["libreoffice", "msoffice_word"]

    tab = LayoutTab(vm)
    priority_list = tab._priority_list  # pyright: ignore[reportPrivateUsage]
    priority_ids = [priority_list.item(index).data(0x0100) for index in range(priority_list.count())]

    assert priority_ids == ["libreoffice", "msoffice_word"]
    assert "wps_writer" not in priority_ids
    assert "unknown_backend" not in priority_ids

    config.software_priority.pdf_to_office = ["wps_writer"]
    vm.load_full_config(config)
    assert vm.config.software_priority.pdf_to_office == ["msoffice_word", "libreoffice"]


def test_spreadsheet_software_priority_buttons_write_back(qapp) -> None:
    from docwen_gui.models.settings_config import SettingsConfig
    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings.spreadsheet_tab import SpreadsheetTab

    vm = SettingsViewModel(config=SettingsConfig())
    tab = SpreadsheetTab(vm)

    priority_list = tab._priority_lists["spreadsheet_processors"]  # pyright: ignore[reportPrivateUsage]
    move_up = tab._move_up_btns["spreadsheet_processors"]  # pyright: ignore[reportPrivateUsage]
    move_down = tab._move_down_btns["spreadsheet_processors"]  # pyright: ignore[reportPrivateUsage]

    assert priority_list.currentRow() == 0
    assert move_up.isEnabled() is False
    assert move_down.isEnabled() is True

    move_down.click()

    assert vm.config.software_priority.spreadsheet_processors == [
        "msoffice_excel",
        "wps_spreadsheets",
        "libreoffice",
    ]
    assert priority_list.currentRow() == 1


def test_spreadsheet_to_pdf_priority_buttons_write_back(qapp) -> None:
    from docwen_gui.models.settings_config import SettingsConfig
    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings.spreadsheet_tab import SpreadsheetTab

    vm = SettingsViewModel(config=SettingsConfig())
    tab = SpreadsheetTab(vm)

    priority_list = tab._priority_lists["spreadsheet_to_pdf"]  # pyright: ignore[reportPrivateUsage]
    move_down = tab._move_down_btns["spreadsheet_to_pdf"]  # pyright: ignore[reportPrivateUsage]

    move_down.click()

    assert vm.config.software_priority.spreadsheet_to_pdf == [
        "msoffice_excel",
        "wps_spreadsheets",
        "libreoffice",
    ]
    assert priority_list.currentRow() == 1


def test_spreadsheet_ods_priority_filters_wps_and_unknown_backends(qapp) -> None:
    from docwen_gui.models.settings_config import SettingsConfig
    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings.spreadsheet_tab import SpreadsheetTab

    config = SettingsConfig()
    config.software_priority.ods_conversion = [
        "wps_spreadsheets",
        "libreoffice",
        "unknown_backend",
        "msoffice_excel",
        "libreoffice",
    ]
    vm = SettingsViewModel(config=config)

    assert vm.config.software_priority.ods_conversion == ["libreoffice", "msoffice_excel"]

    tab = SpreadsheetTab(vm)
    priority_list = tab._priority_lists["ods"]  # pyright: ignore[reportPrivateUsage]
    priority_ids = [priority_list.item(index).data(0x0100) for index in range(priority_list.count())]

    assert priority_ids == ["libreoffice", "msoffice_excel"]
    assert "wps_spreadsheets" not in priority_ids
    assert "unknown_backend" not in priority_ids

    config.software_priority.ods_conversion = ["wps_spreadsheets"]
    vm.load_full_config(config)
    assert vm.config.software_priority.ods_conversion == ["msoffice_excel", "libreoffice"]

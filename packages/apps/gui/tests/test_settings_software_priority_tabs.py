from __future__ import annotations

import pytest

pytestmark = pytest.mark.gui


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

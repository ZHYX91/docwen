"""External converter priorities in one settings page."""

from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QListWidgetItem

from ...i18n import t
from ...view_models.settings_vm import SECTION_SOFTWARE_PRIORITY, SettingsViewModel
from .base_tab import BaseSettingsTab
from .priority_editor import SoftwarePriorityEditor

_SOFTWARE_LABEL_KEYS = {
    "wps_writer": "settings.document.software.wps_writer",
    "msoffice_word": "settings.document.software.msoffice_word",
    "libreoffice": "settings.document.software.libreoffice",
    "wps_spreadsheets": "settings.spreadsheet.software.wps_spreadsheets",
    "msoffice_excel": "settings.spreadsheet.software.excel",
    "wps_presentation": "settings.software.wps_presentation",
    "msoffice_powerpoint": "settings.software.powerpoint",
}

_PRIORITY_LABEL_KEYS = {
    "word_processors": "settings.document.word_processors_label",
    "odt_conversion": "settings.document.odt_conversion_label",
    "document_to_pdf": "settings.document.document_to_pdf_label",
    "spreadsheet_processors": "settings.spreadsheet.spreadsheet_processors_label",
    "ods_conversion": "settings.spreadsheet.ods_conversion_label",
    "spreadsheet_to_pdf": "settings.spreadsheet.spreadsheet_to_pdf_label",
    "presentation_processors": "settings.software.presentation_processors",
    "pdf_to_office": "settings.layout.pdf_to_doc_label",
}


class SoftwareTab(BaseSettingsTab):
    def __init__(self, view_model: SettingsViewModel) -> None:
        self._vm = view_model
        self._editors: dict[str, SoftwarePriorityEditor] = {}
        super().__init__()
        self.reload_from_config()

    def _create_interface(self) -> None:
        self.set_tab_description(
            t(
                "settings.software.description",
                "Try installed software in this order. Availability also depends on the selected conversion.",
            )
        )
        _card, form = self.add_settings_card(
            t("settings.software.priority", "Software Priority"), object_name="softwarePriorityCard"
        )
        for key, label_key in _PRIORITY_LABEL_KEYS.items():
            editor = SoftwarePriorityEditor(t(label_key), self._scroll_container)
            self._editors[key] = editor
            editor.list_widget.currentRowChanged.connect(lambda _row, k=key: self._refresh_buttons(k))
            editor.move_up_button.clicked.connect(lambda _checked=False, k=key: self._move_item(k, -1))
            editor.move_down_button.clicked.connect(lambda _checked=False, k=key: self._move_item(k, 1))
            form.addRow(editor)

    def reload_from_config(self) -> None:
        for key, editor in self._editors.items():
            widget = editor.list_widget
            selected = widget.currentItem().data(Qt.ItemDataRole.UserRole) if widget.currentItem() else None
            widget.clear()
            for software in getattr(self._vm.config.software_priority, key):
                label_key = _SOFTWARE_LABEL_KEYS.get(software)
                label = t(label_key) if label_key else software
                item = QListWidgetItem(label)
                item.setData(Qt.ItemDataRole.UserRole, software)
                item.setToolTip(label)
                widget.addItem(item)
                if software == selected:
                    widget.setCurrentItem(item)
            if widget.currentRow() < 0 and widget.count():
                widget.setCurrentRow(0)
            self._refresh_buttons(key)

    def _refresh_buttons(self, key: str) -> None:
        editor = self._editors[key]
        row = editor.list_widget.currentRow()
        editor.move_up_button.setEnabled(row > 0)
        editor.move_down_button.setEnabled(0 <= row < editor.list_widget.count() - 1)

    def _move_item(self, key: str, offset: int) -> None:
        widget = self._editors[key].list_widget
        row = widget.currentRow()
        target = row + offset
        if row < 0 or not 0 <= target < widget.count():
            return
        item = widget.takeItem(row)
        widget.insertItem(target, item)
        widget.setCurrentRow(target)
        self._refresh_buttons(key)
        self._vm.set_field(
            SECTION_SOFTWARE_PRIORITY,
            key,
            [str(widget.item(i).data(Qt.ItemDataRole.UserRole)) for i in range(widget.count())],
        )

"""Spreadsheet settings tab — extraction options, merge settings, software priority.

Matches old SpreadsheetTab (DynamicSettingsTab + 3 software priority QListWidgets).
"""

from __future__ import annotations

from PySide6.QtWidgets import (
    QHBoxLayout,
    QListWidget,
    QListWidgetItem,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from ...i18n import t
from ...view_models.settings_vm import SECTION_SOFTWARE_PRIORITY, SettingsViewModel
from .base_tab import DynamicSettingsTab

_SS_SOFTWARE_LABEL_KEYS: dict[str, str] = {
    "wps_spreadsheets": "settings.spreadsheet.software.wps_spreadsheets",
    "msoffice_excel": "settings.spreadsheet.software.excel",
    "libreoffice": "settings.spreadsheet.software.libreoffice",
}


class SpreadsheetTab(DynamicSettingsTab):
    """Spreadsheet (XLSX/XLS/ODS/CSV) settings tab."""

    def __init__(self, view_model: SettingsViewModel) -> None:
        schema = [
            {
                "title": t("settings.spreadsheet.extraction_section", "Extraction Options"),
                "presentation": "card",
                "fields": [
                    {
                        "key": "to_md_keep_images",
                        "type": "checkbox",
                        "text": t("settings.spreadsheet.keep_images", "Keep images during conversion"),
                        "tooltip": t(
                            "settings.spreadsheet.keep_images_tooltip", "Extract embedded images from spreadsheets"
                        ),
                    },
                    {
                        "key": "to_md_enable_ocr",
                        "type": "checkbox",
                        "text": t("settings.spreadsheet.enable_ocr", "Enable OCR on images"),
                        "tooltip": t("settings.spreadsheet.enable_ocr_tooltip", "Perform OCR on extracted images"),
                    },
                ],
            },
            {
                "title": t("settings.table_export.section", "Merge Cell Export"),
                "description": t(
                    "settings.table_export.desc", "Controls how merged cells are represented in Markdown output."
                ),
                "presentation": "card",
                "fields": [
                    {
                        "key": "to_md_table_merge_export_strategy",
                        "type": "combobox",
                        "label": t("settings.table_export.merge_strategy_label", "Merge Cell Export Strategy:"),
                        "items": [
                            (t("settings.table_export.strategies.fill", "Fill"), "fill"),
                            (t("settings.table_export.strategies.empty", "Empty"), "empty"),
                            (t("settings.table_export.strategies.marker", "Marker"), "marker"),
                        ],
                    },
                ],
            },
            {
                "title": t("settings.spreadsheet.merge_mode_section", "Table Merge Mode"),
                "description": t(
                    "settings.spreadsheet.merge_mode_desc", "Default merge mode for combining multiple sheets."
                ),
                "presentation": "card",
                "fields": [
                    {
                        "key": "merge_mode",
                        "type": "combobox",
                        "label": t("settings.spreadsheet.default_merge_mode_label", "Default Merge Mode:"),
                        "items": [
                            (t("settings.spreadsheet.merge_modes.by_row", "By Row"), 1),
                            (t("settings.spreadsheet.merge_modes.by_column", "By Column"), 2),
                            (t("settings.spreadsheet.merge_modes.by_cell", "By Cell"), 3),
                        ],
                    },
                ],
            },
        ]
        self._vm = view_model
        self._priority_lists: dict[str, QListWidget] = {}
        self._move_up_btns: dict[str, QPushButton] = {}
        self._move_down_btns: dict[str, QPushButton] = {}
        super().__init__(None, "conversion_defaults", "spreadsheet", schema)
        self._load_values()
        self._create_software_priority_section()
        self._load_software_priority_values()

    def _load_values(self) -> None:
        data = self._vm.config.conversion_defaults.spreadsheet
        if data:
            self.load_values_from_dict(data)

    def reload_from_config(self) -> None:
        self._load_values()
        self._load_software_priority_values()

    def _create_software_priority_section(self) -> None:
        _card, form = self.add_settings_card(
            t("settings.spreadsheet.software_section", "Software Priority"),
            object_name="spreadsheetSoftwarePriorityCard",
        )
        categories = {
            "spreadsheet_processors": t("settings.spreadsheet.spreadsheet_processors_label", "Spreadsheet Processors:"),
            "ods": t("settings.spreadsheet.ods_conversion_label", "ODS Conversion:"),
            "spreadsheet_to_pdf": t("settings.spreadsheet.spreadsheet_to_pdf_label", "Spreadsheet to PDF:"),
        }
        for cat, label in categories.items():
            row = QWidget(self._scroll_container)
            row_layout = QHBoxLayout(row)
            row_layout.setContentsMargins(0, 0, 0, 0)
            row_layout.setSpacing(8)

            lst = QListWidget(row)
            lst.setObjectName("settingsPriorityList")
            lst.currentRowChanged.connect(lambda _r, c=cat: self._refresh_buttons(c))
            row_layout.addWidget(lst, 1)

            btn_container = QWidget(row)
            btn_layout = QVBoxLayout(btn_container)
            btn_layout.setContentsMargins(0, 0, 0, 0)
            btn_layout.setSpacing(6)

            up = QPushButton(t("editors.common.move_up", "Move Up"), btn_container)
            up.clicked.connect(lambda _checked=False, c=cat: self._move_item(c, -1))
            btn_layout.addWidget(up)

            down = QPushButton(t("editors.common.move_down", "Move Down"), btn_container)
            down.clicked.connect(lambda _checked=False, c=cat: self._move_item(c, 1))
            btn_layout.addWidget(down)
            btn_layout.addStretch(1)

            row_layout.addWidget(btn_container)
            self.add_form_row(form, label, row)
            self._priority_lists[cat] = lst
            self._move_up_btns[cat] = up
            self._move_down_btns[cat] = down
            self._refresh_buttons(cat)

    def _load_software_priority_values(self) -> None:
        sp = self._vm.config.software_priority
        defaults = {
            "spreadsheet_processors": sp.spreadsheet_processors,
            "ods": sp.ods_conversion,
            "spreadsheet_to_pdf": sp.spreadsheet_to_pdf,
        }
        for cat, lst in self._priority_lists.items():
            lst.clear()
            for sid in defaults.get(cat, []):
                label = t(_SS_SOFTWARE_LABEL_KEYS.get(sid, ""), sid)
                item = QListWidgetItem(label)
                item.setData(0x0100, sid)
                item.setToolTip(label)
                lst.addItem(item)
            if lst.count() > 0:
                lst.setCurrentRow(0)
            self._refresh_buttons(cat)

    def _get_priority(self, category: str) -> list[str]:
        lst = self._priority_lists[category]
        return [str(lst.item(i).data(0x0100)) for i in range(lst.count())]

    def _refresh_buttons(self, category: str) -> None:
        lst = self._priority_lists[category]
        row = lst.currentRow()
        cnt = lst.count()
        self._move_up_btns[category].setEnabled(row > 0)
        self._move_down_btns[category].setEnabled(0 <= row < cnt - 1)

    def _move_item(self, category: str, offset: int) -> None:
        lst = self._priority_lists[category]
        cur = lst.currentRow()
        tgt = cur + offset
        if cur < 0 or not (0 <= tgt < lst.count()):
            return
        item = lst.takeItem(cur)
        lst.insertItem(tgt, item)
        lst.setCurrentRow(tgt)
        self._refresh_buttons(category)

        mapping = {
            "spreadsheet_processors": "spreadsheet_processors",
            "ods": "ods_conversion",
            "spreadsheet_to_pdf": "spreadsheet_to_pdf",
        }
        key = mapping.get(category, category)
        self._vm.set_field(SECTION_SOFTWARE_PRIORITY, key, self._get_priority(category))

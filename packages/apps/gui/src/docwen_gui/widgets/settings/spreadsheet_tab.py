"""Spreadsheet settings tab — extraction options, merge settings, software priority.

Matches old SpreadsheetTab (DynamicSettingsTab + 3 software priority QListWidgets).
"""

from __future__ import annotations

from ...i18n import t
from ...view_models.settings_vm import SettingsViewModel
from .base_tab import DynamicSettingsTab


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
        super().__init__(None, "conversion_defaults", "spreadsheet", schema)
        self._load_values()

    def _load_values(self) -> None:
        data = self._vm.config.conversion_defaults.spreadsheet
        if data:
            self.load_values_from_dict(data)

    def reload_from_config(self) -> None:
        self._load_values()

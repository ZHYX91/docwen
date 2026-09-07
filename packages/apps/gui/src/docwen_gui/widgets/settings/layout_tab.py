"""Layout settings tab — extraction, optimization, DPI, software priority.

Matches old LayoutTab (DynamicSettingsTab + 1 software priority QListWidget).
"""

from __future__ import annotations

from ...i18n import t
from ...view_models.settings_vm import SettingsViewModel
from .base_tab import DynamicSettingsTab


class LayoutTab(DynamicSettingsTab):
    """Layout (PDF) settings tab."""

    def __init__(self, view_model: SettingsViewModel) -> None:
        schema = [
            {
                "title": t("settings.layout.extraction_section", "Extraction Options"),
                "presentation": "card",
                "fields": [
                    {
                        "key": "to_md_keep_images",
                        "type": "checkbox",
                        "text": t("settings.layout.keep_images", "Keep images during conversion"),
                        "tooltip": "",
                    },
                    {
                        "key": "to_md_enable_ocr",
                        "type": "checkbox",
                        "text": t("settings.layout.enable_ocr", "Enable OCR on images"),
                        "tooltip": "",
                    },
                ],
            },
            {
                "title": t("settings.layout.optimization_section", "Optimization"),
                "presentation": "card",
                "fields": [
                    {
                        "key": "to_md_enable_optimization",
                        "type": "checkbox",
                        "text": t("settings.layout.enable_optimization", "Enable content optimization"),
                        "tooltip": "",
                    },
                    {
                        "key": "to_md_optimization_type",
                        "type": "combobox",
                        "label": t("settings.layout.optimization_type_label", "Optimization Type:"),
                        "items": [],
                    },
                ],
            },
            {
                "title": t("settings.layout.dpi_section", "DPI Settings"),
                "description": t("settings.layout.dpi_desc", "Set render DPI for PDF to image conversion."),
                "presentation": "card",
                "fields": [
                    {
                        "key": "render_dpi",
                        "type": "combobox",
                        "label": t("settings.layout.render_dpi_label", "Render DPI:"),
                        "items": [
                            (t("settings.layout.dpi_min", "150 DPI (draft)"), 150),
                            (t("settings.layout.dpi_medium", "300 DPI (standard)"), 300),
                            (t("settings.layout.dpi_high", "600 DPI (high)"), 600),
                        ],
                    },
                ],
            },
        ]
        self._vm = view_model
        super().__init__(None, "conversion_defaults", "layout", schema)
        self._load_values()

    def _load_values(self) -> None:
        data = self._vm.config.conversion_defaults.layout
        if data:
            self.load_values_from_dict(data)

    def reload_from_config(self) -> None:
        self._load_values()

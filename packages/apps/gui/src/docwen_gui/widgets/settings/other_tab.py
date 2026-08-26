"""Other settings tab — catch-all conversion defaults for unclassified files.

Matches old OtherTab (DynamicSettingsTab, 2 checkboxes: keep_images, enable_ocr).
"""

from __future__ import annotations

from ...i18n import t
from ...view_models.settings_vm import SettingsViewModel
from .base_tab import DynamicSettingsTab


class OtherTab(DynamicSettingsTab):
    """Other (unclassified files) settings tab."""

    def __init__(self, view_model: SettingsViewModel) -> None:
        schema = [
            {
                "title": t("settings.other.extraction_section", "Extraction Options"),
                "presentation": "card",
                "fields": [
                    {
                        "key": "to_md_keep_images",
                        "type": "checkbox",
                        "text": t("settings.other.keep_images", "Keep images during conversion"),
                        "tooltip": t(
                            "settings.other.keep_images_tooltip", "Extract and save images from other file types"
                        ),
                    },
                    {
                        "key": "to_md_enable_ocr",
                        "type": "checkbox",
                        "text": t("settings.other.enable_ocr", "Enable OCR on images"),
                        "tooltip": t("settings.other.enable_ocr_tooltip", "Perform OCR on extracted images"),
                    },
                ],
            },
        ]
        self._vm = view_model
        super().__init__(None, "conversion_defaults", "other", schema)
        self._load_values()

    def _load_values(self) -> None:
        data = self._vm.config.conversion_defaults.other
        if data:
            self.load_values_from_dict(data)

    def reload_from_config(self) -> None:
        self._load_values()

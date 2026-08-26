"""Image settings tab — extraction, compression, PDF/TIFF options.

Matches old ImageTab (DynamicSettingsTab, 8 fields across 3 cards).
"""

from __future__ import annotations

from ...i18n import t
from ...view_models.settings_vm import SettingsViewModel
from .base_tab import DynamicSettingsTab


class ImageTab(DynamicSettingsTab):
    """Image (PNG/BMP/GIF/TIF/WebP/JPG) settings tab."""

    def __init__(self, view_model: SettingsViewModel) -> None:
        schema = [
            {
                "title": t("settings.image.extraction_section", "Extraction Options"),
                "presentation": "card",
                "fields": [
                    {
                        "key": "to_md_keep_images",
                        "type": "checkbox",
                        "text": t("settings.image.keep_images", "Keep images during conversion"),
                        "tooltip": "",
                    },
                    {
                        "key": "to_md_enable_ocr",
                        "type": "checkbox",
                        "text": t("settings.image.enable_ocr", "Enable OCR on images"),
                        "tooltip": "",
                    },
                    {
                        "key": "ocr_language",
                        "type": "combobox",
                        "label": t("settings.image.ocr_language_label", "OCR Language:"),
                        "tooltip": t(
                            "settings.image.ocr_language_tooltip", "Select OCR language for image text extraction"
                        ),
                        "items": [
                            (t("settings.image.ocr_language_auto", "Auto-detect"), "auto"),
                            (t("settings.image.ocr_language_chinese", "Chinese"), "chinese"),
                            (t("settings.image.ocr_language_chinese_cht", "Traditional Chinese"), "chinese_cht"),
                            (t("settings.image.ocr_language_english", "English Only"), "english"),
                            (t("settings.image.ocr_language_japanese", "Japanese"), "japanese"),
                            (t("settings.image.ocr_language_korean", "Korean"), "korean"),
                            (t("settings.image.ocr_language_latin", "Latin"), "latin"),
                            (t("settings.image.ocr_language_cyrillic", "Cyrillic"), "cyrillic"),
                        ],
                    },
                ],
            },
            {
                "title": t("settings.image.compress_section", "Compression"),
                "description": t("settings.image.compress_desc", "Configure image compression behavior."),
                "presentation": "card",
                "fields": [
                    {
                        "key": "compress_mode",
                        "type": "combobox",
                        "label": t("settings.image.compress_mode_label", "Compression Mode:"),
                        "items": [
                            (t("settings.image.compress_lossless", "Lossless"), "lossless"),
                            (t("settings.image.compress_limit_size", "Limit Size"), "limit_size"),
                        ],
                    },
                    {
                        "key": "size_limit",
                        "type": "spinbox",
                        "label": t("settings.image.size_limit_label", "Size Limit:"),
                        "min": 10,
                        "max": 10000,
                    },
                    {
                        "key": "size_unit",
                        "type": "combobox",
                        "label": t("settings.image.size_unit_label", "Unit:"),
                        "items": [("KB", "KB"), ("MB", "MB")],
                    },
                ],
            },
            {
                "title": t("settings.image.optimization_section", "Optimization"),
                "presentation": "card",
                "fields": [
                    {
                        "key": "to_md_enable_optimization",
                        "type": "checkbox",
                        "text": t("settings.image.enable_optimization", "Enable content optimization"),
                        "tooltip": "",
                    },
                    {
                        "key": "to_md_optimization_type",
                        "type": "combobox",
                        "label": t("settings.image.optimization_type_label", "Optimization Type:"),
                        "items": [],
                    },
                ],
            },
            {
                "title": t("settings.image.pdf_tiff_section", "PDF / TIFF Options"),
                "presentation": "card",
                "fields": [
                    {
                        "key": "pdf_quality",
                        "type": "combobox",
                        "label": t("settings.image.pdf_size_label", "PDF Page Size:"),
                        "items": [
                            (t("settings.image.pdf_original", "Original"), "original"),
                            (t("settings.image.pdf_a4", "Fit A4"), "fit_a4"),
                            (t("settings.image.pdf_a3", "Fit A3"), "fit_a3"),
                        ],
                    },
                    {
                        "key": "tiff_mode",
                        "type": "combobox",
                        "label": t("settings.image.tiff_mode_label", "TIFF Mode:"),
                        "items": [
                            (t("settings.image.tiff_keep_alpha", "Smart (keep alpha)"), "smart"),
                            (t("settings.image.tiff_no_alpha", "RGB (no alpha)"), "rgb"),
                        ],
                    },
                ],
            },
        ]
        self._vm = view_model
        super().__init__(None, "conversion_defaults", "image", schema)
        self._load_values()

    def _load_values(self) -> None:
        data = self._vm.config.conversion_defaults.image
        if data:
            self.load_values_from_dict(data)

    def reload_from_config(self) -> None:
        self._load_values()

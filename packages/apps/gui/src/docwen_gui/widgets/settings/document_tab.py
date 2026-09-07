"""Document settings tab — extraction options, optimization, software priority.

Matches old DocumentTab (DynamicSettingsTab + software priority QListWidgets).
"""

from __future__ import annotations

from PySide6.QtCore import QSignalBlocker
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
)

from ... import numbering_schemes
from ...i18n import t
from ...view_models.settings_vm import (
    SettingsViewModel,
)
from .base_tab import DynamicSettingsTab
from .content_controls import DocumentContentControls
from .numbering_editors import NumberingEditors


class DocumentTab(DynamicSettingsTab):
    """Document (DOCX/DOC/ODT/RTF) settings tab."""

    def __init__(self, view_model: SettingsViewModel) -> None:
        scheme_items = numbering_schemes.get_numbering_scheme_items(
            config_data=view_model.config.text.numbering_schemes,
        )
        schema = [
            {
                "title": t("settings.document.extraction_section", "Extraction Options"),
                "presentation": "card",
                "fields": [
                    {
                        "key": "to_md_keep_images",
                        "type": "checkbox",
                        "text": t("settings.document.keep_images", "Keep images during conversion"),
                        "tooltip": t(
                            "settings.document.keep_images_tooltip", "Extract and save embedded images from documents"
                        ),
                    },
                    {
                        "key": "to_md_enable_ocr",
                        "type": "checkbox",
                        "text": t("settings.document.enable_ocr", "Enable OCR on images"),
                        "tooltip": t(
                            "settings.document.enable_ocr_tooltip",
                            "Perform optical character recognition on extracted images",
                        ),
                    },
                ],
            },
            {
                "title": t("settings.document.numbering.section", "Heading Numbering (Document to Markdown)"),
                "presentation": "card",
                "fields": [
                    {
                        "key": "to_md_remove_numbering",
                        "type": "checkbox",
                        "text": t(
                            "settings.document.numbering.remove",
                            "Remove original document heading numbering by default",
                        ),
                        "tooltip": t(
                            "settings.document.numbering.remove_tooltip",
                            "Remove existing heading numbering when converting a document to Markdown",
                        ),
                    },
                    {
                        "key": "to_md_add_numbering",
                        "type": "checkbox",
                        "text": t(
                            "settings.document.numbering.add",
                            "Add heading numbering to Markdown by default",
                        ),
                    },
                    {
                        "key": "to_md_default_scheme",
                        "type": "combobox",
                        "label": t("settings.document.numbering.scheme_label", "Default numbering scheme:"),
                        "items": scheme_items,
                    },
                ],
            },
            {
                "title": t("settings.document.optimization_section", "Optimization"),
                "presentation": "card",
                "fields": [
                    {
                        "key": "to_md_enable_optimization",
                        "type": "checkbox",
                        "text": t("settings.document.enable_optimization", "Enable content optimization"),
                        "tooltip": "",
                    },
                    {
                        "key": "to_md_optimization_type",
                        "type": "combobox",
                        "label": t("settings.document.optimization_type_label", "Optimization Type:"),
                        "items": [],
                    },
                ],
            },
            {
                "title": t("settings.table_export.section", "Table Export"),
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
        ]
        self._vm = view_model
        super().__init__(None, "conversion_defaults", "document", schema)
        self._load_values()
        self._wire_numbering_controls()
        self._content_controls = DocumentContentControls(self, self._vm)
        self._numbering_editors = NumberingEditors(self, self._vm, self._refresh_scheme_combo_items)

    def _load_values(self) -> None:
        self._refresh_scheme_combo_items()
        data = self._vm.config.conversion_defaults.document
        if data:
            self.load_values_from_dict(data)

    def _refresh_scheme_combo_items(self) -> None:
        scheme = self._widgets.get("to_md_default_scheme")
        if isinstance(scheme, QComboBox):
            with QSignalBlocker(scheme):
                scheme.clear()
                for label, key in numbering_schemes.get_numbering_scheme_items(
                    config_data=self._vm.config.text.numbering_schemes
                ):
                    scheme.addItem(label, key)
                selected = self._vm.config.conversion_defaults.document.get(
                    "to_md_default_scheme", "hierarchical_standard"
                )
                self.set_combo_data(scheme, selected)

    def reload_from_config(self) -> None:
        self._content_controls.reload_from_config()
        self._load_values()
        self._sync_numbering_controls()

    def _wire_numbering_controls(self) -> None:
        add_numbering = self._widgets.get("to_md_add_numbering")
        scheme = self._widgets.get("to_md_default_scheme")
        if isinstance(add_numbering, QCheckBox) and isinstance(scheme, QComboBox):
            add_numbering.toggled.connect(scheme.setEnabled)
        self._sync_numbering_controls()

    def _sync_numbering_controls(self) -> None:
        add_numbering = self._widgets.get("to_md_add_numbering")
        scheme = self._widgets.get("to_md_default_scheme")
        if isinstance(add_numbering, QCheckBox) and isinstance(scheme, QComboBox):
            scheme.setEnabled(add_numbering.isChecked())

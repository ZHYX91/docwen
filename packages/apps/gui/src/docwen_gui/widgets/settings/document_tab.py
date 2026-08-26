"""Document settings tab — extraction options, optimization, software priority.

Matches old DocumentTab (DynamicSettingsTab + software priority QListWidgets).
"""

from __future__ import annotations

from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QHBoxLayout,
    QListWidget,
    QListWidgetItem,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from ... import numbering_schemes
from ...i18n import t
from ...view_models.settings_vm import (
    SECTION_SOFTWARE_PRIORITY,
    SettingsViewModel,
)
from .base_tab import DynamicSettingsTab

_SOFTWARE_LABEL_KEYS: dict[str, str] = {
    "wps_writer": "settings.document.software.wps_writer",
    "msoffice_word": "settings.document.software.msoffice_word",
    "libreoffice": "settings.document.software.libreoffice",
}


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
                            (t("settings.table_export.strategies.marker", "Marker"), "marker"),
                        ],
                    },
                ],
            },
        ]
        self._vm = view_model
        self._priority_lists: dict[str, QListWidget] = {}
        self._move_up_btns: dict[str, QPushButton] = {}
        self._move_down_btns: dict[str, QPushButton] = {}
        super().__init__(None, "conversion_defaults", "document", schema)
        self._load_values()
        self._wire_numbering_controls()
        self._create_software_priority_section()
        self._load_software_priority_values()

    def _load_values(self) -> None:
        data = self._vm.config.conversion_defaults.document
        if data:
            self.load_values_from_dict(data)

    def reload_from_config(self) -> None:
        self._load_values()
        self._sync_numbering_controls()
        self._load_software_priority_values()

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

    def _create_software_priority_section(self) -> None:
        _card, form = self.add_settings_card(
            t("settings.document.software_section", "Software Priority"),
            object_name="documentSoftwarePriorityCard",
        )
        categories = {
            "word_processors": t("settings.document.word_processors_label", "Word Processors:"),
            "odt": t("settings.document.odt_conversion_label", "ODT Conversion:"),
            "document_to_pdf": t("settings.document.document_to_pdf_label", "Document to PDF:"),
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
            "word_processors": sp.word_processors,
            "odt": sp.odt_conversion,
            "document_to_pdf": sp.document_to_pdf,
        }
        for cat, lst in self._priority_lists.items():
            lst.clear()
            for sid in defaults.get(cat, []):
                label = t(_SOFTWARE_LABEL_KEYS.get(sid, ""), sid)
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

        # Write back to VM
        values = self._get_priority(category)
        mapping = {
            "word_processors": "word_processors",
            "odt": "odt_conversion",
            "document_to_pdf": "document_to_pdf",
        }
        key = mapping.get(category, category)
        self._vm.set_field(SECTION_SOFTWARE_PRIORITY, key, values)

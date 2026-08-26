"""Layout settings tab — extraction, optimization, DPI, software priority.

Matches old LayoutTab (DynamicSettingsTab + 1 software priority QListWidget).
"""

from __future__ import annotations

from typing import cast as _cast

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

_LAYOUT_SOFTWARE_LABEL_KEYS: dict[str, str] = {
    "msoffice_word": "settings.document.software.msoffice_word",
    "libreoffice": "settings.document.software.libreoffice",
}


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
        self._priority_list: QListWidget = _cast(QListWidget, None)
        self._move_up_btn: QPushButton = _cast(QPushButton, None)
        self._move_down_btn: QPushButton = _cast(QPushButton, None)
        super().__init__(None, "conversion_defaults", "layout", schema)
        self._load_values()
        self._create_software_priority_section()
        self._load_software_priority_values()

    def _load_values(self) -> None:
        data = self._vm.config.conversion_defaults.layout
        if data:
            self.load_values_from_dict(data)

    def reload_from_config(self) -> None:
        self._load_values()
        self._load_software_priority_values()

    def _create_software_priority_section(self) -> None:
        _card, form = self.add_settings_card(
            t("settings.layout.software_section", "Software Priority"),
            object_name="layoutSoftwarePriorityCard",
        )
        self._priority_list = QListWidget(self._scroll_container)
        self._priority_list.setObjectName("settingsPriorityList")
        self._priority_list.currentRowChanged.connect(self._refresh_buttons)

        btn_container = QWidget(self._scroll_container)
        btn_layout = QVBoxLayout(btn_container)
        btn_layout.setContentsMargins(0, 0, 0, 0)
        btn_layout.setSpacing(6)

        self._move_up_btn = QPushButton(t("editors.common.move_up", "Move Up"), btn_container)
        self._move_up_btn.clicked.connect(lambda: self._move_item(-1))
        btn_layout.addWidget(self._move_up_btn)

        self._move_down_btn = QPushButton(t("editors.common.move_down", "Move Down"), btn_container)
        self._move_down_btn.clicked.connect(lambda: self._move_item(1))
        btn_layout.addWidget(self._move_down_btn)
        btn_layout.addStretch(1)

        priority_row = QWidget(self._scroll_container)
        priority_layout = QHBoxLayout(priority_row)
        priority_layout.setContentsMargins(0, 0, 0, 0)
        priority_layout.setSpacing(8)
        priority_layout.addWidget(self._priority_list, 1)
        priority_layout.addWidget(btn_container)

        self.add_form_row(form, t("settings.layout.pdf_to_doc_label", "PDF to Office Priority:"), priority_row)
        self._refresh_buttons()

    def _load_software_priority_values(self) -> None:
        if self._priority_list is None:
            return
        self._priority_list.clear()
        for sid in self._vm.config.software_priority.pdf_to_office:
            label = t(_LAYOUT_SOFTWARE_LABEL_KEYS.get(sid, ""), sid)
            item = QListWidgetItem(label)
            item.setData(0x0100, sid)
            item.setToolTip(label)
            self._priority_list.addItem(item)
        if self._priority_list.count() > 0:
            self._priority_list.setCurrentRow(0)
        self._refresh_buttons()

    def _get_priority(self) -> list[str]:
        if self._priority_list is None:
            return []
        return [str(self._priority_list.item(i).data(0x0100)) for i in range(self._priority_list.count())]

    def _refresh_buttons(self, _row: int | None = None) -> None:
        if self._priority_list is None:
            return
        row = self._priority_list.currentRow()
        cnt = self._priority_list.count()
        if self._move_up_btn:
            self._move_up_btn.setEnabled(row > 0)
        if self._move_down_btn:
            self._move_down_btn.setEnabled(0 <= row < cnt - 1)

    def _move_item(self, offset: int) -> None:
        if self._priority_list is None:
            return
        cur = self._priority_list.currentRow()
        tgt = cur + offset
        if cur < 0 or not (0 <= tgt < self._priority_list.count()):
            return
        item = self._priority_list.takeItem(cur)
        self._priority_list.insertItem(tgt, item)
        self._priority_list.setCurrentRow(tgt)
        self._refresh_buttons()
        self._vm.set_field(SECTION_SOFTWARE_PRIORITY, "pdf_to_office", self._get_priority())

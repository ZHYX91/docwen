"""Official-document recognition defaults and Word template field formatting."""

from __future__ import annotations

from PySide6.QtCore import QSignalBlocker

from docwen_gui.i18n import t
from docwen_gui.view_models.settings_vm import SECTION_TEXT, SettingsViewModel

from .base_tab import BaseSettingsTab
from .optimization_default import DefaultOptimizationControl


class GongwenTab(BaseSettingsTab):
    def __init__(self, view_model: SettingsViewModel) -> None:
        self._vm = view_model
        super().__init__()
        self._load_fields_value()
        self._vm.config_changed.connect(self._on_config_changed)
        self._vm.config_reloaded.connect(self._load_fields_value)

    def _create_interface(self) -> None:
        self.set_tab_description(t("settings.gongwen.description"))
        _to_markdown_card, to_markdown_form = self.add_settings_card(
            t("settings.gongwen.to_markdown_title"),
            t("settings.gongwen.to_markdown_description"),
            object_name="gongwenToMarkdownCard",
        )
        self._document_default = DefaultOptimizationControl(
            self._vm,
            category="document",
            resource_id="gongwen",
            resource_label=t("settings.gongwen.title"),
            parent=self._scroll_container,
        )
        to_markdown_form.addRow(self._document_default)
        self.add_settings_link(
            to_markdown_form,
            t("settings.gongwen.document_settings"),
            "document",
            object_name="gongwenDocumentSettingsButton",
        )
        _from_markdown_card, from_markdown_form = self.add_settings_card(
            t("settings.gongwen.from_markdown_title"),
            object_name="gongwenFromMarkdownCard",
        )
        row, self._gongwen_fields = self._create_toggle_with_info(
            t("settings.gongwen.fields_label"),
            t("settings.gongwen.fields_help"),
        )
        from_markdown_form.addRow(row)
        self._gongwen_fields.toggled.connect(
            lambda enabled: self._vm.set_field(SECTION_TEXT, "gongwen_fields_enabled", enabled)
        )

    def _load_fields_value(self) -> None:
        with QSignalBlocker(self._gongwen_fields):
            self._gongwen_fields.setChecked(self._vm.config.text.gongwen_fields_enabled)

    def _on_config_changed(self, section: str, key: str, _value: object) -> None:
        if section == SECTION_TEXT and key == "gongwen_fields_enabled":
            self._load_fields_value()

    def reload_from_config(self) -> None:
        self._document_default.reload_from_config()
        self._load_fields_value()

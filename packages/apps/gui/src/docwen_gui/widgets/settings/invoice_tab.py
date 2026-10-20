"""Chinese-invoice defaults for fixed-layout files and images."""

from __future__ import annotations

from docwen_gui.i18n import t
from docwen_gui.view_models.settings_vm import SettingsViewModel

from .base_tab import BaseSettingsTab
from .optimization_default import DefaultOptimizationControl


class InvoiceTab(BaseSettingsTab):
    def __init__(self, view_model: SettingsViewModel) -> None:
        self._vm = view_model
        super().__init__()

    def _create_interface(self) -> None:
        self.set_tab_description(t("settings.invoice.description"))
        _layout_card, layout_form = self.add_settings_card(
            t("settings.invoice.layout_title"),
            t("settings.invoice.layout_description"),
            object_name="invoiceLayoutToMarkdownCard",
        )
        self._layout_default = DefaultOptimizationControl(
            self._vm,
            category="layout",
            resource_id="invoice_cn",
            resource_label=t("settings.invoice.title"),
            parent=self._scroll_container,
        )
        layout_form.addRow(self._layout_default)
        self.add_settings_link(
            layout_form,
            t("settings.invoice.layout_settings"),
            "layout",
            object_name="invoiceLayoutSettingsButton",
        )
        _image_card, image_form = self.add_settings_card(
            t("settings.invoice.image_title"),
            t("settings.invoice.image_description"),
            object_name="invoiceImageToMarkdownCard",
        )
        self._image_default = DefaultOptimizationControl(
            self._vm,
            category="image",
            resource_id="invoice_cn",
            resource_label=t("settings.invoice.title"),
            parent=self._scroll_container,
        )
        image_form.addRow(self._image_default)
        self.add_settings_link(
            image_form,
            t("settings.invoice.ocr_settings"),
            "export",
            object_name="invoiceOcrSettingsButton",
        )

    def reload_from_config(self) -> None:
        self._layout_default.reload_from_config()
        self._image_default.reload_from_config()

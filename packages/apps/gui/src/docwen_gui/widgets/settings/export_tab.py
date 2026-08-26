"""Export settings tab — image mode, OCR placement, Base64 compression.

Matches old ExportTab: image_mode (file/base64), ocr_mode (image_md/main_md),
OCR title, Base64 compress toggle + threshold.
Base64 mode locks OCR mode to main_md.
"""

from __future__ import annotations

from typing import cast as _cast

from PySide6.QtWidgets import QCheckBox, QComboBox, QHBoxLayout, QLineEdit, QPushButton, QSpinBox, QWidget

from ...i18n import t
from ...view_models.settings_vm import SECTION_EXPORT, SettingsViewModel
from .base_tab import BaseSettingsTab, _create_info_button


class ExportTab(BaseSettingsTab):
    """Export/image extraction settings tab."""

    def __init__(self, view_model: SettingsViewModel) -> None:
        self._vm = view_model
        self._image_mode: QComboBox = _cast(QComboBox, None)
        self._ocr_mode: QComboBox = _cast(QComboBox, None)
        self._ocr_title_enabled: QCheckBox = _cast(QCheckBox, None)
        self._ocr_title_text: QLineEdit = _cast(QLineEdit, None)
        self._ocr_title_reset_btn: QPushButton = _cast(QPushButton, None)
        self._compress_enabled: QCheckBox = _cast(QCheckBox, None)
        self._compress_threshold: QSpinBox = _cast(QSpinBox, None)
        self._prev_ocr_mode: str = _cast(str, None)
        super().__init__()
        self._load_values()

    def _create_interface(self) -> None:
        _card, form = self.add_settings_card(
            t("settings.export.md_export_section", "MD Export Options"),
            t("settings.export.md_export_desc", "Configure how images are handled in Markdown output."),
            object_name="exportMdExportCard",
        )

        self._image_mode = self.create_combobox(
            [
                (t("settings.extraction.image_extraction_mode_file", "Save as File"), "file"),
                (t("settings.extraction.image_extraction_mode_base64", "Embed as Base64"), "base64"),
            ],
            t("settings.extraction.image_extraction_mode_tooltip", "Image extraction mode"),
        )
        self.add_form_row(
            form, t("settings.extraction.image_extraction_mode_label", "Image Extraction Mode:"), self._image_mode
        )

        self._ocr_mode = self.create_combobox(
            [
                (t("settings.extraction.ocr_placement_mode_image_md", "Image MD (per-image)"), "image_md"),
                (t("settings.extraction.ocr_placement_mode_main_md", "Main MD (inline)"), "main_md"),
            ],
            t("settings.extraction.ocr_placement_mode_tooltip", "OCR text placement mode"),
        )
        self.add_form_row(
            form, t("settings.extraction.ocr_placement_mode_label", "OCR Placement Mode:"), self._ocr_mode
        )

        # OCR title
        ocr_title_widget, self._ocr_title_enabled = self._create_toggle_with_info(
            t("settings.extraction.ocr_blockquote_title_enabled_label", "Enable OCR Blockquote Title"),
            t("settings.extraction.ocr_blockquote_title_enabled_tooltip", "Show blockquote title for OCR results"),
        )
        form.addRow(ocr_title_widget)
        self._ocr_title_enabled.toggled.connect(self._refresh_title_controls)
        self._ocr_title_enabled.toggled.connect(lambda v: self._vm.set_field(SECTION_EXPORT, "ocr_title_enabled", v))

        title_row = QWidget(self._scroll_container)
        title_layout = QHBoxLayout(title_row)
        title_layout.setContentsMargins(0, 0, 0, 0)
        title_layout.setSpacing(6)
        self._ocr_title_text = QLineEdit(self._scroll_container)
        self._ocr_title_text.setToolTip(
            t("settings.extraction.ocr_blockquote_title_text_tooltip", "Only affects the current UI language.")
        )
        self._ocr_title_text.textChanged.connect(lambda v: self._vm.set_field(SECTION_EXPORT, "ocr_title_text", v))
        title_layout.addWidget(self._ocr_title_text, 1)
        self._ocr_title_reset_btn = QPushButton(
            t("settings.extraction.ocr_blockquote_title_text_reset", "Reset"), self._scroll_container
        )
        self._ocr_title_reset_btn.clicked.connect(self._reset_ocr_title)
        title_layout.addWidget(self._ocr_title_reset_btn)
        self.add_form_row(
            form, t("settings.extraction.ocr_blockquote_title_text_label", "OCR Blockquote Title:"), title_row
        )

        # Base64 compress
        _card2, form2 = self.add_settings_card(
            t("settings.export.base64_compress_section", "Base64 Compression"),
            t("settings.export.base64_compress_desc", "Compress images before base64 embedding."),
            object_name="exportBase64CompressCard",
        )
        compress_widget, self._compress_enabled = self._create_toggle_with_info(
            t("settings.export.base64_compress_enabled", "Enable Base64 Compression"),
            t("settings.export.base64_compress_tooltip", "Compress images to reduce base64 output size"),
        )
        form2.addRow(compress_widget)
        self._compress_enabled.toggled.connect(self._refresh_compress_controls)
        self._compress_enabled.toggled.connect(
            lambda v: self._vm.set_field(SECTION_EXPORT, "base64_compress_enabled", v)
        )

        self._compress_threshold = self.create_spinbox(
            10,
            10000,
            t(
                "settings.export.base64_compress_threshold_tooltip",
                "Size threshold in KB (images above this are compressed)",
            ),
            default=100,
        )
        self.add_form_row(
            form2,
            t("settings.export.base64_compress_threshold_label", "Compression Threshold (KB):"),
            self._compress_threshold,
        )
        self._compress_threshold.valueChanged.connect(
            lambda v: self._vm.set_field(SECTION_EXPORT, "base64_compress_threshold_kb", v)
        )

        # Wire signals
        self._image_mode.currentIndexChanged.connect(self._on_image_mode_changed)
        self._ocr_mode.currentIndexChanged.connect(self._on_ocr_mode_changed)

    def _create_toggle_with_info(self, text: str, tooltip: str) -> tuple[QWidget, QCheckBox]:
        """Create a settings toggle with label, returning (widget, checkbox)."""
        container = QWidget(self._scroll_container)
        layout = QHBoxLayout(container)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(6)
        cb = self.create_settings_toggle(text, tooltip)
        layout.addWidget(cb)
        if tooltip:
            layout.addWidget(_create_info_button(tooltip, container))
        layout.addStretch(1)
        return container, cb

    def _on_image_mode_changed(self, _idx: int) -> None:
        mode = self.get_combo_data(self._image_mode)
        if mode:
            self._vm.set_field(SECTION_EXPORT, "image_mode", mode)
        self._refresh_ocr_mode_state(mode)

    def _on_ocr_mode_changed(self, _idx: int) -> None:
        mode = self.get_combo_data(self._ocr_mode)
        if mode:
            self._vm.set_field(SECTION_EXPORT, "ocr_mode", mode)
        self._refresh_title_controls()

    def _refresh_ocr_mode_state(self, mode: str | None = None) -> None:
        if mode is None:
            mode = self.get_combo_data(self._image_mode)
        if mode == "base64":
            cur = self.get_combo_data(self._ocr_mode)
            if cur != "main_md":
                self._prev_ocr_mode = cur  # pyright: ignore[reportAttributeAccessIssue]
                self.set_combo_data(self._ocr_mode, "main_md")
                self._vm.set_field(SECTION_EXPORT, "ocr_mode", "main_md")
            self._ocr_mode.setEnabled(False)
        else:
            self._ocr_mode.setEnabled(True)
            if self._prev_ocr_mode:
                self.set_combo_data(self._ocr_mode, self._prev_ocr_mode)
                self._vm.set_field(SECTION_EXPORT, "ocr_mode", self._prev_ocr_mode)
                self._prev_ocr_mode = None  # pyright: ignore[reportAttributeAccessIssue]
        self._refresh_title_controls()

    def _refresh_title_controls(self, _idx: int | None = None) -> None:
        allow = self.get_combo_data(self._ocr_mode) == "main_md"
        self._ocr_title_enabled.setEnabled(allow)
        self._refresh_title_text_controls()

    def _refresh_title_text_controls(self) -> None:
        enabled = self._ocr_title_enabled.isEnabled() and self._ocr_title_enabled.isChecked()
        if self._ocr_title_text:
            self._ocr_title_text.setEnabled(enabled)
        if self._ocr_title_reset_btn:
            self._ocr_title_reset_btn.setEnabled(enabled)

    def _reset_ocr_title(self) -> None:
        if self._ocr_title_text:
            self._ocr_title_text.setText(t("conversion.ocr_output.blockquote_prefix", "🖼️ **Image OCR**:"))

    def _refresh_compress_controls(self) -> None:
        if self._compress_threshold:
            self._compress_threshold.setEnabled(self._compress_enabled.isChecked())

    def _load_values(self) -> None:
        exp = self._vm.config.export
        self.set_combo_data(self._image_mode, exp.image_mode)
        self.set_combo_data(self._ocr_mode, exp.ocr_mode)
        self._ocr_title_enabled.setChecked(exp.ocr_title_enabled)
        if self._ocr_title_text:
            self._ocr_title_text.setText(exp.ocr_title_text)
        self._compress_enabled.setChecked(exp.base64_compress_enabled)
        if self._compress_threshold:
            self._compress_threshold.setValue(exp.base64_compress_threshold_kb)
        self._prev_ocr_mode = None  # pyright: ignore[reportAttributeAccessIssue]
        self._refresh_ocr_mode_state(exp.image_mode)
        self._refresh_compress_controls()

    def reload_from_config(self) -> None:
        self._load_values()

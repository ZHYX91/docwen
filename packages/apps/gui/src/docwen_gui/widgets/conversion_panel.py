"""ConversionPanel widget — format conversion panel with 4 category layouts.

Renders the conversion panel UI based on ConversionPanelViewModel state.
Does NOT call runtime/plugins directly — all actions go through the ViewModel.

Supports 4 category layouts:
  - Document: DOCX/DOC/ODT/RTF conversion + PDF save-as + proofread
  - Spreadsheet: XLSX/XLS/ODS/CSV conversion + PDF save-as + table merge
  - Image: PNG/BMP/GIF/TIF/WebP/JPG conversion + compression + PDF save-as + TIFF merge
  - Layout: PDF conversion + DOCX/DOC/ODT/RTF export + PNG/JPG/TIF render + PDF merge/split

Widget structure:
  - 3 neutral PanelCard sections: conversion, layout output, extra
  - Each group has one concise description plus its controls
"""

from __future__ import annotations

import logging
from collections.abc import Sequence
from pathlib import Path
from typing import TYPE_CHECKING
from typing import cast as _cast

from PySide6.QtCore import QEvent, QRectF, QSignalBlocker, Qt, QTimer, Signal
from PySide6.QtGui import QColor, QIcon, QPainter, QPixmap
from PySide6.QtWidgets import (
    QButtonGroup,
    QCheckBox,
    QComboBox,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QRadioButton,
    QScrollArea,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)

from docwen_gui.format_presentation import FormatChoice, presentation_for
from docwen_gui.i18n import t as _t
from docwen_gui.view_models.conversion_panel_vm import BUTTON_COLORS

from .panel_card import ActionFooter, ChoiceGroup, FormatSelector, FormRow, InlineNotice, PanelCard, WrappingLabel

if TYPE_CHECKING:
    from ..view_models.conversion_panel_vm import ConversionPanelViewModel

logger = logging.getLogger(__name__)

# ── Design constants ─────────────────────────────────────────────────────
_SPACING_XS = 4
_SPACING_SM = 8

# ── Format swatch icons ────────────────────────────────────────────────
# Semantic class per format comes from the shared format presentation registry.
_SWATCH_SIZE = 12
_SWATCH_RADIUS = 3

_format_icon_cache: dict[tuple[str, str], QIcon] = {}
_empty_swatch_icon: QIcon | None = None


def format_swatch_icon(format_name: str) -> QIcon | None:
    """Return a theme-aware 12x12 color dot icon for a known format name.

    Unknown formats return ``None``; the color is resolved from the current
    theme via :func:`docwen_gui.styles.theme_semantics.get_theme_class_color`.
    """
    normalized = str(format_name or "").strip().upper()
    theme_class = BUTTON_COLORS.get(normalized)
    if theme_class is None:
        return None

    from docwen_gui.styles.theme_manager import ThemeManager
    from docwen_gui.styles.theme_semantics import get_theme_class_color

    theme_name = ThemeManager.get_instance().get_current_theme()
    key = (theme_name, normalized)
    cached = _format_icon_cache.get(key)
    if cached is not None:
        return cached

    pixmap = QPixmap(_SWATCH_SIZE, _SWATCH_SIZE)
    pixmap.fill(Qt.GlobalColor.transparent)
    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing)
    painter.setPen(Qt.PenStyle.NoPen)
    painter.setBrush(QColor(get_theme_class_color(theme_class, theme_name)))
    painter.drawRoundedRect(
        QRectF(0.5, 0.5, _SWATCH_SIZE - 1, _SWATCH_SIZE - 1),
        _SWATCH_RADIUS,
        _SWATCH_RADIUS,
    )
    painter.end()
    icon = QIcon(pixmap)
    _format_icon_cache[key] = icon
    return icon


def apply_format_swatch_icons(combo: QComboBox | None) -> None:
    """Set format dots on a combo's items; no-op when no format is known.

    Unknown formats receive a transparent placeholder of the same size so
    icon and icon-less item text stays left-aligned inside the popup.
    Combos whose items are not formats (e.g. DPI or unit pickers) are left
    completely untouched.
    """
    if combo is None:
        return
    if not any(format_swatch_icon(combo.itemText(index)) is not None for index in range(combo.count())):
        return
    global _empty_swatch_icon
    if _empty_swatch_icon is None:
        placeholder = QPixmap(_SWATCH_SIZE, _SWATCH_SIZE)
        placeholder.fill(Qt.GlobalColor.transparent)
        _empty_swatch_icon = QIcon(placeholder)
    for index in range(combo.count()):
        combo.setItemIcon(index, format_swatch_icon(combo.itemText(index)) or _empty_swatch_icon)


class _WrappingCheckLabel(WrappingLabel):
    """Word-wrapping label that keeps the adjacent checkbox label-click affordance."""

    activated = Signal()

    def mouseReleaseEvent(self, event) -> None:
        if event.button() == Qt.MouseButton.LeftButton and self.isEnabled():
            self.activated.emit()
            event.accept()
            return
        super().mouseReleaseEvent(event)


class ConversionPanel(QWidget):
    """Format conversion panel widget.

    Renders 3 neutral cards (conversion/layout output/extra) whose content changes
    based on the current file category.  All user actions are delegated
    to the ``ConversionPanelViewModel``.
    """

    # Custom signal for focus chain integration
    panel_focused = Signal()

    def __init__(
        self,
        view_model: ConversionPanelViewModel,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self._vm = view_model
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        self.setObjectName("conversionPanelRoot")

        # Internal widget refs — cleared and rebuilt on category change
        self._conversion_combo: FormatSelector = _cast(FormatSelector, None)
        self._conversion_button: QPushButton = _cast(QPushButton, None)
        self._saveas_combo: QComboBox = _cast(QComboBox, None)
        self._saveas_button: QPushButton = _cast(QPushButton, None)
        self._validate_button: QPushButton = _cast(QPushButton, None)
        self._merge_tables_button: QPushButton = _cast(QPushButton, None)
        self._merge_tiff_button: QPushButton = _cast(QPushButton, None)

        # Layout section widgets
        self._layout_export_combo: QComboBox = _cast(QComboBox, None)
        self._layout_export_button: QPushButton = _cast(QPushButton, None)
        self._layout_render_format_combo: QComboBox = _cast(QComboBox, None)
        self._layout_render_dpi_combo: QComboBox = _cast(QComboBox, None)
        self._layout_render_button: QPushButton = _cast(QPushButton, None)
        self._merge_pdfs_button: QPushButton = _cast(QPushButton, None)
        self._split_pdf_button: QPushButton = _cast(QPushButton, None)

        # Image section sub-widgets
        self._size_limit_edit: QLineEdit = _cast(QLineEdit, None)
        self._size_unit_combo: QComboBox = _cast(QComboBox, None)
        self._size_warning_label: QLabel = _cast(QLabel, None)
        self._compress_btn_group: QButtonGroup = _cast(QButtonGroup, None)
        self._tiff_btn_group: QButtonGroup = _cast(QButtonGroup, None)
        self._pdf_quality_group: QButtonGroup = _cast(QButtonGroup, None)

        # Spreadsheet section sub-widgets
        self._merge_mode_group: QButtonGroup = _cast(QButtonGroup, None)
        self._reference_table_label: QLabel = _cast(QLabel, None)
        self._spreadsheet_password_edit: QLineEdit = _cast(QLineEdit, None)
        self._spreadsheet_protection_loss_checkbox: QCheckBox = _cast(QCheckBox, None)

        # Validation checkboxes
        self._validation_checkboxes: dict[str, QCheckBox] = {}

        # Layout section sub-widgets
        self._page_input_edit: QLineEdit = _cast(QLineEdit, None)
        self._pdf_info_label: QLabel = _cast(QLabel, None)
        self._page_warning_label: QLabel = _cast(QLabel, None)

        self._format_buttons: dict[str, QPushButton] = {}
        self._hint_label: QLabel = _cast(QLabel, None)
        self._render_key: tuple[object, ...] | None = None

        self._build_ui()
        self._wire_vm()

    def changeEvent(self, event: QEvent) -> None:
        super().changeEvent(event)
        if event.type() != QEvent.Type.PaletteChange:
            return
        for combo in self.findChildren(QComboBox):
            apply_format_swatch_icons(combo)

    # ── UI Construction ──────────────────────────────────────────────────

    def _build_ui(self) -> None:
        """Build the widget skeleton with three neutral operation cards."""
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        scroll_area = QScrollArea(self)
        scroll_area.setObjectName("conversionPanelScrollArea")
        scroll_area.setWidgetResizable(True)
        scroll_area.setFrameShape(QFrame.Shape.NoFrame)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)

        scroll_content = QWidget(scroll_area)
        scroll_content.setObjectName("conversionPanelScrollContent")
        content_layout = QVBoxLayout(scroll_content)
        content_layout.setContentsMargins(0, 0, 0, _SPACING_SM * 2)
        content_layout.setSpacing(_SPACING_SM)
        content_layout.setAlignment(Qt.AlignmentFlag.AlignTop)

        # Conversion group
        self._conversion_group = self._make_section_group(
            _t("conversion_panel.format_conversion", "Format Conversion"),
            "conversionPrimaryGroup",
        )
        content_layout.addWidget(self._conversion_group)

        # Save-as group
        self._saveas_group = self._make_section_group(
            _t("conversion_panel.save_as", "Save As"),
            "conversionSecondaryGroup",
        )
        content_layout.addWidget(self._saveas_group)

        # Extra group (initially hidden)
        self._extra_group = self._make_section_group("", "conversionExtraGroup")
        self._extra_group.setVisible(False)
        content_layout.addWidget(self._extra_group)

        # Hint label (shown when no file selected)
        self._hint_label = WrappingLabel(
            _t("conversion_panel.select_file_hint", "Select a file to see conversion options")
        )
        self._hint_label.setObjectName("hintLabel")
        self._hint_label.setWordWrap(True)
        content_layout.addWidget(self._hint_label)

        scroll_area.setWidget(scroll_content)
        root.addWidget(scroll_area)

    def _make_section_group(self, title: str, object_name: str) -> PanelCard:
        """Create a compact card with an internal centred title."""
        group = PanelCard(title, self, level="card")
        group.setObjectName(object_name)
        group.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Maximum)

        layout = group.content_layout
        layout.setSpacing(_SPACING_XS)
        layout.setAlignment(Qt.AlignmentFlag.AlignTop)

        desc_label = WrappingLabel(parent=group)
        desc_label.setObjectName("conversionSectionDescription")
        desc_label.setWordWrap(True)
        layout.addWidget(desc_label)
        setattr(self, "_" + object_name + "_desc", desc_label)

        content = QVBoxLayout()
        content.setSpacing(_SPACING_XS)
        content.setAlignment(Qt.AlignmentFlag.AlignTop)
        layout.addLayout(content)
        setattr(self, "_" + object_name + "_content", content)

        return group

    def _get_primary_content(self) -> QVBoxLayout:
        return self._conversionPrimaryGroup_content  # pyright: ignore[reportAttributeAccessIssue]

    def _get_secondary_content(self) -> QVBoxLayout:
        return self._conversionSecondaryGroup_content  # pyright: ignore[reportAttributeAccessIssue]

    def _get_extra_content(self) -> QVBoxLayout:
        return self._conversionExtraGroup_content  # pyright: ignore[reportAttributeAccessIssue]

    # ── ViewModel Wiring ─────────────────────────────────────────────────

    def _wire_vm(self) -> None:
        """Connect ViewModel signals to widget updates."""
        self._vm.state_changed.connect(self._on_vm_state_changed)
        self._rebuild_from_vm()

    def _current_render_key(self) -> tuple[object, ...]:
        """Return only the state that changes the panel's widget structure."""
        route_result = self._vm.route_choices_result
        return (
            self._vm.file_category,
            self._vm.current_format,
            self._vm.ui_mode,
            route_result.status,
            route_result.targets,
            getattr(route_result.error, "code", None),
            self._vm.spreadsheet_protection_info,
            self._vm.has_xlsx_inputs,
            self._vm.spreadsheet_analysis_pending,
        )

    def _on_vm_state_changed(self) -> None:
        """Rebuild only for structural context changes; otherwise update in place."""
        if self._render_key != self._current_render_key():
            # Analysis completes asynchronously. Preserve the user's target
            # when only protection controls changed for the same route context.
            target = (
                self._conversion_combo.currentText()
                if self._conversion_combo is not None
                and self._render_key is not None
                and self._render_key[:6] == self._current_render_key()[:6]
                else ""
            )
            self._rebuild_from_vm()
            if target and self._conversion_combo is not None:
                self._conversion_combo.setCurrentText(target)
            return
        self._sync_controls_from_vm()

    def _rebuild_from_vm(self) -> None:
        """Rebuild content after the file category, source format, or UI mode changes."""
        self._clear_all_content()
        self._update_section_descriptions()
        self._render_key = self._current_render_key()

        category = self._vm.file_category
        if category is None:
            self._conversion_group.setVisible(False)
            self._saveas_group.setVisible(False)
            self._extra_group.setVisible(False)
            if self._hint_label:
                self._hint_label.setVisible(True)
            return

        route_result = self._vm.route_choices_result
        if route_result.status != "ready":
            self._conversion_group.setVisible(False)
            self._saveas_group.setVisible(False)
            self._extra_group.setVisible(False)
            if self._hint_label:
                if route_result.status == "failed":
                    self._hint_label.setText(
                        _t(
                            "conversion_panel.route_options_unavailable",
                            "Available operations could not be loaded; actions are disabled to avoid an incorrect request.",
                        )
                    )
                else:
                    self._hint_label.setText(
                        _t(
                            "conversion_panel.no_compatible_operation",
                            "No compatible operation is available for this file.",
                        )
                    )
                self._hint_label.setProperty("routeState", route_result.status)
                self._hint_label.setVisible(True)
            return

        if self._hint_label:
            self._hint_label.setText(_t("conversion_panel.select_file_hint", "Select a file to see conversion options"))
            self._hint_label.setProperty("routeState", None)
            self._hint_label.setVisible(False)

        if category == "document":
            self._build_document_section()
        elif category == "spreadsheet":
            self._build_spreadsheet_section()
        elif category == "image":
            self._build_image_section()
        elif category == "layout":
            self._build_layout_section()
        for content in (self._get_primary_content(), self._get_secondary_content(), self._get_extra_content()):
            self._arrange_section_actions(content)
        self._sync_controls_from_vm()

    def _arrange_section_actions(self, layout: QVBoxLayout) -> None:
        """Use the same target/options/action reading order in each workflow card."""
        footers: list[QWidget] = []
        widgets = [item.widget() for index in range(layout.count()) if (item := layout.itemAt(index)) is not None]
        for widget in widgets:
            if isinstance(widget, ActionFooter):
                combo = widget.findChild(FormatSelector)
                if combo is not None:
                    widget.content_layout.removeWidget(combo)
                    layout.insertWidget(
                        layout.indexOf(widget), FormRow(_t("conversion_panel.target_format", "Target format"), combo)
                    )
                    notice = InlineNotice(parent=self, tone="warning")
                    notice.setObjectName("targetAvailabilityNotice")
                    layout.insertWidget(layout.indexOf(widget), notice)
                    combo.choices_changed.connect(lambda c=combo, n=notice: self._sync_target_notice(c, n))
                    self._sync_target_notice(combo, notice)
                footers.append(widget)
            elif isinstance(widget, QPushButton):
                footers.append(widget)
        for footer in dict.fromkeys(footers):
            layout.removeWidget(footer)
            layout.addWidget(footer)

    def _sync_target_notice(self, combo: FormatSelector, notice: InlineNotice) -> None:
        available = any(bool(combo.itemData(index, Qt.ItemDataRole.UserRole)) for index in range(combo.count()))
        notice.setVisible(not available)
        if not available:
            message = (
                _t("conversion_panel.no_common_target")
                if self._vm.ui_mode == "batch"
                else _t("conversion_panel.no_compatible_operation")
            )
            notice.setText(message)

    def _clear_all_content(self) -> None:
        """Remove all dynamic widgets from content layouts."""
        for obj_attr in (
            "_conversionPrimaryGroup_content",
            "_conversionSecondaryGroup_content",
            "_conversionExtraGroup_content",
        ):
            layout = getattr(self, obj_attr, None)
            if layout is None:
                continue
            while layout.count():
                item = layout.takeAt(0)
                if item is None:
                    continue
                w = item.widget()
                if w is not None:
                    w.setVisible(False)
                    w.setParent(None)
                    w.deleteLater()

        self._extra_group.setVisible(False)
        self._extra_group.setTitle("")
        self._conversion_group.setVisible(True)
        self._saveas_group.setVisible(True)
        self._format_buttons.clear()
        self._validation_checkboxes.clear()

        # Reset all internal refs — pyright: ignore reportAttributeAccessIssue
        # since _cast() declares fields as non-Optional for usage sites
        for _attr_name in (
            "_conversion_combo",
            "_conversion_button",
            "_saveas_combo",
            "_saveas_button",
            "_validate_button",
            "_merge_tables_button",
            "_merge_tiff_button",
            "_layout_export_combo",
            "_layout_export_button",
            "_layout_render_format_combo",
            "_layout_render_dpi_combo",
            "_layout_render_button",
            "_size_limit_edit",
            "_size_unit_combo",
            "_size_warning_label",
            "_compress_btn_group",
            "_tiff_btn_group",
            "_pdf_quality_group",
            "_merge_mode_group",
            "_reference_table_label",
            "_spreadsheet_password_edit",
            "_spreadsheet_protection_loss_checkbox",
            "_merge_pdfs_button",
            "_split_pdf_button",
            "_page_input_edit",
            "_pdf_info_label",
            "_page_warning_label",
        ):
            setattr(self, _attr_name, None)  # pyright: ignore[reportAttributeAccessIssue]

    def _update_section_descriptions(self) -> None:
        """Update the concise, action-specific descriptions for all sections."""
        cat = self._vm.file_category
        descs: dict[str, tuple[str, str, str]] = {
            "document": (
                _t("conversion_panel.document.convert_to_other_formats", "Convert to other document formats"),
                _t("conversion_panel.document.convert_to_layout", "Save as layout format"),
                _t("conversion_panel.document.proofread_hint", "Proofread document"),
            ),
            "spreadsheet": (
                _t("conversion_panel.spreadsheet.convert_to_other_formats", "Convert to other spreadsheet formats"),
                _t("conversion_panel.spreadsheet.convert_to_layout", "Save as layout format"),
                _t("conversion_panel.spreadsheet.merge_tables_hint", "Merge tables"),
            ),
            "image": (
                _t("conversion_panel.image.convert_to_other_formats", "Convert to other image formats"),
                _t("conversion_panel.image.convert_to_layout", "Save as layout format"),
                _t("conversion_panel.image.merge_images_hint", "Merge images"),
            ),
            "layout": (
                _t("conversion_panel.layout.convert_to_other_formats", "Convert layout"),
                _t("conversion_panel.layout.convert_to_document_or_image", "Export or render"),
                _t("conversion_panel.layout.merge_split_hint", "Merge or split PDFs"),
            ),
        }
        defaults: tuple[str, str, str] = ("", "", "")
        primary, secondary, extra = descs.get(cat or "", defaults)
        if cat == "layout" and self._vm.ui_mode == "single":
            extra = ""
        for prefix, text in [
            ("conversionPrimaryGroup", primary),
            ("conversionSecondaryGroup", secondary),
            ("conversionExtraGroup", extra),
        ]:
            desc_label = getattr(self, "_" + prefix + "_desc", None)
            if desc_label:
                desc_label.setText(text)
                desc_label.setVisible(bool(text))

    @staticmethod
    def _combo_items(combo: QComboBox) -> list[str]:
        return [combo.itemText(index) for index in range(combo.count())]

    @classmethod
    def _sync_combo_items(
        cls,
        combo: QComboBox | None,
        items: list[str],
        *,
        selected: str | None = None,
    ) -> None:
        """Synchronize a combo without disturbing a still-valid user selection."""
        if combo is None:
            return
        previous = combo.currentText()
        target = selected if selected in items else previous if previous in items else items[0] if items else ""
        with QSignalBlocker(combo):
            if cls._combo_items(combo) != items:
                combo.clear()
                combo.addItems(items)
                apply_format_swatch_icons(combo)
            if target:
                combo.setCurrentText(target)

    @staticmethod
    def _sync_button_id(group: QButtonGroup | None, button_id: int) -> None:
        if group is None:
            return
        button = group.button(button_id)
        if button is None:
            return
        with QSignalBlocker(group):
            button.setChecked(True)

    @staticmethod
    def _sync_button_property(group: QButtonGroup | None, property_name: str, value: object) -> None:
        if group is None:
            return
        with QSignalBlocker(group):
            for button in group.buttons():
                if button.property(property_name) == value:
                    button.setChecked(True)
                    return

    def _sync_controls_from_vm(self) -> None:
        """Project non-structural ViewModel state onto existing controls in place."""
        self._sync_action_targets()
        self._sync_document_controls()
        self._sync_image_controls()
        self._sync_spreadsheet_controls()
        self._sync_layout_controls()
        self._sync_aggregate_actions()

    def _sync_aggregate_actions(self) -> None:
        for action, button in (
            ("merge_tables", self._merge_tables_button),
            ("merge_images_to_tiff", self._merge_tiff_button),
            ("merge_pdfs", self._merge_pdfs_button),
        ):
            if button is not None:
                count = self._vm.aggregate_count(action)
                reason = _t("main_window.aggregate_need_two", count=count) if count < 2 else ""
                button.setEnabled(count >= 2)
                button.setToolTip(reason)
                notice = self._extra_group.findChild(QLabel, "aggregateAvailabilityNotice")
                if notice is None:
                    notice = WrappingLabel(parent=self._extra_group)
                    notice.setObjectName("aggregateAvailabilityNotice")
                    notice.setWordWrap(True)
                    layout = self._get_extra_content()
                    anchor: QWidget = button
                    if layout.indexOf(anchor) < 0:
                        anchor = button.parentWidget() or button
                    layout.insertWidget(layout.indexOf(anchor) + 1, notice)
                notice.setText(reason)
                notice.setVisible(bool(reason))

    def _sync_action_targets(self) -> None:
        if self._conversion_combo is not None:
            self._conversion_combo.set_choices(self._vm.get_conversion_format_choices())
            apply_format_swatch_icons(self._conversion_combo)
        self._sync_combo_items(self._saveas_combo, self._vm.get_saveas_formats())
        self._sync_combo_items(self._layout_export_combo, self._vm.get_layout_export_formats())
        self._sync_combo_items(
            self._layout_render_format_combo,
            self._vm.get_layout_render_formats(),
            selected=self._vm.render_format,
        )
        self._sync_combo_items(
            self._layout_render_dpi_combo,
            ["150", "300", "600"],
            selected=str(self._vm.render_dpi),
        )

        self._sync_conversion_button_enabled()
        if self._saveas_button is not None and self._saveas_combo is not None:
            self._saveas_button.setEnabled(self._saveas_combo.count() > 0)
        if self._layout_export_button is not None and self._layout_export_combo is not None:
            self._layout_export_button.setEnabled(self._layout_export_combo.count() > 0)
        if self._layout_render_button is not None and self._layout_render_format_combo is not None:
            self._layout_render_button.setEnabled(self._layout_render_format_combo.count() > 0)

    def _sync_document_controls(self) -> None:
        validation_options = self._vm.validation_options
        for key, checkbox in self._validation_checkboxes.items():
            with QSignalBlocker(checkbox):
                checkbox.setChecked(validation_options.get(key, False))
        if self._validate_button is not None:
            self._validate_button.setEnabled(self._vm.is_any_validation_option_checked)

    def _sync_image_controls(self) -> None:
        is_limit = self._vm.compress_mode == "limit_size"
        self._sync_button_id(self._compress_btn_group, 1 if is_limit else 0)
        if self._size_limit_edit is not None:
            self._size_limit_edit.setEnabled(is_limit)
            if not self._size_limit_edit.hasFocus():
                with QSignalBlocker(self._size_limit_edit):
                    self._size_limit_edit.setText(str(self._vm.size_limit))
        if self._size_unit_combo is not None:
            self._size_unit_combo.setEnabled(is_limit)
            self._sync_combo_items(self._size_unit_combo, ["KB", "MB"], selected=self._vm.size_unit)
        if not is_limit and self._size_warning_label is not None:
            self._size_warning_label.clear()

        quality_ids = {"original": 0, "a4": 1, "a3": 2}
        self._sync_button_id(self._pdf_quality_group, quality_ids.get(self._vm.pdf_quality, 0))
        self._sync_button_property(self._tiff_btn_group, "tiff_value", self._vm.tiff_mode)

    def _sync_spreadsheet_controls(self) -> None:
        self._sync_button_id(self._merge_mode_group, self._vm.merge_mode)
        if self._reference_table_label is not None:
            self._reference_table_label.setText(
                _t(
                    "conversion_panel.spreadsheet.reference_table",
                    "Reference table: {name}",
                    name=self._vm.reference_table_name,
                )
                if self._vm.reference_table_name
                else _t("conversion_panel.spreadsheet.no_table_selected", "No table selected")
            )

    def _sync_conversion_button_enabled(self) -> None:
        """Apply target availability and protection consent as one gate."""

        button = self._conversion_button
        combo = self._conversion_combo
        if button is None or combo is None:
            return
        enabled = combo.current_choice_enabled()
        if (
            enabled
            and self._vm.file_category == "spreadsheet"
            and self._vm.has_xlsx_inputs
            and combo.currentText().strip().lower() == "ods"
        ):
            if self._vm.spreadsheet_analysis_pending or self._vm.spreadsheet_unknown_files:
                enabled = False
            elif self._vm.spreadsheet_protected_files:
                enabled = (
                    self._vm.ui_mode == "single"
                    and self._spreadsheet_protection_loss_checkbox is not None
                    and self._spreadsheet_protection_loss_checkbox.isChecked()
                    and (
                        not self._vm.spreadsheet_password_required
                        or (
                            self._spreadsheet_password_edit is not None and bool(self._spreadsheet_password_edit.text())
                        )
                    )
                )
        button.setEnabled(enabled)

    def _sync_layout_controls(self) -> None:
        if self._page_input_edit is not None:
            if not self._page_input_edit.hasFocus():
                with QSignalBlocker(self._page_input_edit):
                    self._page_input_edit.setText(self._vm.page_input)
            self._update_page_validation(self._page_input_edit.text().strip())
        if self._pdf_info_label is not None:
            self._pdf_info_label.setText(
                _t(
                    "conversion_panel.layout.selected_split_file",
                    "Selected file: {pages} pages",
                ).format(pages=self._vm.pdf_total_pages)
            )
            self._pdf_info_label.setToolTip(self._vm.pdf_file_name)
            self._pdf_info_label.setVisible(self._vm.pdf_total_pages > 0)

    # ── Widget factory helpers ───────────────────────────────────────────

    def _make_button_row(self) -> tuple[QWidget, QHBoxLayout]:
        row = QWidget(self)
        row.setObjectName("conversionButtonRow")
        layout = QHBoxLayout(row)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(_SPACING_SM)
        return row, layout

    def _make_combo(self, items: list[str], parent: QWidget | None = None) -> QComboBox:
        combo = QComboBox(parent or self)
        combo.addItems(items)
        apply_format_swatch_icons(combo)
        combo.setMinimumWidth(100)
        return combo

    def _make_action_button(self, text: str, parent: QWidget | None = None) -> QPushButton:
        btn = QPushButton(text, parent or self)
        btn.setObjectName("conversionSecondaryButton")
        btn.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        return btn

    def _make_subtle_divider(self) -> QFrame:
        """Create a subtle divider QFrame for visual separation."""
        sep = QFrame(self)
        sep.setObjectName("conversionSubtleDivider")
        sep.setFrameShape(QFrame.Shape.HLine)
        sep.setFrameShadow(QFrame.Shadow.Sunken)
        return sep

    @staticmethod
    def _enabled_format_choice(name: str) -> FormatChoice:
        presentation = presentation_for(name)
        return FormatChoice(
            key=presentation.key,
            display_name=presentation.display_name,
            tone=presentation.tone,
            enabled=True,
            disabled_reason="",
            help_text=_t(
                "conversion_panel.format_target_help",
                "Convert to {format}",
                format=presentation.display_name,
            ),
        )

    def _make_dropdown_action_row(
        self,
        items: Sequence[str | FormatChoice],
        button_text: str,
    ) -> tuple[QWidget, FormatSelector, QPushButton]:
        """Create a row with combo + action button."""
        footer = ActionFooter(self)
        choices = [item if isinstance(item, FormatChoice) else self._enabled_format_choice(item) for item in items]
        combo = FormatSelector(footer)
        combo.set_choices(choices)
        apply_format_swatch_icons(combo)
        footer.content_layout.addWidget(combo, alignment=Qt.AlignmentFlag.AlignVCenter)
        btn = self._make_action_button(button_text, parent=footer)
        btn.setEnabled(combo.current_choice_enabled())
        combo.currentIndexChanged.connect(lambda _index, c=combo, b=btn: b.setEnabled(c.current_choice_enabled()))
        footer.content_layout.addWidget(btn, alignment=Qt.AlignmentFlag.AlignVCenter)
        return footer, combo, btn

    def _make_checkbox(self, text: str, checked: bool = False) -> QCheckBox:
        try:
            from qfluentwidgets import CheckBox as FluentCheckBox

            cb = FluentCheckBox(text, self)
        except ImportError:
            cb = QCheckBox(text, self)
        cb.setChecked(checked)
        return cb

    def _make_wrapping_checkbox(self, text: str, checked: bool = False) -> tuple[QWidget, QCheckBox]:
        row = QWidget(self)
        row.setObjectName("conversionWrappingCheckRow")
        layout = QHBoxLayout(row)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(_SPACING_SM)

        checkbox = self._make_checkbox("", checked=checked)
        checkbox.setAccessibleName(text)
        checkbox.setToolTip(text)
        layout.addWidget(checkbox, alignment=Qt.AlignmentFlag.AlignTop)

        label = _WrappingCheckLabel(text, row)
        label.setObjectName("conversionWrappingCheckLabel")
        label.setWordWrap(True)
        label.setBuddy(checkbox)
        label.setToolTip(text)
        label.activated.connect(checkbox.toggle)
        layout.addWidget(label, stretch=1)
        QTimer.singleShot(0, label, label.sync_wrapped_height)
        return row, checkbox

    def _make_radio(self, text: str, checked: bool = False) -> QRadioButton:
        rb = QRadioButton(text, self)
        rb.setObjectName("conversionRadioButton")
        rb.setChecked(checked)
        return rb

    def _make_option_group(self, title: str) -> PanelCard:
        group = PanelCard(title, self, level="section")
        group.setObjectName("conversionOptionGroup")
        layout = group.content_layout
        layout.setSpacing(_SPACING_SM)
        return group

    # ── Document Section ────────────────────────────────────────────────

    def _build_document_section(self) -> None:
        """Build the document category: conversion row + save-as row + proofread extra."""
        conv_layout = self._get_primary_content()

        # Conversion row
        conversion_formats = self._vm.get_conversion_format_choices()
        row_container, combo, btn = self._make_dropdown_action_row(
            conversion_formats, _t("conversion_panel.convert", "Convert")
        )
        self._conversion_combo = combo
        self._conversion_button = btn
        btn.clicked.connect(self._on_convert_button_clicked)
        conv_layout.addWidget(row_container)

        # Save-as row
        saveas_layout = self._get_secondary_content()
        saveas_formats = self._vm.get_saveas_formats()
        row_container2, combo2, btn2 = self._make_dropdown_action_row(
            saveas_formats, _t("conversion_panel.save_as_action", "Save As")
        )
        self._saveas_combo = combo2
        self._saveas_button = btn2
        btn2.clicked.connect(self._on_saveas_button_clicked)
        saveas_layout.addWidget(row_container2)

        # Extra: proofread
        self._build_document_proofread_extra()

    def _build_document_proofread_extra(self) -> None:
        """Build the proofread extra section for document category."""
        extra_layout = self._get_extra_content()
        self._extra_group.setTitle(_t("conversion_panel.document.proofread_document", "Proofread Document"))
        self._extra_group.setVisible(True)

        validate_btn = self._make_action_button(_t("conversion_panel.document.proofread_button", "Proofread"))
        validate_btn.setEnabled(self._vm.is_any_validation_option_checked)
        validate_btn.clicked.connect(self._on_validate_clicked)
        self._validate_button = validate_btn
        extra_layout.addWidget(validate_btn)

        # Options group
        opts_group = self._make_option_group(_t("conversion_panel.document.proofread_options", "Proofread Options"))
        opts_layout = QVBoxLayout()
        opts_group.content_layout.addLayout(opts_layout)

        options_spec: list[tuple[str, str, str]] = [
            (
                "symbol_pairing",
                _t("conversion_panel.document.symbol_pairing", "Symbol Pairing"),
                _t("conversion_panel.document.symbol_pairing_tooltip", "Check bracket/quote pairing"),
            ),
            (
                "typos_rule",
                _t("conversion_panel.document.typos_rule", "Typos Rule"),
                _t("conversion_panel.document.typos_rule_tooltip", "Check common typos"),
            ),
            (
                "symbol_correction",
                _t("conversion_panel.document.symbol_correction", "Symbol Correction"),
                _t("conversion_panel.document.symbol_correction_tooltip", "Correct symbol usage"),
            ),
            (
                "sensitive_word",
                _t("conversion_panel.document.sensitive_word", "Sensitive Word"),
                _t("conversion_panel.document.sensitive_word_tooltip", "Check sensitive words"),
            ),
        ]

        opts = self._vm.validation_options
        self._validation_checkboxes.clear()
        for key, label, tooltip in options_spec:
            cb = self._make_checkbox(label, checked=opts.get(key, False))
            cb.setToolTip(tooltip)
            cb.stateChanged.connect(lambda state, k=key: self._on_validation_option_changed(k, bool(state)))
            self._validation_checkboxes[key] = cb
            opts_layout.addWidget(cb)

        extra_layout.addWidget(opts_group)

    # ── Spreadsheet Section ─────────────────────────────────────────────

    def _build_spreadsheet_section(self) -> None:
        """Build the spreadsheet category: conversion row + save-as row + table merge extra."""
        conv_layout = self._get_primary_content()

        conversion_formats = self._vm.get_conversion_format_choices()
        row_container, combo, btn = self._make_dropdown_action_row(
            conversion_formats, _t("conversion_panel.convert", "Convert")
        )
        self._conversion_combo = combo
        self._conversion_button = btn
        btn.clicked.connect(self._on_convert_button_clicked)
        conv_layout.addWidget(row_container)

        if self._vm.has_xlsx_inputs:
            policy_group = self._make_option_group(
                _t("conversion_panel.spreadsheet.ods_delivery_options", "ODS Delivery Options")
            )
            policy_layout = policy_group.content_layout
            is_single = self._vm.ui_mode == "single"
            protected_files = self._vm.spreadsheet_protected_files
            unknown_files = self._vm.spreadsheet_unknown_files

            policy_hint = QLabel(
                _t(
                    "conversion_panel.spreadsheet.ods_delivery_hint",
                    "External formulas use cached values. Publishing protected files as ODS requires explicit consent.",
                ),
                self,
            )
            policy_hint.setObjectName("conversionDetailLabel")
            policy_hint.setWordWrap(True)
            policy_layout.addWidget(policy_hint)

            if self._vm.spreadsheet_analysis_pending:
                policy_layout.addWidget(InlineNotice(_t("conversion_panel.spreadsheet.protection_checking"), self))
            elif unknown_files:
                unknown_notice = InlineNotice(
                    "\n".join(
                        (
                            _t(
                                "conversion_panel.spreadsheet.protection_status_unknown",
                                "The file's protection status could not be inspected reliably. "
                                "Conversion is blocked to prevent an unprotected delivery.",
                            ),
                            *(f"• {Path(path).name}" for path in unknown_files),
                        )
                    ),
                    self,
                    tone="warning",
                )
                policy_layout.addWidget(unknown_notice)

            elif is_single and protected_files:
                password_edit = QLineEdit(self)
                password_edit.setObjectName("spreadsheetProtectionPasswordEdit")
                password_edit.setEchoMode(QLineEdit.EchoMode.Password)
                password_edit.setClearButtonEnabled(True)
                password_edit.setPlaceholderText(
                    _t("conversion_panel.spreadsheet.protection_password_required", "Password (required)")
                    if self._vm.spreadsheet_password_required
                    else _t("conversion_panel.spreadsheet.protection_password", "Protection password (optional)")
                )
                password_edit.setAccessibleName(password_edit.placeholderText())
                password_edit.setToolTip(password_edit.placeholderText())
                self._spreadsheet_password_edit = password_edit
                password_edit.textChanged.connect(self._sync_conversion_button_enabled)
                policy_layout.addWidget(password_edit)

                consent_row, consent = self._make_wrapping_checkbox(
                    _t(
                        "conversion_panel.spreadsheet.allow_protection_loss",
                        "I understand the delivered ODS will not retain password protection",
                    ),
                    checked=False,
                )
                self._spreadsheet_protection_loss_checkbox = consent
                policy_layout.addWidget(consent_row)
                consent.toggled.connect(self._sync_conversion_button_enabled)

            elif not is_single and protected_files:
                batch_hint = QLabel(
                    "\n".join(
                        (
                            _t(
                                "conversion_panel.spreadsheet.protected_batch_single_only",
                                "Passwords are not reused across a batch. Convert protected files individually.",
                            ),
                            *(f"• {Path(path).name}" for path in protected_files),
                        )
                    ),
                    self,
                )
                batch_hint.setObjectName("warningLabel")
                batch_hint.setWordWrap(True)
                policy_layout.addWidget(batch_hint)

            conv_layout.addWidget(policy_group)
            policy_group.setVisible(combo.currentText().strip().upper() == "ODS")
            combo.currentTextChanged.connect(
                lambda value, group=policy_group: group.setVisible(value.strip().upper() == "ODS")
            )
            combo.currentTextChanged.connect(self._sync_conversion_button_enabled)

        saveas_layout = self._get_secondary_content()
        saveas_formats = self._vm.get_saveas_formats()
        if saveas_formats:
            row_container2, combo2, btn2 = self._make_dropdown_action_row(
                saveas_formats, _t("conversion_panel.save_as_action", "Save As")
            )
            self._saveas_combo = combo2
            self._saveas_button = btn2
            btn2.clicked.connect(self._on_saveas_button_clicked)
            saveas_layout.addWidget(row_container2)
        else:
            self._saveas_group.setVisible(False)

        self._build_spreadsheet_merge_extra()

    def _build_spreadsheet_merge_extra(self) -> None:
        """Build the table merge extra section for spreadsheet category."""
        if self._vm.ui_mode != "batch":
            self._extra_group.hide()
            return
        extra_layout = self._get_extra_content()
        self._extra_group.setTitle(_t("conversion_panel.spreadsheet.merge_tables", "Merge Tables"))
        self._extra_group.setVisible(True)

        merge_btn = self._make_action_button(_t("conversion_panel.spreadsheet.merge_tables_button", "Merge Tables"))
        merge_btn.clicked.connect(self._on_merge_tables_clicked)
        self._merge_tables_button = merge_btn
        extra_layout.addWidget(merge_btn)

        # Merge mode radios
        mode_group = self._make_option_group(_t("conversion_panel.spreadsheet.merge_options", "Merge Options"))
        mode_layout = mode_group.content_layout

        self._merge_mode_group = QButtonGroup(mode_group)
        self._merge_mode_group.setExclusive(True)

        radio_defs: list[tuple[str, int]] = [
            (_t("conversion_panel.spreadsheet.merge_by_row", "By Row"), 1),
            (_t("conversion_panel.spreadsheet.merge_by_column", "By Column"), 2),
            (_t("conversion_panel.spreadsheet.merge_by_cell", "By Cell"), 3),
        ]
        for label, value in radio_defs:
            rb = self._make_radio(label, checked=value == self._vm.merge_mode)
            self._merge_mode_group.addButton(rb, value)
            mode_layout.addWidget(rb)

        self._merge_mode_group.idToggled.connect(self._on_merge_mode_changed)
        extra_layout.addWidget(mode_group)

        # Reference table label
        ref_text = (
            _t(
                "conversion_panel.spreadsheet.reference_table",
                "Reference table: {name}",
                name=self._vm.reference_table_name,
            )
            if self._vm.reference_table_name
            else _t("conversion_panel.spreadsheet.no_table_selected", "No table selected")
        )
        ref_label = QLabel(ref_text, self)
        ref_label.setObjectName("conversionDetailLabel")
        ref_label.setWordWrap(True)
        self._reference_table_label = ref_label
        extra_layout.addWidget(ref_label)

    # ── Image Section ───────────────────────────────────────────────────

    def _build_image_section(self) -> None:
        """Build the image category: conversion row + compression + save-as + TIFF merge."""
        conv_layout = self._get_primary_content()

        # Conversion row
        conversion_formats = self._vm.get_conversion_format_choices()
        row_container, combo, btn = self._make_dropdown_action_row(
            conversion_formats, _t("conversion_panel.convert", "Convert")
        )
        self._conversion_combo = combo
        self._conversion_button = btn
        btn.clicked.connect(self._on_convert_button_clicked)
        conv_layout.addWidget(row_container)

        # Compression options
        self._build_image_compress_section(conv_layout)

        # Save-as row
        saveas_layout = self._get_secondary_content()
        saveas_formats = self._vm.get_saveas_formats()
        row_container2, combo2, btn2 = self._make_dropdown_action_row(
            saveas_formats, _t("conversion_panel.save_as_action", "Save As")
        )
        self._saveas_combo = combo2
        self._saveas_button = btn2
        btn2.clicked.connect(self._on_convert_to_pdf_clicked)
        saveas_layout.addWidget(row_container2)

        # PDF quality options
        self._build_pdf_quality_section(saveas_layout)

        # TIFF merge extra
        self._build_tiff_merge_extra()

    def _build_image_compress_section(self, parent_layout: QVBoxLayout) -> None:
        """Build compression options for image category."""
        compress_group = self._make_option_group(
            _t("conversion_panel.image.compression_options", "Compression Options")
        )
        compress_layout = compress_group.content_layout

        self._compress_btn_group = QButtonGroup(compress_group)
        self._compress_btn_group.setExclusive(True)
        choice_group = ChoiceGroup(compress_group)

        radio_defs: list[tuple[str, str, int]] = [
            (_t("conversion_panel.image.highest_quality", "Lossless"), "lossless", 0),
            (_t("conversion_panel.image.limit_file_size", "Limit Size"), "limit_size", 1),
        ]
        is_limit = self._vm.compress_mode == "limit_size"
        for label, mode_val, btn_id in radio_defs:
            rb = self._make_radio(label, checked=mode_val == self._vm.compress_mode)
            self._compress_btn_group.addButton(rb, btn_id)
            rb.setProperty("compress_value", mode_val)
            choice_group.content_layout.addWidget(rb)

        compress_layout.addWidget(choice_group)

        self._compress_btn_group.idToggled.connect(self._on_compress_mode_changed)

        # Size limit row
        size_control = QWidget(self)
        size_layout = QHBoxLayout(size_control)
        size_layout.setContentsMargins(0, 0, 0, 0)
        size_layout.setSpacing(_SPACING_SM)

        size_edit = QLineEdit(str(self._vm.size_limit), self)
        size_edit.setEnabled(is_limit)
        size_edit.textChanged.connect(self._on_size_input_changed)
        self._size_limit_edit = size_edit
        size_layout.addWidget(size_edit, stretch=1)

        unit_combo = self._make_combo(["KB", "MB"], parent=size_control)
        unit_combo.setCurrentText(self._vm.size_unit)
        unit_combo.setEnabled(is_limit)
        unit_combo.setMinimumWidth(84)
        unit_combo.setSizePolicy(QSizePolicy.Policy.Minimum, QSizePolicy.Policy.Fixed)
        unit_combo.currentTextChanged.connect(self._on_size_input_changed)
        self._size_unit_combo = unit_combo
        size_layout.addWidget(unit_combo)

        size_row = FormRow(
            _t("conversion_panel.image.file_size_limit", "File Size Limit"),
            size_control,
            compress_group,
        )
        compress_layout.addWidget(size_row)

        warning_label = QLabel("", self)
        warning_label.setObjectName("warningLabel")
        self._size_warning_label = warning_label
        compress_layout.addWidget(warning_label)

        parent_layout.addWidget(compress_group)

    def _build_pdf_quality_section(self, parent_layout: QVBoxLayout) -> None:
        """Build PDF quality radio options for image category."""
        quality_group = self._make_option_group(_t("conversion_panel.image.size_options", "Quality Options"))
        quality_layout = quality_group.content_layout

        self._pdf_quality_group = QButtonGroup(quality_group)
        self._pdf_quality_group.setExclusive(True)

        for label, value, btn_id in [
            (_t("conversion_panel.image.original_embed", "Original"), "original", 0),
            (_t("conversion_panel.image.fit_a4", "Fit A4"), "a4", 1),
            (_t("conversion_panel.image.fit_a3", "Fit A3"), "a3", 2),
        ]:
            rb = self._make_radio(label, checked=value == self._vm.pdf_quality)
            self._pdf_quality_group.addButton(rb, btn_id)
            rb.setProperty("quality_value", value)
            quality_layout.addWidget(rb)

        self._pdf_quality_group.buttonClicked.connect(
            lambda btn: setattr(self._vm, "pdf_quality", btn.property("quality_value"))
        )

        parent_layout.addWidget(quality_group)

    def _build_tiff_merge_extra(self) -> None:
        """Build TIFF merge extra section for image category."""
        if self._vm.ui_mode != "batch":
            self._extra_group.hide()
            return
        extra_layout = self._get_extra_content()
        self._extra_group.setTitle(_t("conversion_panel.image.merge_images", "Merge Images"))
        self._extra_group.setVisible(True)

        merge_btn = self._make_action_button(_t("conversion_panel.image.merge_to_tif", "Merge to TIFF"))
        merge_btn.clicked.connect(self._on_merge_tiff_clicked)
        self._merge_tiff_button = merge_btn
        extra_layout.addWidget(merge_btn)

        tiff_group = self._make_option_group(_t("conversion_panel.image.conversion_options", "TIFF Options"))
        tiff_layout = tiff_group.content_layout

        self._tiff_btn_group = QButtonGroup(tiff_group)
        self._tiff_btn_group.setExclusive(True)

        for label, value in [
            (_t("conversion_panel.image.preserve_transparency", "Smart (Preserve Transparency)"), "smart"),
            (_t("conversion_panel.image.no_transparency", "RGB (No Transparency)"), "RGB"),
        ]:
            rb = self._make_radio(label, checked=value == self._vm.tiff_mode)
            rb.setProperty("tiff_value", value)
            self._tiff_btn_group.addButton(rb)
            tiff_layout.addWidget(rb)

        self._tiff_btn_group.buttonClicked.connect(
            lambda btn: setattr(self._vm, "tiff_mode", btn.property("tiff_value"))
        )

        extra_layout.addWidget(tiff_group)

    # ── Layout Section ──────────────────────────────────────────────────

    def _build_layout_section(self) -> None:
        """Build the layout category: conversion row + export/render + merge/split."""
        conv_layout = self._get_primary_content()

        # Conversion row (PDF)
        conversion_formats = self._vm.get_conversion_formats()
        if conversion_formats:
            row_container, combo, btn = self._make_dropdown_action_row(
                conversion_formats, _t("conversion_panel.convert", "Convert")
            )
            self._conversion_combo = combo
            self._conversion_button = btn
            btn.clicked.connect(self._on_convert_button_clicked)
            conv_layout.addWidget(row_container)

        else:
            self._conversion_group.setVisible(False)

        # Save-as section
        saveas_layout = self._get_secondary_content()
        self._build_layout_saveas_section(saveas_layout)

        # Merge/split extra
        if self._vm.supports_layout_pdf_operations:
            self._build_layout_merge_split_extra()

    def _build_layout_saveas_section(self, parent_layout: QVBoxLayout) -> None:
        """Build export + render rows for layout category."""
        # Export row
        export_formats = self._vm.get_layout_export_formats()
        render_formats = self._vm.get_layout_render_formats()
        if not export_formats and not render_formats:
            self._saveas_group.setVisible(False)
            return
        if export_formats:
            row_container, combo, btn = self._make_dropdown_action_row(
                export_formats, _t("conversion_panel.export", "Export")
            )
            self._layout_export_combo = combo
            self._layout_export_button = btn
            btn.clicked.connect(self._on_layout_export_clicked)
            parent_layout.addWidget(row_container)

        if export_formats and render_formats:
            sep = QFrame()
            sep.setFrameShape(QFrame.Shape.HLine)
            parent_layout.addWidget(sep)

        if not render_formats:
            return

        # Render controls need one more field than the common combo/action row.
        # Let selectors stack in narrow viewports; the action always has its own row.
        render_row_container = QWidget(self)
        render_row_container.setObjectName("conversionButtonRow")
        render_row = QVBoxLayout(render_row_container)
        render_row.setContentsMargins(0, 0, 0, 0)
        render_row.setSpacing(_SPACING_SM)
        selectors = ChoiceGroup(render_row_container, responsive=True, spacing=_SPACING_SM)
        render_row.addWidget(selectors)

        render_format_combo = self._make_combo(render_formats, parent=render_row_container)
        render_format_combo.setCurrentText(
            self._vm.render_format if self._vm.render_format in render_formats else render_formats[0]
        )
        render_format_combo.setMinimumWidth(100)
        render_format_combo.currentTextChanged.connect(lambda value: setattr(self._vm, "render_format", value))
        self._layout_render_format_combo = render_format_combo
        selectors.content_layout.addWidget(render_format_combo, 1)

        dpi_combo = self._make_combo(["150", "300", "600"], parent=render_row_container)
        dpi_combo.setCurrentText(str(self._vm.render_dpi))
        dpi_combo.setMinimumWidth(80)
        dpi_combo.currentTextChanged.connect(lambda value: setattr(self._vm, "render_dpi", int(value)))
        self._layout_render_dpi_combo = dpi_combo
        selectors.content_layout.addWidget(dpi_combo, 1)

        render_btn = self._make_action_button(_t("conversion_panel.render", "Render"), parent=render_row_container)
        render_btn.clicked.connect(self._on_layout_render_clicked)
        self._layout_render_button = render_btn
        render_row.addWidget(render_btn)

        parent_layout.addWidget(render_row_container)

    def _build_layout_merge_split_extra(self) -> None:
        """Build merge/split extra section for layout category."""
        extra_layout = self._get_extra_content()
        self._extra_group.setTitle(
            _t("conversion_panel.layout.merge_split", "Merge & Split")
            if self._vm.ui_mode == "batch"
            else _t("conversion_panel.layout.split_to_pdf", "Split PDF")
        )
        self._extra_group.setVisible(True)

        if self._vm.ui_mode == "batch":
            merge_row_container, merge_row = self._make_button_row()

            merge_pdfs_btn = self._make_action_button(
                _t("conversion_panel.layout.merge_to_pdf", "Merge PDFs"), parent=merge_row_container
            )
            merge_pdfs_btn.clicked.connect(self._on_merge_pdfs_clicked)
            self._merge_pdfs_button = merge_pdfs_btn
            merge_row.addWidget(merge_pdfs_btn)

            extra_layout.addWidget(merge_row_container)

            sep = QFrame()
            sep.setFrameShape(QFrame.Shape.HLine)
            extra_layout.addWidget(sep)

        # Split row
        split_row_container, split_row = self._make_button_row()

        split_pdf_btn = self._make_action_button(
            _t("conversion_panel.layout.split_to_pdf", "Split PDF"), parent=split_row_container
        )
        split_pdf_btn.setEnabled(False)
        split_pdf_btn.clicked.connect(self._on_split_pdf_clicked)
        self._split_pdf_button = split_pdf_btn
        split_row.addWidget(split_pdf_btn)

        # Page range input row
        page_edit = QLineEdit(self._vm.page_input, self)
        page_edit.setMinimumWidth(160)
        page_edit.setPlaceholderText(_t("conversion_panel.layout.page_range_placeholder", "e.g., 1-5,7,9-12 or *"))
        page_edit.textChanged.connect(self._on_page_input_changed)
        self._page_input_edit = page_edit
        page_label = _t("conversion_panel.layout.split_page_range", "Page Range")
        page_edit.setAccessibleName(page_label)
        extra_layout.addWidget(FormRow(page_label, page_edit, self))

        # PDF info
        pdf_info_text = _t(
            "conversion_panel.layout.selected_split_file",
            "Selected file: {pages} pages",
        ).format(pages=self._vm.pdf_total_pages)
        pdf_info = WrappingLabel(pdf_info_text, self)
        pdf_info.setObjectName("hintLabel")
        pdf_info.setWordWrap(True)
        self._pdf_info_label = pdf_info
        extra_layout.addWidget(pdf_info)

        page_warning = WrappingLabel("", self)
        page_warning.setObjectName("warningLabel")
        page_warning.setWordWrap(True)
        page_warning.hide()
        self._page_warning_label = page_warning
        extra_layout.addWidget(page_warning)
        extra_layout.addWidget(split_row_container)

    # ── Button Click Handlers ───────────────────────────────────────────

    def _on_convert_button_clicked(self) -> None:
        combo = self._conversion_combo
        self._sync_conversion_button_enabled()
        if combo is None or self._conversion_button is None or not self._conversion_button.isEnabled():
            return
        target = combo.currentText().strip()
        if not target:
            return
        options = self._conversion_options_for_target(target)
        self._vm.request_conversion(target.lower(), options=options)
        if "spreadsheet_password" in options and self._spreadsheet_password_edit is not None:
            self._spreadsheet_password_edit.clear()
            if self._spreadsheet_protection_loss_checkbox is not None:
                self._spreadsheet_protection_loss_checkbox.setChecked(False)

    def _on_saveas_button_clicked(self) -> None:
        combo = self._saveas_combo
        if combo is None:
            return
        target = combo.currentText().strip()
        if not target:
            return
        self._vm.request_conversion(target.lower())

    def _on_validate_clicked(self) -> None:
        self._vm.request_named_action("validate", options=self._vm.validation_options)

    def _on_validation_option_changed(self, key: str, checked: bool) -> None:
        self._vm.set_validation_option(key, checked)
        if self._validate_button:
            self._validate_button.setEnabled(self._vm.is_any_validation_option_checked)

    def _on_merge_tables_clicked(self) -> None:
        self._vm.request_named_action("merge_tables", options={"merge_mode": self._spreadsheet_merge_mode_option()})

    def _on_merge_mode_changed(self, btn_id: int, checked: bool) -> None:
        if checked:
            self._vm.merge_mode = btn_id

    def _on_convert_to_pdf_clicked(self) -> None:
        self._vm.request_conversion("pdf", options={"quality_mode": self._vm.pdf_quality})

    def _on_merge_tiff_clicked(self) -> None:
        self._vm.request_named_action("merge_images_to_tiff", options={"mode": self._vm.tiff_mode})

    def _on_compress_mode_changed(self, btn_id: int, checked: bool) -> None:
        if not checked:
            return
        mode = "limit_size" if btn_id == 1 else "lossless"
        self._vm.compress_mode = mode
        is_limit = mode == "limit_size"
        if self._size_limit_edit:
            self._size_limit_edit.setEnabled(is_limit)
        if self._size_unit_combo:
            self._size_unit_combo.setEnabled(is_limit)
        if not is_limit and self._size_warning_label:
            self._size_warning_label.setText("")

    def _on_size_input_changed(self) -> None:
        if not self._size_limit_edit or not self._size_unit_combo:
            return
        text = self._size_limit_edit.text().strip()
        unit = self._size_unit_combo.currentText()
        if not text:
            if self._size_warning_label:
                self._size_warning_label.setText(
                    _t("conversion_panel.image.enter_size_limit_warning", "Enter a size limit")
                )
            return

        is_valid = self._vm.validate_size_input(text, unit)
        if is_valid:
            if self._size_warning_label:
                self._size_warning_label.setText("")
            self._vm.size_limit = int(text)
            self._vm.size_unit = unit
        else:
            try:
                int(text)
                msg = (
                    _t("conversion_panel.image.kb_range_warning", "KB must be 1-10240")
                    if unit == "KB"
                    else _t("conversion_panel.image.mb_range_warning", "MB must be 1-100")
                )
            except ValueError:
                msg = _t("conversion_panel.image.enter_valid_number_warning", "Enter a valid number")
            if self._size_warning_label:
                self._size_warning_label.setText(msg)

    def _conversion_options_for_target(self, target: str) -> dict[str, object]:
        if self._vm.file_category == "image":
            options: dict[str, object] = {"compress_mode": self._vm.compress_mode}
            if self._vm.compress_mode == "limit_size":
                options["size_limit"] = self._vm.size_limit
                options["size_unit"] = self._vm.size_unit
            return options
        if (
            self._vm.file_category == "spreadsheet"
            and self._vm.current_format == "xlsx"
            and target.strip().lower() == "ods"
            and self._vm.ui_mode == "single"
        ):
            options = {}
            if self._spreadsheet_password_edit is not None:
                password = self._spreadsheet_password_edit.text()
                if password:
                    options["spreadsheet_password"] = password
            if (
                self._spreadsheet_protection_loss_checkbox is not None
                and self._spreadsheet_protection_loss_checkbox.isChecked()
            ):
                options["allow_spreadsheet_protection_loss"] = True
            return options
        return {}

    def _spreadsheet_merge_mode_option(self) -> str:
        return {1: "row", 2: "col", 3: "cell"}.get(self._vm.merge_mode, "cell")

    def _on_layout_export_clicked(self) -> None:
        combo = self._layout_export_combo
        if combo is None:
            return
        self._vm.request_conversion(combo.currentText().lower())

    def _on_layout_render_clicked(self) -> None:
        fmt_combo = self._layout_render_format_combo
        dpi_combo = self._layout_render_dpi_combo
        if fmt_combo is None or dpi_combo is None:
            return
        fmt = fmt_combo.currentText().lower()
        dpi = int(dpi_combo.currentText())
        self._vm.request_conversion(fmt, options={"render_dpi": dpi})

    def _on_merge_pdfs_clicked(self) -> None:
        self._vm.request_named_action("merge_pdfs")

    def _on_split_pdf_clicked(self) -> None:
        if not self._page_input_edit:
            return
        text = self._page_input_edit.text().strip()
        if not text:
            return
        try:
            split_mode, pages = self._vm.parse_split_input(text)
        except ValueError:
            return
        if split_mode != "custom" and self._vm.pdf_total_pages <= 1:
            if self._page_warning_label:
                self._page_warning_label.setText(
                    _t(
                        "conversion_panel.layout.split_mode_single_page_warning",
                        "Split mode not applicable for single-page PDF",
                    )
                )
                self._page_warning_label.show()
            return

        options: dict = {"split_mode": split_mode}
        if split_mode == "custom" and pages:
            options["pages"] = pages
        self._vm.request_named_action("split_pdf", options=options)

    def _on_page_input_changed(self) -> None:
        if not self._page_input_edit:
            return
        raw_text = self._page_input_edit.text()
        if self._vm.page_input != raw_text:
            self._vm.page_input = raw_text
        self._update_page_validation(raw_text.strip())

    def _update_page_validation(self, text: str) -> None:
        """Update split availability and warning text without rebuilding controls."""
        if not text:
            if self._split_pdf_button:
                self._split_pdf_button.setEnabled(False)
            if self._page_warning_label:
                self._page_warning_label.setText("")
                self._page_warning_label.hide()
            return
        is_valid = self._vm.validate_page_input(text, self._vm.pdf_total_pages)
        if self._split_pdf_button:
            self._split_pdf_button.setEnabled(is_valid)
        if text in ("*", "#") and 0 < self._vm.pdf_total_pages <= 1:
            if self._page_warning_label:
                self._page_warning_label.setText(
                    _t(
                        "conversion_panel.layout.split_mode_single_page_warning",
                        "Split mode not applicable for single-page PDF",
                    )
                )
                self._page_warning_label.show()
        else:
            if self._page_warning_label:
                self._page_warning_label.setText("")
                self._page_warning_label.hide()

    # ── Focus Management ─────────────────────────────────────────────────

    def focusInEvent(self, event) -> None:
        """Auto-focus the first enabled interactive child widget."""
        super().focusInEvent(event)
        self.panel_focused.emit()
        self._focus_first_control()

    def _focus_first_control(self) -> None:
        """Set focus to the first enabled combo or interactive widget."""
        try:
            import shiboken6

            if not shiboken6.isValid(self):
                return
        except ImportError:
            pass

        for combo in (self._conversion_combo, self._saveas_combo):
            if combo is not None and combo.isEnabled() and combo.isVisible():
                combo.setFocus(Qt.FocusReason.TabFocusReason)
                return

        for child in self.findChildren(QWidget):
            if isinstance(child, (QPushButton, QCheckBox, QComboBox)) and child.isEnabled() and child.isVisible():
                child.setFocus(Qt.FocusReason.TabFocusReason)
                return

    # ── Public API ───────────────────────────────────────────────────────

    @property
    def view_model(self) -> ConversionPanelViewModel:
        """The ViewModel driving this panel."""
        return self._vm

    @property
    def conversion_combo(self) -> QComboBox | None:
        """The conversion format combo box (for testing)."""
        return self._conversion_combo

    @property
    def conversion_button(self) -> QPushButton | None:
        """The convert button (for testing)."""
        return self._conversion_button

    @property
    def saveas_combo(self) -> QComboBox | None:
        """The save-as format combo box (for testing)."""
        return self._saveas_combo

    @property
    def saveas_button(self) -> QPushButton | None:
        """The save-as button (for testing)."""
        return self._saveas_button

    @property
    def extra_group(self) -> PanelCard:
        """The extra operation card (for testing visibility)."""
        return self._extra_group


__all__ = ["ConversionPanel"]

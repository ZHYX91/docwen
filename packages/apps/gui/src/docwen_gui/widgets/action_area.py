"""ActionArea widget — dynamic action button area with 7 setup modes.

Renders the action area UI based on ActionAreaViewModel state.
Does NOT call runtime/plugins directly — all actions go through the ViewModel.

Supports 7 setup_for_* modes:
  1. setup_for_document_file — Document -> MD (show_numbering=True, show_optimize=True)
  2. setup_for_spreadsheet_file — Spreadsheet -> MD (show_numbering=False, show_optimize=True)
  3. setup_for_image_file — Image -> MD (OCR=True by default)
  4. setup_for_layout_file — Layout -> MD (show_numbering=False, show_optimize=True)
  5. setup_for_other_file — Other -> MD (show_numbering=False, show_optimize=False)
  6. setup_for_md_to_document — MD -> Document (numbering + proofread grid)
  7. setup_for_md_to_spreadsheet — MD -> Spreadsheet (format combo + generate button)

Widget structure:
  - QStackedWidget with 2 pages:
    - page 0: content area (action button + options)
    - page 1: cancel page (cancel button + hint)
  - Initially hidden (setVisible(False))
  - Shows on file selection, hides on clear/cancel
"""

from __future__ import annotations

import logging
from typing import TYPE_CHECKING
from typing import cast as _cast

from PySide6.QtCore import QEvent, QSignalBlocker, Qt, QTimer, Signal
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QSizePolicy,
    QStackedWidget,
    QVBoxLayout,
    QWidget,
)
from qfluentwidgets import CheckBox as FluentCheckBox
from qfluentwidgets import PushButton as FluentPushButton

from docwen_gui import numbering_schemes
from docwen_gui.i18n import t as _t
from docwen_gui.styles.design_tokens import Sizing
from docwen_gui.styles.theme_semantics import apply_theme_class

if TYPE_CHECKING:
    from ..view_models.action_area_vm import ActionAreaViewModel

logger = logging.getLogger(__name__)

# ── Design constants ─────────────────────────────────────────────────────
_SPACING_XS = 4
_SPACING_SM = 8
_SPACING_MD = 12
_SPACING_LG = 16


class _ActionCheckBox(FluentCheckBox):
    """Fluent checkbox that retains DocWen's semantic control height."""

    _METRIC_EVENTS = frozenset(
        {
            QEvent.Type.Polish,
            QEvent.Type.StyleChange,
            QEvent.Type.FontChange,
            QEvent.Type.ApplicationFontChange,
        }
    )

    def event(self, event: QEvent) -> bool:
        handled = super().event(event)
        if event.type() in self._METRIC_EVENTS and self.minimumHeight() < Sizing.CONTROL_HEIGHT:
            # qfluentwidgets' private CheckBox QSS specifies min-height: 22px.
            # Reassert the application token after that stylesheet is polished.
            self.setMinimumHeight(Sizing.CONTROL_HEIGHT)
        return handled


class ActionArea(QWidget):
    """Dynamic action area widget.

    Presents conversion actions and options that change based on the current
    file type.  All user actions are delegated to the ``ActionAreaViewModel``.
    """

    # Signal emitted when the panel receives focus
    panel_focused = Signal()

    def __init__(
        self,
        view_model: ActionAreaViewModel,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self._vm = view_model
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        self.setObjectName("actionAreaRoot")

        # Widget refs — cleared on mode switch
        self._button_stack: QStackedWidget = _cast(QStackedWidget, None)
        self._content_layout: QVBoxLayout = _cast(QVBoxLayout, None)
        self._cancel_button: QPushButton = _cast(QPushButton, None)
        self._cancel_hint_label: QLabel = _cast(QLabel, None)

        # Primary action button refs (for focus/Enter key scanning)
        self.convert_docx_button: QPushButton = _cast(QPushButton, None)
        self.convert_excel_button: QPushButton = _cast(QPushButton, None)
        self.document_to_md_button: QPushButton = _cast(QPushButton, None)
        self.spreadsheet_to_md_button: QPushButton = _cast(QPushButton, None)
        self.image_to_md_button: QPushButton = _cast(QPushButton, None)
        self.layout_to_md_button: QPushButton = _cast(QPushButton, None)

        # Option widget refs
        self._image_cb: QCheckBox = _cast(QCheckBox, None)
        self._ocr_cb: QCheckBox = _cast(QCheckBox, None)
        self._optimize_combo: QComboBox = _cast(QComboBox, None)
        self.doc_remove_numbering_cb: QCheckBox = _cast(QCheckBox, None)
        self.doc_add_numbering_cb: QCheckBox = _cast(QCheckBox, None)
        self.doc_numbering_scheme_combo: QComboBox = _cast(QComboBox, None)
        self.md_remove_numbering_cb: QCheckBox = _cast(QCheckBox, None)
        self.md_add_numbering_cb: QCheckBox = _cast(QCheckBox, None)
        self.md_numbering_scheme_combo: QComboBox = _cast(QComboBox, None)
        self.md_document_format_combo: QComboBox = _cast(QComboBox, None)
        self.md_spreadsheet_format_combo: QComboBox = _cast(QComboBox, None)
        self.checkbox_vars: dict[str, QCheckBox] = {}
        self._content_signature: tuple[object, ...] | None = None
        self._responsive_numbering_rows: list[tuple[QWidget, QGridLayout, QCheckBox, QComboBox]] = []
        self._responsive_numbering_states: dict[int, bool] = {}
        self._proofread_grid_widget: QWidget | None = None
        self._proofread_grid: QGridLayout | None = None
        self._proofread_columns = 0

        self._build_ui()
        self._wire_vm()
        self.setVisible(False)

    def changeEvent(self, event: QEvent) -> None:
        super().changeEvent(event)
        if event.type() != QEvent.Type.PaletteChange:
            return
        from docwen_gui.widgets.conversion_panel import apply_format_swatch_icons

        for combo in self.findChildren(QComboBox):
            apply_format_swatch_icons(combo)

    # ── UI Construction ──────────────────────────────────────────────────

    def _build_ui(self) -> None:
        """Build the action area skeleton."""
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(_SPACING_XS)
        root.setAlignment(Qt.AlignmentFlag.AlignTop)

        # Stacked widget: page 0 = content, page 1 = cancel
        self._button_stack = QStackedWidget(self)
        root.addWidget(self._button_stack)

        # Page 0: content area
        content_page = QFrame()
        content_page.setObjectName("actionContentCard")
        self._content_layout = QVBoxLayout(content_page)
        self._content_layout.setContentsMargins(_SPACING_MD, _SPACING_SM, _SPACING_MD, _SPACING_MD)
        self._content_layout.setSpacing(_SPACING_XS)
        self._button_stack.addWidget(content_page)

        # Page 1: cancel area
        cancel_page = QFrame()
        cancel_page.setObjectName("actionCancelCard")
        cancel_layout = QVBoxLayout(cancel_page)
        cancel_layout.setContentsMargins(_SPACING_LG, _SPACING_LG, _SPACING_LG, _SPACING_LG)

        cancel_text = _t("common.cancel", "Cancel")
        cancel_btn = FluentPushButton(cancel_text, cancel_page)
        cancel_btn.setObjectName("actionCancelButton")
        cancel_btn.setProperty("usesFluentActionButton", True)
        cancel_btn.setProperty("actionButtonRole", "cancel")
        apply_theme_class(cancel_btn, "secondary")
        cancel_btn.setMinimumHeight(Sizing.CONTROL_HEIGHT)
        cancel_btn.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        cancel_btn.setToolTip(cancel_text)
        cancel_btn.setAccessibleName(cancel_text)
        cancel_btn.clicked.connect(self._on_cancel_clicked)
        self._cancel_button = cancel_btn
        cancel_layout.addWidget(cancel_btn)
        self._button_stack.addWidget(cancel_page)

        # Cancel hint label (below stacked widget)
        self._cancel_hint_label = QLabel(_t("components.action_area.cancel_hint", "Operation in progress..."))
        self._cancel_hint_label.setWordWrap(True)
        self._cancel_hint_label.setObjectName("hintLabel")
        self._cancel_hint_label.hide()
        root.addWidget(self._cancel_hint_label)

    # ── ViewModel Wiring ─────────────────────────────────────────────────

    def _wire_vm(self) -> None:
        """Connect ViewModel signals to widget updates."""
        self._vm.state_changed.connect(self._sync_from_vm)
        self._sync_from_vm()

    def _sync_from_vm(self) -> None:
        """Sync widget visibility and cancel stack from ViewModel."""
        self.setVisible(self._vm.visible)

        if self._vm.cancel_visible:
            if self._button_stack:
                self._button_stack.setCurrentIndex(1)
            if self._cancel_button:
                self._cancel_button.setEnabled(True)
            if self._cancel_hint_label:
                self._cancel_hint_label.setVisible(True)
        else:
            if self._button_stack:
                self._button_stack.setCurrentIndex(0)
            if self._cancel_hint_label:
                self._cancel_hint_label.setVisible(False)

        # Rebuild only when the structural mode changes.  Option toggles emit
        # state_changed as well; rebuilding the whole tree on every toggle used
        # to discard focus, popup state and scroll position.
        if self._vm.visible and not self._vm.cancel_visible:
            signature = self._content_structure_signature()
            if signature != self._content_signature:
                self._rebuild_content()
                self._content_signature = signature
            else:
                self._sync_dynamic_controls()
        elif not self._vm.visible:
            self._content_signature = None
        self._enforce_control_metrics()

    def _content_structure_signature(self) -> tuple[object, ...]:
        """Describe only facts that require a different dynamic widget tree."""
        optimization_result = self._vm.optimization_choices_result
        target_result = self._vm.target_route_choices_result
        optimization_items = tuple((choice.id, choice.label) for choice in optimization_result.choices)
        optimization_error = optimization_result.error
        numbering_items: tuple[tuple[str, str], ...] = ()
        if self._vm.show_numbering:
            numbering_items = tuple(self._numbering_scheme_items())
        return (
            self._vm.file_type,
            self._vm.show_numbering,
            self._vm.show_optimize,
            self._vm.show_proofread,
            tuple(self._vm.available_target_formats),
            target_result.status,
            target_result.targets,
            getattr(target_result.error, "code", None),
            optimization_result.status,
            optimization_items,
            getattr(optimization_error, "code", None),
            str(optimization_error) if optimization_error is not None else None,
            numbering_items,
        )

    def _rebuild_content(self) -> None:
        """Clear and rebuild content based on current ViewModel mode."""
        if self._content_layout is None:
            return
        self._clear_content()

        title = _t("action_area.export_options", "Export Options")
        if self._vm.file_type in {"docx", "md_to_spreadsheet"}:
            title = _t("action_area.generation_options", "Generation Options")
        title_label = QLabel(title)
        title_label.setObjectName("actionPanelTitle")
        self._content_layout.addWidget(title_label)

        ft = self._vm.file_type
        if ft in ("document", "spreadsheet", "image", "layout"):
            self._build_file_to_md()
        elif ft == "docx":
            self._build_md_to_document()
        elif ft == "md_to_spreadsheet":
            self._build_md_to_spreadsheet()
        elif ft is not None:
            # Other file -> MD (simplified, no numbering, no optimize)
            self._build_other_to_md()
        self._sync_dynamic_controls()
        QTimer.singleShot(0, self._update_responsive_layouts)

    @staticmethod
    def _set_checkbox_checked(checkbox: QCheckBox | None, checked: bool) -> None:
        if checkbox is None or checkbox.isChecked() == checked:
            return
        blocker = QSignalBlocker(checkbox)
        checkbox.setChecked(checked)
        del blocker

    @staticmethod
    def _set_combo_data(combo: QComboBox | None, value: object) -> None:
        if combo is None:
            return
        index = combo.findData(value)
        if index < 0 or combo.currentIndex() == index:
            return
        blocker = QSignalBlocker(combo)
        combo.setCurrentIndex(index)
        del blocker

    def _sync_dynamic_controls(self) -> None:
        """Update option widgets in place without replacing their QWidget tree."""
        self._set_checkbox_checked(self._image_cb, self._vm.extract_image)
        self._set_checkbox_checked(self._ocr_cb, self._vm.extract_ocr)
        self._set_combo_data(self._optimize_combo, self._vm.optimize_for_type or None)
        self._set_checkbox_checked(self.doc_remove_numbering_cb, self._vm.doc_remove_numbering)
        self._set_checkbox_checked(self.doc_add_numbering_cb, self._vm.doc_add_numbering)
        self._set_combo_data(self.doc_numbering_scheme_combo, self._vm.doc_numbering_scheme)
        if self.doc_numbering_scheme_combo is not None:
            self.doc_numbering_scheme_combo.setEnabled(self._vm.doc_add_numbering)
        self._set_checkbox_checked(self.md_remove_numbering_cb, self._vm.md_remove_numbering)
        self._set_checkbox_checked(self.md_add_numbering_cb, self._vm.md_add_numbering)
        self._set_combo_data(self.md_numbering_scheme_combo, self._vm.md_numbering_scheme)
        self._set_combo_data(self.md_document_format_combo, self._vm.target_format)
        self._set_combo_data(self.md_spreadsheet_format_combo, self._vm.target_format)
        for key, checkbox in self.checkbox_vars.items():
            self._set_checkbox_checked(checkbox, self._vm.proofread_options.get(key, False))
        if self.md_numbering_scheme_combo is not None:
            self.md_numbering_scheme_combo.setEnabled(self._vm.md_add_numbering)

    def _interactive_controls(self) -> tuple[QWidget, ...]:
        """Return the current operation controls without stale rebuilt widgets."""
        candidates: tuple[QWidget | None, ...] = (
            self._cancel_button,
            self.convert_docx_button,
            self.convert_excel_button,
            self.document_to_md_button,
            self.spreadsheet_to_md_button,
            self.image_to_md_button,
            self.layout_to_md_button,
            self.md_document_format_combo,
            self.md_spreadsheet_format_combo,
            self._image_cb,
            self._ocr_cb,
            self._optimize_combo,
            self.doc_remove_numbering_cb,
            self.doc_add_numbering_cb,
            self.doc_numbering_scheme_combo,
            self.md_remove_numbering_cb,
            self.md_add_numbering_cb,
            self.md_numbering_scheme_combo,
            *self.checkbox_vars.values(),
        )
        controls: list[QWidget] = []
        seen: set[int] = set()
        for candidate in candidates:
            if candidate is None or id(candidate) in seen:
                continue
            seen.add(id(candidate))
            controls.append(candidate)
        return tuple(controls)

    def _enforce_control_metrics(self) -> None:
        """Reassert the shared interaction target after Fluent widget polish.

        Some Fluent controls recompute their own 22px minimum while the
        dynamic panel is polished or rebuilt.  Applying the semantic control
        metric after every state projection keeps all modes consistent without
        freezing widths or replacing the widgets.
        """
        for control in self._interactive_controls():
            if control.minimumHeight() < Sizing.CONTROL_HEIGHT:
                control.setMinimumHeight(Sizing.CONTROL_HEIGHT)

    def _clear_content(self) -> None:
        """Remove all dynamic widgets from the content layout."""
        if self._content_layout is None:
            return
        while self._content_layout.count():
            item = self._content_layout.takeAt(0)
            if item is None:
                continue
            w = item.widget()
            if w is not None:
                w.hide()
                w.setParent(None)
                w.deleteLater()

        # Reset all widget refs — pyright: ignore reportAttributeAccessIssue
        # since _cast() declares fields as non-Optional for usage sites
        for _attr_name in (
            "convert_docx_button",
            "convert_excel_button",
            "document_to_md_button",
            "spreadsheet_to_md_button",
            "image_to_md_button",
            "layout_to_md_button",
            "md_document_format_combo",
            "md_spreadsheet_format_combo",
            "_image_cb",
            "_ocr_cb",
            "_optimize_combo",
            "doc_remove_numbering_cb",
            "doc_add_numbering_cb",
            "doc_numbering_scheme_combo",
            "md_remove_numbering_cb",
            "md_add_numbering_cb",
            "md_numbering_scheme_combo",
        ):
            setattr(self, _attr_name, None)  # pyright: ignore[reportAttributeAccessIssue]
        self.checkbox_vars.clear()
        self._responsive_numbering_rows.clear()
        self._responsive_numbering_states.clear()
        self._proofread_grid_widget = None
        self._proofread_grid = None
        self._proofread_columns = 0

    # ── Widget Factory Helpers ───────────────────────────────────────────

    def _make_button(self, text: str, parent: QWidget | None = None) -> QPushButton:
        btn = QPushButton(text, parent or self)
        btn.setObjectName("actionPrimaryButton")
        btn.setMinimumHeight(Sizing.CONTROL_HEIGHT)
        btn.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        return btn

    def _make_checkbox(self, text: str, checked: bool = False, parent: QWidget | None = None) -> QCheckBox:
        cb = _ActionCheckBox(text, parent or self)
        cb.setChecked(checked)
        cb.setMinimumHeight(Sizing.CONTROL_HEIGHT)
        cb.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Fixed)
        return cb

    def _make_combo(self, items: list[str] | None = None, parent: QWidget | None = None) -> QComboBox:
        combo = QComboBox(parent or self)
        combo.setMinimumHeight(Sizing.CONTROL_HEIGHT)
        # Option rows already end in a stretch.  A preferred-width combo keeps
        # the control visually distinct from a text field without claiming the
        # whole row, while still allowing the layout to shrink when necessary.
        combo.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Fixed)
        if items:
            combo.addItems(items)
        return combo

    @staticmethod
    def _make_format_combo_compact(combo: QComboBox) -> None:
        """Size a short format selector from its contents instead of pixels.

        The previous fixed 112 px width was simultaneously too wide for a
        four-letter choice and too narrow once the global font/DPI and arrow
        padding were applied, so ``DOCX`` could be elided.  Qt's content-based
        size hint includes both the longest item and the drop-down affordance.
        """
        combo.setSizeAdjustPolicy(QComboBox.SizeAdjustPolicy.AdjustToContents)
        combo.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)

    def _set_target_format_from_combo(self, combo: QComboBox, index: int) -> None:
        """Project a user-selected target into the ViewModel's canonical state."""
        value = combo.itemData(index)
        if isinstance(value, str) and value:
            self._vm.target_format = value

    def _numbering_scheme_items(self) -> list[tuple[str, str]]:
        return numbering_schemes.get_numbering_scheme_items(
            config_data=self._vm.numbering_scheme_config(),
        )

    def _make_option_row(self, parent: QWidget | None = None) -> tuple[QWidget, QHBoxLayout]:
        row = QWidget(parent or self)
        row.setObjectName("actionOptionRow")
        layout = QHBoxLayout(row)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(_SPACING_XS)
        return row, layout

    def _make_responsive_option_grid(
        self,
        parent: QWidget | None = None,
    ) -> tuple[QWidget, QGridLayout]:
        """Create an option row that can stack its controls when space is tight."""

        row = QWidget(parent or self)
        row.setObjectName("actionOptionRow")
        layout = QGridLayout(row)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setHorizontalSpacing(_SPACING_XS)
        layout.setVerticalSpacing(_SPACING_XS)
        return row, layout

    def resizeEvent(self, event) -> None:
        """Keep compact option groups readable across panel widths and locales."""

        super().resizeEvent(event)
        self._update_responsive_layouts()

    def _update_responsive_layouts(self) -> None:
        """Use compact rows when they fit and lossless stacked rows otherwise."""

        for row, layout, checkbox, combo in self._responsive_numbering_rows:
            available = row.contentsRect().width()
            if available <= 0:
                available = max(0, self.contentsRect().width() - (2 * _SPACING_MD))
            required = checkbox.sizeHint().width() + combo.sizeHint().width() + layout.horizontalSpacing()
            stacked = available > 0 and required > available
            state_key = id(row)
            if self._responsive_numbering_states.get(state_key) is stacked:
                continue

            layout.removeWidget(checkbox)
            layout.removeWidget(combo)
            if stacked:
                layout.addWidget(checkbox, 0, 0)
                layout.addWidget(combo, 1, 0)
                layout.setColumnStretch(0, 0)
                layout.setColumnStretch(1, 1)
                layout.setColumnStretch(2, 0)
            else:
                layout.addWidget(checkbox, 0, 0)
                layout.addWidget(combo, 0, 1)
                layout.setColumnStretch(0, 0)
                layout.setColumnStretch(1, 0)
                layout.setColumnStretch(2, 1)
            self._responsive_numbering_states[state_key] = stacked

        grid = self._proofread_grid
        grid_widget = self._proofread_grid_widget
        checkboxes = list(self.checkbox_vars.values())
        if grid is None or grid_widget is None or not checkboxes:
            return

        widest_checkbox = max(checkbox.sizeHint().width() for checkbox in checkboxes)
        required_two_columns = (2 * widest_checkbox) + grid.horizontalSpacing()
        # The proofreading choices are a stable compact 2x2 group. Preserve
        # that shape for every locale and give the
        # group a truthful minimum width instead of silently clipping labels or
        # turning the compact group into four rows.
        grid_widget.setMinimumWidth(required_two_columns)
        if self._proofread_columns == 2:
            return

        for checkbox in checkboxes:
            grid.removeWidget(checkbox)
        for index, checkbox in enumerate(checkboxes):
            row, column = divmod(index, 2)
            grid.addWidget(checkbox, row, column)
        grid.setColumnStretch(0, 1)
        grid.setColumnStretch(1, 1)
        self._proofread_columns = 2

    # ── File -> MD Layout ───────────────────────────────────────────────

    def _build_file_to_md(self) -> None:
        """Build the file->MD layout: button + options (extract_image, OCR, optimize, numbering)."""
        if self._content_layout is None:
            return

        ft = self._vm.file_type or "document"

        # Main action button row
        button_row, button_layout = self._make_option_row()
        button_layout.addStretch(1)

        btn = self._make_button(self._vm.get_button_label(), parent=button_row)
        btn.setToolTip(self._vm.get_button_tooltip())
        btn.clicked.connect(lambda: self._on_file_to_md_clicked("md"))

        # Assign to the correct attr_name slot
        attr_map = {
            "document": "document_to_md_button",
            "spreadsheet": "spreadsheet_to_md_button",
            "image": "image_to_md_button",
            "layout": "layout_to_md_button",
        }
        attr_name = attr_map.get(ft, "document_to_md_button")
        setattr(self, attr_name, btn)

        button_layout.addWidget(btn)
        button_layout.addStretch(1)
        self._content_layout.addWidget(button_row)

        # Options section
        self._build_file_to_md_options()

    def _build_file_to_md_options(self) -> None:
        """Build extract_image, OCR, optimize, and numbering options rows."""
        if self._content_layout is None:
            return

        # The two primary boolean options belong to one compact row, matching
        # the old panel's scan pattern while remaining keyboard accessible.
        img_row, img_layout = self._make_option_row()
        self._image_cb = self._make_checkbox(
            _t("action_area.extract_images", "Extract Images"),
            checked=self._vm.extract_image,
            parent=img_row,
        )
        self._image_cb.stateChanged.connect(lambda state: self._vm.set_file_to_md_option("extract_image", bool(state)))
        img_layout.addWidget(self._image_cb)
        self._ocr_cb = self._make_checkbox(
            _t("action_area.ocr", "Enable OCR"),
            checked=self._vm.extract_ocr,
            parent=img_row,
        )
        self._ocr_cb.stateChanged.connect(lambda state: self._vm.set_file_to_md_option("extract_ocr", bool(state)))
        img_layout.addWidget(self._ocr_cb)
        img_layout.addStretch(1)
        self._content_layout.addWidget(img_row)

        # Optimize option (if shown)
        if self._vm.show_optimize:
            self._build_optimize_row()

        # Numbering options (if shown)
        if self._vm.show_numbering:
            self._build_file_to_md_numbering_rows()

    def _build_optimize_row(self) -> None:
        """Build the optimize-for-type combo row."""
        if self._content_layout is None:
            return
        opt_row, opt_layout = self._make_option_row()
        optimization_result = self._vm.optimization_choices_result
        if optimization_result.status == "failed":
            notice = QLabel(
                _t(
                    "action_area.optimization_unavailable",
                    "Optimization options are unavailable; standard conversion remains available.",
                ),
                opt_row,
            )
            notice.setObjectName("actionOptimizationUnavailable")
            notice.setProperty("settingsRole", "tabDescription")
            notice.setTextFormat(Qt.TextFormat.PlainText)
            notice.setWordWrap(True)
            error = optimization_result.error
            if error is not None:
                notice.setProperty("errorCode", getattr(error, "code", "capability_unavailable"))
                notice.setToolTip(str(getattr(error, "code", "capability_unavailable")))
            opt_layout.addWidget(notice)
            opt_layout.addStretch(1)
            self._content_layout.addWidget(opt_row)
            return

        ft = self._vm.file_type or "document"

        label_key_map = {
            "document": "action_area.document.optimize_for_type",
            "spreadsheet": "action_area.spreadsheet.optimize_for_type",
            "image": "action_area.image.optimize_for_type",
            "layout": "action_area.layout.optimize_for_type",
        }
        label_text = _t(label_key_map.get(ft, "action_area.layout.optimize_for_type"), "Optimize For")

        opt_label = QLabel(label_text, opt_row)
        opt_layout.addWidget(opt_label)

        optimization_choices = optimization_result.choices
        if not optimization_choices:
            opt_row.deleteLater()
            return

        combo = self._make_combo(parent=opt_row)
        combo.setAccessibleName(label_text)
        opt_label.setBuddy(combo)
        combo.addItem(_t("common.default", "Default"), None)
        for choice in optimization_choices:
            combo.addItem(choice.label, choice.id)

        selected_idx = combo.findData(self._vm.optimize_for_type)
        combo.setCurrentIndex(selected_idx if selected_idx >= 0 else 0)
        combo.currentIndexChanged.connect(
            lambda idx: self._vm.set_file_to_md_option("optimize_for_type", combo.itemData(idx))
        )
        self._optimize_combo = combo
        opt_layout.addWidget(combo)
        opt_layout.addStretch(1)
        self._content_layout.addWidget(opt_row)

    def _build_file_to_md_numbering_rows(self) -> None:
        """Build remove/add numbering option rows for document->MD."""
        if self._content_layout is None:
            return

        # Remove numbering
        rem_row, rem_layout = self._make_option_row()
        self.doc_remove_numbering_cb = self._make_checkbox(
            _t("action_area.document.remove_existing_numbering", "Remove Existing Numbering"),
            checked=self._vm.doc_remove_numbering,
            parent=rem_row,
        )
        self.doc_remove_numbering_cb.stateChanged.connect(
            lambda state: self._vm.set_file_to_md_option("remove_numbering", bool(state))
        )
        rem_layout.addWidget(self.doc_remove_numbering_cb)
        rem_layout.addStretch(1)
        self._content_layout.addWidget(rem_row)

        # Add numbering
        add_row, add_layout = self._make_responsive_option_grid()
        self.doc_add_numbering_cb = self._make_checkbox(
            _t("action_area.document.add_new_numbering", "Add New Numbering"),
            checked=self._vm.doc_add_numbering,
            parent=add_row,
        )
        self.doc_add_numbering_cb.stateChanged.connect(
            lambda state: self._vm.set_file_to_md_option("add_numbering", bool(state))
        )
        add_layout.addWidget(self.doc_add_numbering_cb, 0, 0)

        # Numbering scheme combo
        scheme_combo = self._make_combo(parent=add_row)
        scheme_combo.setAccessibleName(
            _t("settings.document.numbering.scheme_label", "Default numbering scheme:").rstrip(":：")
        )
        for label, scheme_id in self._numbering_scheme_items():
            scheme_combo.addItem(label, scheme_id)
        idx = scheme_combo.findData(self._vm.doc_numbering_scheme)
        if idx >= 0:
            scheme_combo.setCurrentIndex(idx)
        scheme_combo.setEnabled(self._vm.doc_add_numbering)
        scheme_combo.currentIndexChanged.connect(
            lambda idx: self._vm.set_file_to_md_option(
                "numbering_scheme",
                scheme_combo.itemData(idx),
            )
        )
        self.doc_numbering_scheme_combo = scheme_combo
        self.doc_add_numbering_cb.stateChanged.connect(lambda state: scheme_combo.setEnabled(bool(state)))
        add_layout.addWidget(scheme_combo, 0, 1)
        add_layout.setColumnStretch(2, 1)
        self._responsive_numbering_rows.append((add_row, add_layout, self.doc_add_numbering_cb, scheme_combo))
        self._content_layout.addWidget(add_row)

    # ── Other -> MD Layout ───────────────────────────────────────────────

    def _build_other_to_md(self) -> None:
        """Simplified file->MD layout for 'other' category (no numbering, no optimize)."""
        if self._content_layout is None:
            return

        # Action button
        button_row, button_layout = self._make_option_row()
        button_layout.addStretch(1)

        btn = self._make_button(self._vm.get_button_label(), parent=button_row)
        btn.setToolTip(self._vm.get_button_tooltip())
        btn.clicked.connect(lambda: self._on_file_to_md_clicked("md"))
        self.document_to_md_button = btn

        button_layout.addWidget(btn)
        button_layout.addStretch(1)
        self._content_layout.addWidget(button_row)

        # Keep the two universal boolean options on one compact row.
        img_row, img_layout = self._make_option_row()
        self._image_cb = self._make_checkbox(
            _t("action_area.extract_images", "Extract Images"),
            checked=self._vm.extract_image,
            parent=img_row,
        )
        self._image_cb.stateChanged.connect(lambda state: self._vm.set_file_to_md_option("extract_image", bool(state)))
        img_layout.addWidget(self._image_cb)
        self._ocr_cb = self._make_checkbox(
            _t("action_area.ocr", "Enable OCR"),
            checked=self._vm.extract_ocr,
            parent=img_row,
        )
        self._ocr_cb.stateChanged.connect(lambda state: self._vm.set_file_to_md_option("extract_ocr", bool(state)))
        img_layout.addWidget(self._ocr_cb)
        img_layout.addStretch(1)
        self._content_layout.addWidget(img_row)

        if self._vm.show_optimize:
            self._build_optimize_row()

    # ── MD -> Document Layout ───────────────────────────────────────────

    def _build_md_to_document(self) -> None:
        """Build the MD->Document layout: format combo + generate button + numbering + proofread."""
        if self._content_layout is None:
            return

        # Generate row
        gen_row, gen_layout = self._make_option_row()
        gen_layout.addStretch(1)

        format_combo = self._make_combo(parent=gen_row)
        for label in self._vm.available_target_formats:
            format_combo.addItem(label, label.lower())
        from docwen_gui.widgets.conversion_panel import apply_format_swatch_icons

        apply_format_swatch_icons(format_combo)
        idx = format_combo.findData(self._vm.target_format)
        if idx >= 0:
            format_combo.setCurrentIndex(idx)
        self._make_format_combo_compact(format_combo)
        route_ready = self._vm.target_route_choices_result.status == "ready"
        format_combo.setEnabled(route_ready)
        format_combo.setAccessibleName(
            _t("action_area.md_to_document.format_combo_tooltip", "Choose target document format")
        )
        format_combo.setToolTip(_t("action_area.md_to_document.format_combo_tooltip", "Choose target document format"))
        format_combo.currentIndexChanged.connect(lambda index: self._set_target_format_from_combo(format_combo, index))
        self.md_document_format_combo = format_combo
        gen_layout.addWidget(format_combo)

        generate_btn = self._make_button(_t("action_area.generate", "Generate"), parent=gen_row)
        generate_btn.setEnabled(route_ready)
        generate_btn.clicked.connect(lambda: self._on_md_to_document_clicked(format_combo.currentData()))
        self.convert_docx_button = generate_btn

        gen_layout.addWidget(generate_btn)
        gen_layout.addStretch(1)
        self._content_layout.addWidget(gen_row)

        self._add_target_route_notice()

        # Numbering rows
        self._build_md_numbering_rows()

        # Divider
        sep = QFrame()
        sep.setObjectName("actionSectionDividerLine")
        sep.setProperty("actionDividerVariant", "subtle")
        sep.setFrameShape(QFrame.Shape.HLine)
        self._content_layout.addWidget(sep)

        # Proofread grid
        self._build_proofread_grid()

    def _build_md_numbering_rows(self) -> None:
        """Build numbering options for MD->Document mode.

        Layout:
          1. Remove Existing Numbering
          2. Add New Numbering  +  [Numbering Scheme]

        The infrequently changed text/Word-native rendering preference lives
        in Text settings.  The ViewModel still reads and submits that persisted
        value; duplicating it here made the per-document panel noisy and let a
        transient choice disagree with Settings.
        """
        if self._content_layout is None:
            return

        # Row 1: Remove numbering
        rem_row, rem_layout = self._make_option_row()
        self.md_remove_numbering_cb = self._make_checkbox(
            _t("action_area.md_to_document.remove_existing_numbering", "Remove Existing Numbering"),
            checked=self._vm.md_remove_numbering,
            parent=rem_row,
        )
        self.md_remove_numbering_cb.stateChanged.connect(
            lambda state: self._vm.set_md_to_doc_option("remove_numbering", bool(state))
        )
        rem_layout.addWidget(self.md_remove_numbering_cb)
        rem_layout.addStretch(1)
        self._content_layout.addWidget(rem_row)

        # Row 2: Add numbering checkbox + scheme combo
        add_row, add_layout = self._make_responsive_option_grid()
        self.md_add_numbering_cb = self._make_checkbox(
            _t("action_area.md_to_document.add_new_numbering", "Add New Numbering"),
            checked=self._vm.md_add_numbering,
            parent=add_row,
        )
        self.md_add_numbering_cb.stateChanged.connect(
            lambda state: self._vm.set_md_to_doc_option("add_numbering", bool(state))
        )
        add_layout.addWidget(self.md_add_numbering_cb, 0, 0)

        scheme_combo = self._make_combo(parent=add_row)
        scheme_combo.setAccessibleName(_t("settings.text.scheme_label", "Default numbering scheme:").rstrip(":："))
        for label, scheme_id in self._numbering_scheme_items():
            scheme_combo.addItem(label, scheme_id)
        idx = scheme_combo.findData(self._vm.md_numbering_scheme)
        if idx >= 0:
            scheme_combo.setCurrentIndex(idx)
        scheme_combo.setEnabled(self._vm.md_add_numbering)
        scheme_combo.currentIndexChanged.connect(
            lambda idx: self._vm.set_md_to_doc_option(
                "numbering_scheme",
                scheme_combo.itemData(idx),
            )
        )
        self.md_numbering_scheme_combo = scheme_combo
        self.md_add_numbering_cb.stateChanged.connect(lambda state: scheme_combo.setEnabled(bool(state)))
        add_layout.addWidget(scheme_combo, 0, 1)
        add_layout.setColumnStretch(2, 1)
        self._responsive_numbering_rows.append((add_row, add_layout, self.md_add_numbering_cb, scheme_combo))
        self._content_layout.addWidget(add_row)

    def _build_proofread_grid(self) -> None:
        """Build proofread checkboxes for MD->Document mode."""
        if self._content_layout is None:
            return

        proofread_defs: list[tuple[str, str]] = [
            ("symbol_pairing", _t("conversion_panel.document.symbol_pairing", "Symbol Pairing")),
            ("typos_rule", _t("conversion_panel.document.typos_rule", "Typos Rule")),
            ("symbol_correction", _t("conversion_panel.document.symbol_correction", "Symbol Correction")),
            ("sensitive_word", _t("conversion_panel.document.sensitive_word", "Sensitive Word")),
        ]

        grid_widget = QWidget(self)
        grid = QGridLayout(grid_widget)
        grid.setContentsMargins(0, 0, 0, 0)
        grid.setHorizontalSpacing(_SPACING_MD)
        grid.setVerticalSpacing(_SPACING_XS)
        grid.setColumnStretch(0, 1)
        grid.setColumnStretch(1, 1)
        self._proofread_grid_widget = grid_widget
        self._proofread_grid = grid
        self._proofread_columns = 2

        opts = self._vm.proofread_options
        self.checkbox_vars.clear()
        for i, (key, label) in enumerate(proofread_defs):
            cb = self._make_checkbox(label, checked=opts.get(key, False), parent=grid_widget)
            cb.stateChanged.connect(lambda state, k=key: self._vm.set_proofread_option(k, bool(state)))
            self.checkbox_vars[key] = cb
            row, column = divmod(i, 2)
            grid.addWidget(cb, row, column)

        self._content_layout.addWidget(grid_widget)

    # ── MD -> Spreadsheet Layout ────────────────────────────────────────

    def _build_md_to_spreadsheet(self) -> None:
        """Build the MD->Spreadsheet layout: format combo + generate button."""
        if self._content_layout is None:
            return

        gen_row, gen_layout = self._make_option_row()
        gen_layout.addStretch(1)

        format_combo = self._make_combo(parent=gen_row)
        for label in self._vm.available_target_formats:
            format_combo.addItem(label, label.lower())
        from docwen_gui.widgets.conversion_panel import apply_format_swatch_icons

        apply_format_swatch_icons(format_combo)
        idx = format_combo.findData(self._vm.target_format)
        if idx >= 0:
            format_combo.setCurrentIndex(idx)
        self._make_format_combo_compact(format_combo)
        route_ready = self._vm.target_route_choices_result.status == "ready"
        format_combo.setEnabled(route_ready)
        format_combo.setAccessibleName(
            _t("action_area.md_to_spreadsheet.format_combo_tooltip", "Choose target spreadsheet format")
        )
        format_combo.setToolTip(
            _t("action_area.md_to_spreadsheet.format_combo_tooltip", "Choose target spreadsheet format")
        )
        format_combo.currentIndexChanged.connect(lambda index: self._set_target_format_from_combo(format_combo, index))
        self.md_spreadsheet_format_combo = format_combo
        gen_layout.addWidget(format_combo)

        generate_btn = self._make_button(_t("action_area.generate", "Generate"), parent=gen_row)
        generate_btn.setEnabled(route_ready)
        generate_btn.clicked.connect(lambda: self._on_md_to_spreadsheet_clicked(format_combo.currentData()))
        self.convert_excel_button = generate_btn

        gen_layout.addWidget(generate_btn)
        gen_layout.addStretch(1)
        self._content_layout.addWidget(gen_row)

        self._add_target_route_notice()

    def _add_target_route_notice(self) -> None:
        """Explain why generation is disabled for failed or valid-empty catalogs."""

        if self._content_layout is None:
            return
        result = self._vm.target_route_choices_result
        if result.status == "ready":
            return
        if result.status == "failed":
            message = _t(
                "action_area.route_options_unavailable",
                "Available operations could not be loaded; generation is disabled to avoid an incorrect request.",
            )
        else:
            message = _t(
                "action_area.no_compatible_output",
                "No compatible output is available for this file.",
            )
        label = QLabel(message, self)
        label.setObjectName("routeStateNotice")
        label.setProperty("routeState", result.status)
        label.setWordWrap(True)
        self._content_layout.addWidget(label)

    # ── Click Handlers ──────────────────────────────────────────────────

    def _on_file_to_md_clicked(self, target_format: str) -> None:
        """Handle file->MD conversion button click."""
        if not self._vm.file_path:
            return
        self._vm.request_conversion(target_format)

    def _on_md_to_document_clicked(self, target_format: str) -> None:
        """Handle MD->Document conversion button click."""
        if not self._vm.file_path:
            return
        self._vm.save_last_document_format(target_format)
        self._vm.request_conversion(target_format)

    def _on_md_to_spreadsheet_clicked(self, target_format: str) -> None:
        """Handle MD->Spreadsheet conversion button click."""
        if not self._vm.file_path:
            return
        self._vm.save_last_spreadsheet_format(target_format)
        self._vm.request_conversion(target_format)

    def _on_cancel_clicked(self) -> None:
        """Handle cancel button click."""
        if self._cancel_button:
            self._cancel_button.setEnabled(False)
        self._vm.request_cancel()

    # ── Focus Management ─────────────────────────────────────────────────

    def focusInEvent(self, event) -> None:
        """Auto-focus the first enabled interactive child widget."""
        super().focusInEvent(event)
        self.panel_focused.emit()
        self._focus_first_control()

    def _focus_first_control(self) -> None:
        """Set focus to the first enabled QPushButton/QCheckBox/QComboBox."""
        for child in self.findChildren(QWidget):
            if isinstance(child, (QPushButton, QCheckBox, QComboBox)) and child.isEnabled() and child.isVisible():
                child.setFocus(Qt.FocusReason.TabFocusReason)
                return

    def trigger_primary_action(self) -> bool:
        """Scan primary buttons and click the first visible+enabled one.

        Used by the Enter key shortcut handler. Only works when the
        button_stack is on page 0 (content area).

        Returns:
            True if a button was triggered.
        """
        if self._button_stack and self._button_stack.currentIndex() != 0:
            return False

        buttons = [
            self.convert_docx_button,
            self.convert_excel_button,
            self.document_to_md_button,
            self.spreadsheet_to_md_button,
            self.image_to_md_button,
            self.layout_to_md_button,
        ]
        for btn in buttons:
            if btn is None or not btn.isVisibleTo(self) or not btn.isEnabled():
                continue
            btn.click()
            return True
        return False

    # ── Public API ───────────────────────────────────────────────────────

    @property
    def view_model(self) -> ActionAreaViewModel:
        """The ViewModel driving this action area."""
        return self._vm

    @property
    def cancel_button(self) -> QPushButton | None:
        """The cancel button (for external access / testing)."""
        return self._cancel_button

    @property
    def button_stack(self) -> QStackedWidget | None:
        """The stacked widget (for testing)."""
        return self._button_stack

    def show_panel(self) -> None:
        """Show the action area."""
        self._vm.show()

    def hide_panel(self) -> None:
        """Hide the action area."""
        self._vm.hide()

    def show_cancel(self) -> None:
        """Switch to the cancel page."""
        self._vm.show_cancel()

    def hide_cancel(self) -> None:
        """Switch back to the action page."""
        self._vm.hide_cancel()


__all__ = ["ActionArea"]

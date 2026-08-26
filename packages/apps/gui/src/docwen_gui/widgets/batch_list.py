"""BatchList widget — 6-tab batch file list with status icons, filtering, sorting.

Renders the batch file list UI based on BatchListViewModel state.
Does NOT call runtime/plugins directly — all actions go through the ViewModel.

Widget structure:
  - Summary bar: filter button, summary label, move up/down, sort button
  - FluentPivot tabs (6 categories)
  - QStackedWidget with one ReorderableListWidget per category
  - BatchEntryItemWidget cards for each file entry
  - Right-click context menu on entries
  - Ctrl+Up/Down keyboard reordering
"""

from __future__ import annotations

import contextlib
import logging
from collections import deque
from pathlib import Path
from typing import TYPE_CHECKING
from typing import cast as _cast

from PySide6.QtCore import (
    QEvent,
    QPoint,
    QRect,
    QSize,
    Qt,
    QTimer,
    QVariantAnimation,
    Signal,
)
from PySide6.QtGui import (
    QActionGroup,
    QIcon,
    QKeyEvent,
)
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QBoxLayout,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLayout,
    QListWidget,
    QListWidgetItem,
    QMenu,
    QPushButton,
    QSizePolicy,
    QStackedWidget,
    QVBoxLayout,
    QWidget,
)

from docwen_gui.font_utils import DEFAULT_FONT_SIZE, resolve_font_size_preset
from docwen_gui.i18n import t as _t
from docwen_gui.styles.theme_semantics import apply_theme_class

if TYPE_CHECKING:
    from ..view_models.batch_list_vm import BatchFileEntry, BatchListViewModel

logger = logging.getLogger(__name__)

# ── Design constants ───────────────────────────────────────────────────

_SPACING_XS = 4
_SPACING_SM = 8
_SPACING_MD = 12
_SPACING_LG = 16

_BATCH_ENTRY_COMPACT_WIDTH_THRESHOLD = 340
_BATCH_CATEGORY_PIVOT_NARROW_THRESHOLD = 380
_BATCH_STATUS_PULSE_ENTRY_LIMIT = 40

# Constructing a card creates a substantial QWidget/layout tree.  Keep one
# event-loop slice bounded so a large Explorer drop cannot monopolize the GUI
# thread while every card is materialized.
_BATCH_ENTRY_WIDGET_ATTACH_CHUNK_SIZE = 8

_CATEGORY_ORDER = ["text", "spreadsheet", "document", "image", "layout", "other"]

# ── SVG icon helper ────────────────────────────────────────────────────


def _load_status_icon(status: str) -> QIcon:
    """Load a status icon (16x16) recolored with the theme-aware semantic color."""
    asset_names = {
        "processing": "sync.svg",
        "completed": "complete.svg",
        "skipped": "skip.svg",
        "failed": "error.svg",
        "cancelled": "skip.svg",
    }
    if status == "pending":
        return QIcon()
    asset = asset_names.get(status)
    if asset is None:
        return QIcon()
    from docwen_gui.resources import load_svg_icon
    from docwen_gui.styles.theme_manager import ThemeManager
    from docwen_gui.styles.theme_semantics import get_status_color

    theme_name = ThemeManager.get_instance().get_current_theme()
    icon = load_svg_icon(asset, color=get_status_color(status, theme_name))
    return icon if icon is not None else QIcon()


def _soft_wrap_filename(file_name: str) -> str:
    """Add invisible wrap opportunities after common filename separators."""
    return "".join(f"{char}\u200b" if char in {"-", "_"} else char for char in file_name)


def _get_status_color(status: str) -> str:
    """Return a semantic color name for a status."""
    mapping = {
        "pending": "gray",
        "processing": "blue",
        "completed": "green",
        "skipped": "orange",
        "failed": "red",
        "cancelled": "orange",
    }
    return mapping.get(status, "gray")


def _filter_option_label(filter_key: str, fallback: str) -> str:
    """Return the localized display label for a status filter."""
    if filter_key == "all":
        return _t("components.file_drop.batch_list.filter_all", fallback)
    status_fallback = _t(f"components.file_drop.status.{filter_key}", fallback)
    return _t(f"components.file_drop.batch_list.filter_{filter_key}", status_fallback)


def _sort_option_label(sort_key: str) -> str:
    """Return the localized display label for a sort option."""
    labels = {
        "custom": _t("components.file_drop.batch_list.sort_custom", "Custom order"),
        "name": _t("components.file_drop.batch_list.sort_name", "File name"),
        "type": _t("components.file_drop.batch_list.sort_type", "Type"),
        "size": _t("components.file_drop.batch_list.sort_size", "Size"),
        "mtime": _t("components.file_drop.batch_list.sort_mtime", "Modified time"),
    }
    return labels.get(sort_key, labels["custom"])


# ── ReorderableListWidget ──────────────────────────────────────────────


class ReorderableListWidget(QListWidget):
    """QListWidget subclass supporting Ctrl+Up/Down reordering."""

    item_reordered = Signal()

    def keyPressEvent(self, event: QKeyEvent) -> None:
        if self._should_reorder_from_event(event, Qt.Key.Key_Up) and self.move_current_item_by(-1):
            event.accept()
            return
        if self._should_reorder_from_event(event, Qt.Key.Key_Down) and self.move_current_item_by(1):
            event.accept()
            return
        super().keyPressEvent(event)

    @staticmethod
    def _should_reorder_from_event(event: QKeyEvent, key: Qt.Key) -> bool:
        modifiers = event.modifiers()
        return event.key() == key and bool(modifiers & Qt.KeyboardModifier.ControlModifier)

    def move_current_item_by(self, offset: int) -> bool:
        current_item = self.currentItem()
        if current_item is None or current_item.isHidden():
            return False
        visible_items = self._visible_items()
        try:
            current_index = visible_items.index(current_item)
        except ValueError:
            return False

        target_index = current_index + offset
        if target_index < 0 or target_index >= len(visible_items):
            return False

        current_row = self.row(current_item)
        target_row = self.row(visible_items[target_index])
        item_widget = self.itemWidget(current_item)
        if item_widget is not None:
            self.removeItemWidget(current_item)
            item_widget.setParent(None)
        item = self.takeItem(current_row)
        self.insertItem(target_row, item)
        if item_widget is not None:
            self.setItemWidget(item, item_widget)
        self.setCurrentItem(item)
        self.item_reordered.emit()
        return True

    def _visible_items(self) -> list[QListWidgetItem]:
        return [self.item(i) for i in range(self.count()) if not self.item(i).isHidden()]


class _InteractivePathLabel(QLabel):
    """Keyboard-accessible file label with open and copy-path affordances."""

    activated = Signal()

    def __init__(self, text: str, file_path: str, parent: QWidget | None = None) -> None:
        super().__init__(text, parent)
        self._file_path = file_path
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        self.setCursor(Qt.CursorShape.PointingHandCursor)

    def set_file_path(self, file_path: str) -> None:
        self._file_path = file_path

    def mouseReleaseEvent(self, event) -> None:
        if event.button() == Qt.MouseButton.LeftButton and self.rect().contains(event.position().toPoint()):
            self.activated.emit()
            event.accept()
            return
        super().mouseReleaseEvent(event)

    def keyPressEvent(self, event: QKeyEvent) -> None:
        if event.key() in {Qt.Key.Key_Return, Qt.Key.Key_Enter, Qt.Key.Key_Space}:
            self.activated.emit()
            event.accept()
            return
        super().keyPressEvent(event)

    def contextMenuEvent(self, event) -> None:
        menu = self._create_context_menu()
        menu.exec(event.globalPos())
        event.accept()

    def _create_context_menu(self) -> QMenu:
        menu = QMenu(self)
        copy_action = menu.addAction(_t("info_area.copy_path", "Copy Path"))
        copy_action.triggered.connect(self._copy_path_to_clipboard)
        return menu

    def _copy_path_to_clipboard(self) -> None:
        QApplication.clipboard().setText(self._file_path)


class _MiddleElidedLabel(QLabel):
    """Single-line label that preserves a full path in its tooltip.

    Middle elision keeps both the path root and filename recognizable without
    allowing a long absolute path to widen the whole left panel.
    """

    def __init__(self, text: str = "", parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._full_text = ""
        self.setWordWrap(False)
        self.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Fixed)
        self.set_full_text(text)

    @property
    def full_text(self) -> str:
        return self._full_text

    def set_full_text(self, text: str) -> None:
        self._full_text = text
        self.setToolTip(text)
        self.setAccessibleName(text)
        self._refresh_elision()

    def resizeEvent(self, event) -> None:
        super().resizeEvent(event)
        self._refresh_elision()

    def _refresh_elision(self) -> None:
        available_width = self.contentsRect().width()
        if available_width <= 0:
            super().setText(self._full_text)
            return
        super().setText(
            self.fontMetrics().elidedText(
                self._full_text,
                Qt.TextElideMode.ElideMiddle,
                available_width,
            )
        )


# ── WrapRowLayout ──────────────────────────────────────────────────────


class WrapRowLayout(QLayout):
    """Lightweight flow layout for badge wrapping in narrow mode."""

    def __init__(self, parent: QWidget | None = None, *, spacing: int = _SPACING_XS):
        super().__init__(parent)
        self._items: list[QWidget | QLayout] = []
        self.setContentsMargins(0, 0, 0, 0)
        self.setSpacing(spacing)

    def addItem(self, item) -> None:
        self._items.append(item)

    def count(self) -> int:
        return len(self._items)

    def itemAt(self, index: int):  # pyright: ignore[reportIncompatibleMethodOverride]
        if 0 <= index < len(self._items):
            return self._items[index]
        return None

    def takeAt(self, index: int):  # pyright: ignore[reportIncompatibleMethodOverride]
        if 0 <= index < len(self._items):
            return self._items.pop(index)
        return None

    def expandingDirections(self) -> Qt.Orientation:
        return Qt.Orientation(0)

    def hasHeightForWidth(self) -> bool:
        return True

    def heightForWidth(self, width: int) -> int:
        return self._do_layout(QRect(0, 0, max(0, width), 0), test_only=True)

    def setGeometry(self, rect: QRect) -> None:
        super().setGeometry(rect)
        self._do_layout(rect, test_only=False)

    def sizeHint(self) -> QSize:
        return self.minimumSize()

    def minimumSize(self) -> QSize:
        size = QSize()
        for item in self._items:
            hint = item.sizeHint()
            size = size.expandedTo(hint)
        margins = self.contentsMargins()
        size += QSize(margins.left() + margins.right(), margins.top() + margins.bottom())
        width = size.width()
        if width <= 0:
            return size
        return QSize(width, self.heightForWidth(width))

    def _do_layout(self, rect: QRect, *, test_only: bool) -> int:
        margins = self.contentsMargins()
        effective_rect = rect.adjusted(margins.left(), margins.top(), -margins.right(), -margins.bottom())
        x = effective_rect.x()
        y = effective_rect.y()
        line_height = 0
        max_right = effective_rect.right()
        spacing = self.spacing()

        for item in self._items:
            hint = item.sizeHint()
            next_x = x + hint.width()
            if line_height > 0 and next_x > max_right + 1:
                x = effective_rect.x()
                y += line_height + spacing
                next_x = x + hint.width()
                line_height = 0
            if not test_only:
                item.setGeometry(QRect(QPoint(x, y), hint))
            x = next_x + spacing
            line_height = max(line_height, hint.height())

        used_height = y - effective_rect.y() + line_height
        return used_height + margins.top() + margins.bottom()


# ── BatchEntryItemWidget ───────────────────────────────────────────────


class BatchEntryItemWidget(QWidget):
    """A single file entry card within the batch list.

    Layout (top to bottom):
      1. header_row: sequence/status marker + name_label + info_badge
      2. badge_strip: warning badges (WrapRowLayout)
      3. body_section:
         - path_row: source directory (always visible, middle-elided)
         - detail_row: detail text (visible on selected/current/hovered)
         - output_row: output path (always hidden)
      4. actions_row: primary/retry/remove buttons

    Visibility rules (matches old behavior):
      - detail_row: visible when (selected OR current OR hovered) AND has content
      - retry_button/remove_button: visible when hovered and applicable
      - actions_row: visible when any action button is visible
      - body_section: visible while the source path or detail row is visible
    """

    action_requested = Signal(str, str)  # action_key, file_path

    _COMPACT_WIDTH_THRESHOLD = _BATCH_ENTRY_COMPACT_WIDTH_THRESHOLD

    def __init__(
        self,
        entry: BatchFileEntry,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self._entry = entry
        self._sequence_number: int | None = None
        self._list_widget: QListWidget | None = None
        self._list_item: QListWidgetItem | None = None
        self._is_hovered = False
        self._is_selected = False
        self._is_current = False
        self._is_compact = False
        self._status_badge_pulse: QVariantAnimation | None = None
        self._status_badge_pulse_base_size: int | None = None
        self._typography_layout_timer = QTimer(self)
        self._typography_layout_timer.setSingleShot(True)
        self._typography_layout_timer.timeout.connect(self._sync_typography_layout)
        self._secondary_action_visibility: dict[str, bool] = {
            "retry": False,
            "remove": False,
        }

        self.setObjectName("batchEntryCard")
        self.setMouseTracking(True)

        self._build_ui()
        self._apply_entry(entry)
        self._apply_visibility_for_state()

    def _build_ui(self) -> None:
        """Construct the widget layout tree."""
        root = QVBoxLayout(self)
        root.setContentsMargins(_SPACING_SM, _SPACING_SM, _SPACING_SM, _SPACING_SM)
        root.setSpacing(_SPACING_SM)

        # ── Header row ─────────────────────────────────────────────
        header_row = QWidget(self)
        header_row.setObjectName("batchHeaderRow")
        self._header_layout = QBoxLayout(QBoxLayout.Direction.LeftToRight, header_row)
        self._header_layout.setContentsMargins(0, 0, 0, 0)
        self._header_layout.setSpacing(_SPACING_XS)
        self._header_layout.setAlignment(Qt.AlignmentFlag.AlignTop)

        self.status_icon_label = QLabel(header_row)
        self.status_icon_label.setObjectName("batchEntryStatusIcon")
        self.status_icon_label.setMinimumSize(QSize(20, 16))
        self.status_icon_label.setMaximumHeight(18)
        self.status_icon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._header_layout.addWidget(self.status_icon_label, alignment=Qt.AlignmentFlag.AlignTop)

        self.name_label = _InteractivePathLabel(
            self._entry.file_name,
            self._entry.file_path,
            header_row,
        )
        self.name_label.setObjectName("batchEntryName")
        self.name_label.setStyleSheet("font-weight: bold;")
        self.name_label.setWordWrap(True)
        self.name_label.activated.connect(
            lambda: self.action_requested.emit("open_source_location", self._entry.file_path),
        )
        self._header_layout.addWidget(self.name_label, stretch=1)

        self.info_badge = QLabel(header_row)
        self.info_badge.setObjectName("batchInfoBadge")
        self._header_layout.addWidget(self.info_badge, alignment=Qt.AlignmentFlag.AlignTop)

        root.addWidget(header_row)

        # ── Badge strip ────────────────────────────────────────────
        self.badge_strip = QWidget(self)
        self.badge_strip.setObjectName("batchBadgeRow")
        self._badge_strip_layout = WrapRowLayout(self.badge_strip, spacing=_SPACING_XS)
        root.addWidget(self.badge_strip)

        # ── Body section ───────────────────────────────────────────
        self.body_section = QWidget(self)
        self.body_section.setObjectName("batchBodySection")
        body_layout = QVBoxLayout(self.body_section)
        body_layout.setContentsMargins(0, _SPACING_SM, 0, 0)
        body_layout.setSpacing(_SPACING_XS)

        # path_row
        self.path_row = self._create_body_row(
            "",
            "",
            value_object_name="batchPathLabel",
            elide_middle=True,
        )
        body_layout.addWidget(self.path_row)
        # detail_row
        self.detail_row = self._create_body_row("", "", value_object_name="batchDetailLabel")
        body_layout.addWidget(self.detail_row)
        # output_row
        self.output_row = self._create_body_row(
            "Output",
            "",
            row_object_name="batchOutputRow",
            value_object_name="batchOutputLabel",
        )
        body_layout.addWidget(self.output_row)

        root.addWidget(self.body_section)

        # ── Actions row ────────────────────────────────────────────
        self.actions_row = QWidget(self)
        self.actions_row.setObjectName("batchActionsRow")
        actions_layout = QHBoxLayout(self.actions_row)
        actions_layout.setContentsMargins(0, _SPACING_XS, 0, 0)
        actions_layout.setSpacing(_SPACING_XS)
        self.primary_action_button = QPushButton(self.actions_row)
        self.retry_button = QPushButton(self.actions_row)
        self.remove_button = QPushButton(self.actions_row)
        self.primary_action_button.clicked.connect(self._emit_primary_action)
        self.retry_button.clicked.connect(lambda: self.action_requested.emit("retry_failed", self._entry.file_path))
        self.remove_button.clicked.connect(lambda: self.action_requested.emit("remove_entry", self._entry.file_path))
        for btn in (self.primary_action_button, self.retry_button, self.remove_button):
            btn.setMinimumWidth(72)
            actions_layout.addWidget(btn)
        actions_layout.addStretch(1)
        root.addWidget(self.actions_row)

    @staticmethod
    def _create_body_row(
        label_text: str,
        value_text: str,
        *,
        row_object_name: str = "batchInfoRow",
        value_object_name: str,
        elide_middle: bool = False,
    ) -> QWidget:
        row = QWidget()
        row.setObjectName(row_object_name)
        row_layout = QBoxLayout(QBoxLayout.Direction.LeftToRight, row)
        row_layout.setContentsMargins(0, 0, 0, 0)
        row_layout.setSpacing(_SPACING_XS)
        row_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        title = QLabel(label_text, row)
        title.setObjectName("batchInfoLabel")
        row_layout.addWidget(title, alignment=Qt.AlignmentFlag.AlignTop)
        value = _MiddleElidedLabel(value_text, row) if elide_middle else QLabel(value_text, row)
        value.setObjectName(value_object_name)
        if not elide_middle:
            value.setWordWrap(True)
        row_layout.addWidget(value, stretch=1, alignment=Qt.AlignmentFlag.AlignTop)
        return row

    def _apply_entry(self, entry: BatchFileEntry) -> None:
        """Apply BatchFileEntry data to this card's widgets."""
        self._entry = entry
        self._apply_leading_marker()

        # Completed: show pointing hand cursor for opening output
        if entry.status == "completed" and entry.output_path:
            self.status_icon_label.setCursor(Qt.CursorShape.PointingHandCursor)
            self.status_icon_label.setToolTip(_t("components.file_drop.batch_list.action_open_output"))
        else:
            self.status_icon_label.setCursor(Qt.CursorShape.ArrowCursor)

        # Name
        self.name_label.setText(_soft_wrap_filename(entry.file_name))
        self.name_label.set_file_path(entry.file_path)
        self.name_label.setToolTip(entry.file_path)
        self.name_label.setAccessibleName(entry.file_name)
        self.name_label.setAccessibleDescription(entry.file_path)

        # Info badge
        size_str = _format_size(entry.size_bytes)
        status_label = _t(f"components.file_drop.status.{entry.status}", entry.status.title())
        badge_text = f"{entry.detected_format.upper()} · {size_str} · {status_label}"
        if entry.error_count > 0:
            if entry.status == "failed":
                badge_text += f" · {entry.error_count} errors"
            else:
                badge_text += f" · {entry.error_count} issues"
        self.info_badge.setText(badge_text)

        # Body rows
        self._set_row_text(self.path_row, "", _source_path_text(entry.file_path))
        detail_text = self._get_detail_text(entry)
        detail_label = self._get_detail_label_text(entry)
        self._set_row_text(self.detail_row, detail_label, detail_text)
        self._apply_detail_tone(entry)
        output_text = Path(entry.output_path).name if entry.output_path else ""
        self._set_row_text(self.output_row, "Output", output_text)

        # Action buttons
        self._apply_action_buttons(entry)

    def set_sequence_number(self, number: int) -> None:
        """Show the current list position while the entry has no status icon."""
        self._sequence_number = max(1, int(number))
        self._apply_leading_marker()

    def _apply_leading_marker(self) -> None:
        """Render the old list's sequence-or-terminal-status marker."""
        entry = self._entry
        icon = _load_status_icon(entry.status)
        if icon.isNull():
            self.status_icon_label.clear()
            sequence_text = str(self._sequence_number) if self._sequence_number is not None else ""
            self.status_icon_label.setText(sequence_text)
            self.status_icon_label.setVisible(bool(sequence_text))
            self.status_icon_label.setAccessibleName(sequence_text)
            if entry.status != "completed" or not entry.output_path:
                self.status_icon_label.setToolTip(sequence_text)
        else:
            self.status_icon_label.clear()
            self.status_icon_label.setPixmap(icon.pixmap(QSize(16, 16)))
            self.status_icon_label.setVisible(True)
            status_text = _t(f"components.file_drop.status.{entry.status}", entry.status.title())
            self.status_icon_label.setAccessibleName(status_text)
            if entry.status != "completed" or not entry.output_path:
                self.status_icon_label.setToolTip(status_text)

    @staticmethod
    def _set_row_text(row: QWidget, label: str, value: str) -> None:
        """Set label and value text on a body row."""
        layout = row.layout()
        if layout is None or layout.count() < 2:
            return
        title = layout.itemAt(0)
        val = layout.itemAt(1)
        title_widget = title.widget() if title is not None else None
        if title_widget is not None:
            title_widget.setText(label)  # pyright: ignore[reportOptionalMemberAccess, reportAttributeAccessIssue]
            title_widget.setVisible(bool(label.strip()))
        if val is not None and val.widget() is not None:
            value_widget = val.widget()
            if isinstance(value_widget, _MiddleElidedLabel):
                value_widget.set_full_text(value)
            else:
                value_widget.setText(value)  # pyright: ignore[reportOptionalMemberAccess, reportAttributeAccessIssue]

    def _get_detail_text(self, entry: BatchFileEntry) -> str:
        """Get detail text for the detail row."""
        if entry.error_message:
            return entry.error_message
        if entry.skip_reason:
            return entry.skip_reason
        if entry.warning_message:
            return entry.warning_message
        return ""

    def _get_detail_label_text(self, entry: BatchFileEntry) -> str:
        """Get the label for the detail row."""
        if entry.status == "cancelled":
            return _t("components.file_drop.status.cancelled", "Cancelled")
        if entry.error_message:
            return _t("common.error", "Error")
        if entry.skip_reason:
            return _t("components.file_drop.status.skipped", "Skipped")
        return _t("editors.common.description", "Description")

    def _apply_detail_tone(self, entry: BatchFileEntry) -> None:
        """Bind the visible detail row to the existing semantic QSS contract."""
        title = self._get_row_label_widget(self.detail_row)
        value = self._get_row_value_widget(self.detail_row)
        tone = "secondary"
        title_tone = "secondary"
        if entry.error_message:
            tone = title_tone = "danger"
        elif entry.skip_reason:
            tone = title_tone = "warning"
        elif entry.warning_message:
            tone = "warning"
        apply_theme_class(title, title_tone)
        if value is not None:
            value.setProperty("detailRole", tone)
            apply_theme_class(value, tone)

    def _apply_action_buttons(self, entry: BatchFileEntry) -> None:
        """Configure action buttons based on entry status."""
        status = entry.status

        # Primary action
        if status in {"completed", "failed"} and entry.output_path:
            self._set_button(
                self.primary_action_button,
                _t("components.file_drop.batch_list.action_open_output"),
                True,
            )
            self._primary_action_key = "open_output"
        elif status == "skipped" and entry.skip_reason:
            self._set_button(
                self.primary_action_button,
                _t("components.file_drop.batch_list.action_view_skip"),
                True,
            )
            self._primary_action_key = "show_skip_details"
        elif status == "failed" and entry.error_message:
            self._set_button(
                self.primary_action_button,
                _t("components.file_drop.batch_list.action_view_error"),
                True,
            )
            self._primary_action_key = "show_error_details"
        else:
            self.primary_action_button.hide()
            self._primary_action_key = None

        # Retry button
        if status == "failed":
            self._set_button(
                self.retry_button,
                _t("components.file_drop.batch_list.retry_selected_failed"),
                True,
            )
            self._secondary_action_visibility["retry"] = True
        else:
            self.retry_button.hide()
            self._secondary_action_visibility["retry"] = False

        # Remove button - always present but conditionally visible
        self._set_button(
            self.remove_button,
            _t("components.file_drop.batch_list.action_remove"),
            True,
        )
        self._secondary_action_visibility["remove"] = True

    @staticmethod
    def _set_button(button: QPushButton, text: str, enabled: bool) -> None:
        button.setText(text)
        button.setEnabled(enabled)
        if enabled:
            button.show()
        else:
            button.hide()

    def _emit_primary_action(self) -> None:
        if self._primary_action_key:
            self.action_requested.emit(self._primary_action_key, self._entry.file_path)

    # ── Pulse animation ────────────────────────────────────────────────

    def pulse_processing_transition(self) -> None:
        """Play the 220ms processing pulse animation on the info badge."""
        if self._status_badge_pulse is not None:
            self._status_badge_pulse.stop()
            self._status_badge_pulse.deleteLater()
        self._status_badge_pulse_base_size = self.info_badge.font().pointSize()
        animation = QVariantAnimation(self)
        animation.setDuration(220)
        animation.setStartValue(0)
        animation.setKeyValueAt(0.5, 2)
        animation.setEndValue(0)
        animation.valueChanged.connect(lambda value: self._apply_pulse(int(value)))
        animation.finished.connect(self._finish_pulse)
        self._status_badge_pulse = animation
        animation.start()

    def _apply_pulse(self, delta: int) -> None:
        font = self.info_badge.font()
        base_size = self._status_badge_pulse_base_size
        if base_size is None:
            base_size = font.pointSize()
        if base_size <= 0:
            return
        font.setPointSize(max(1, base_size + int(delta)))
        self.info_badge.setFont(font)

    def _finish_pulse(self) -> None:
        self._apply_pulse(0)
        self._status_badge_pulse_base_size = None

    # ── Visibility / interaction state ─────────────────────────────────

    def set_interaction_state(self, *, selected: bool, current: bool) -> None:
        self._is_selected = selected
        self._is_current = current
        self._refresh_card_state()

    def enterEvent(self, event) -> None:
        self._is_hovered = True
        self._refresh_card_state()
        super().enterEvent(event)

    def leaveEvent(self, event) -> None:
        self._is_hovered = False
        self._refresh_card_state()
        super().leaveEvent(event)

    def _refresh_card_state(self) -> None:
        self.setProperty("stateHovered", self._is_hovered)
        self.setProperty("stateSelected", self._is_selected)
        self.setProperty("stateCurrent", self._is_current)
        self._apply_visibility_for_state()
        self._sync_item_size_hint()
        self.style().unpolish(self)
        self.style().polish(self)
        self.update()

    def _apply_visibility_for_state(self) -> None:
        """Apply visibility rules matching old behavior.

        - badge_strip: always visible
        - info_badge: always visible
        - path_row: always visible when a source directory is available
        - detail_row: visible when (selected OR current OR hovered) AND has content
        - output_row: always hidden
        - retry_button: visible when hovered/selected/current and applicable
        - remove_button: visible when hovered/selected/current and applicable
        - actions_row: visible when any action button is visible
        - body_section: visible while the source path or detail row is visible
        """
        expanded = self._is_selected or self._is_current or self._is_hovered
        self.badge_strip.setVisible(True)
        self.info_badge.setVisible(True)
        path_label = self._get_row_value_widget(self.path_row)
        if isinstance(path_label, _MiddleElidedLabel):
            has_path = bool(path_label.full_text.strip())
        else:
            has_path = bool(path_label.text().strip()) if path_label else False
        self.path_row.setVisible(has_path)
        # detail_row visible when expanded and has content
        detail_label = self._get_row_value_widget(self.detail_row)
        has_detail = bool(detail_label.text().strip()) if detail_label else False
        self.detail_row.setVisible(expanded and has_detail)
        if detail_label is not None:
            line_height = detail_label.fontMetrics().lineSpacing()
            detail_label.setMinimumHeight(line_height * 3 if expanded and has_detail else 0)
        self.output_row.setHidden(True)
        self.retry_button.setVisible(expanded and self._secondary_action_visibility.get("retry", False))
        self.remove_button.setVisible(expanded and self._secondary_action_visibility.get("remove", False))
        any_visible = any(
            btn.isVisible() for btn in (self.primary_action_button, self.retry_button, self.remove_button)
        )
        self.actions_row.setVisible(any_visible)
        self.body_section.setVisible(has_path or not self.detail_row.isHidden())

    @staticmethod
    def _get_row_label_widget(row: QWidget) -> QLabel | None:
        layout = row.layout()
        if layout is None or layout.count() < 1:
            return None
        item = layout.itemAt(0)
        if item is not None and isinstance(item.widget(), QLabel):
            return _cast(QLabel, item.widget())
        return None

    @staticmethod
    def _get_row_value_widget(row: QWidget) -> QLabel | None:
        layout = row.layout()
        if layout is None or layout.count() < 2:
            return None
        item = layout.itemAt(1)
        if item is not None and isinstance(item.widget(), QLabel):
            return _cast(QLabel, item.widget())
        return None

    # ── List item binding ──────────────────────────────────────────────

    def bind_list_item(self, list_widget: QListWidget, item: QListWidgetItem) -> None:
        self._list_widget = list_widget
        self._list_item = item
        self._sync_item_size_hint()

    def _sync_item_size_hint(self) -> None:
        if self._list_item is None:
            return
        try:
            import shiboken6

            if not shiboken6.isValid(self._list_item):
                return
        except ImportError:
            pass
        viewport = self._list_widget.viewport() if self._list_widget is not None else None
        width = viewport.width() if viewport is not None else self.width()
        layout = self.layout()
        if width > 0 and layout is not None:
            hint_height = layout.totalHeightForWidth(width)
            if hint_height > 0:
                try:
                    self._list_item.setSizeHint(QSize(width, hint_height))
                except RuntimeError:
                    return
                if self._list_widget is not None:
                    self._list_widget.doItemsLayout()
                return
        try:
            self._list_item.setSizeHint(self.sizeHint())
            if self._list_widget is not None:
                self._list_widget.doItemsLayout()
        except RuntimeError:
            return

    def resizeEvent(self, event) -> None:
        super().resizeEvent(event)
        self._apply_compact_mode()
        self._sync_item_size_hint()

    def changeEvent(self, event) -> None:
        super().changeEvent(event)
        if event.type() == QEvent.Type.PaletteChange and hasattr(self, "status_icon_label"):
            self._apply_leading_marker()
        if event.type() in {QEvent.Type.FontChange, QEvent.Type.StyleChange} and hasattr(self, "name_label"):
            self._typography_layout_timer.start(0)

    def _sync_typography_layout(self) -> None:
        self._apply_compact_mode()
        self._sync_item_size_hint()

    def _apply_compact_mode(self) -> None:
        _clayout = _cast(QLayout, self.layout())
        content_width = max(
            0,
            self.width() - _clayout.contentsMargins().left() - _clayout.contentsMargins().right(),
        )
        # The decision must not depend on live child geometry or visibility:
        # TopToBottom stretches the marker to the full row width, which used to
        # feed back into this calculation and make compact mode irreversible.
        # Scale the fixed design threshold only with the global typography
        # preset so every status in a category makes the same decision.
        from docwen_gui.styles.theme_manager import ThemeManager

        font_size = resolve_font_size_preset(ThemeManager.get_instance().get_font_size_preset())
        compact_threshold = round(self._COMPACT_WIDTH_THRESHOLD * font_size / DEFAULT_FONT_SIZE)
        compact = 0 < content_width < compact_threshold
        if compact == self._is_compact:
            return
        self._is_compact = compact
        direction = QBoxLayout.Direction.TopToBottom if compact else QBoxLayout.Direction.LeftToRight
        self._header_layout.setDirection(direction)
        spacing = _SPACING_SM if compact else _SPACING_XS
        self._header_layout.setSpacing(spacing)
        if compact:
            self.name_label.setWordWrap(True)
        self.updateGeometry()

    # ── Action helpers ─────────────────────────────────────────────────

    def copy_error_details(self) -> bool:
        """Copy error message to clipboard. Returns True on success."""
        error_text = (self._entry.error_message or "").strip()
        if not error_text:
            return False
        try:
            QApplication.clipboard().setText(error_text)
            return True
        except Exception:
            return False


# ── Helper functions ───────────────────────────────────────────────────


def _format_size(size_bytes: int) -> str:
    value = float(max(0, size_bytes))
    units = ["B", "KB", "MB", "GB"]
    for unit in units:
        if value < 1024 or unit == units[-1]:
            if unit == "B":
                return f"{int(value)} {unit}"
            return f"{value:.1f} {unit}"
        value /= 1024.0
    return f"{int(size_bytes)} B"


def _source_path_text(file_path: str) -> str:
    path = Path(file_path)
    parent_text = str(path.parent)
    if parent_text in {"", "."}:
        return file_path
    return parent_text


def _path_compare_key(file_path: str) -> str:
    """Normalize a path for comparing raw Windows paths and VM keys."""
    return str(Path(file_path)).replace("\\", "/")


# ── BatchList main widget ──────────────────────────────────────────────


class BatchList(QWidget):
    """Main batch file list widget with 6 category tabs.

    This is the user-visible component. All state lives in the ViewModel.
    The widget only renders state and delegates user actions to the ViewModel.

    Signals:
        entry_action_requested: emitted when a user clicks on an entry action.
    """

    entry_action_requested = Signal(str, str)  # action_key, file_path
    selection_changed = Signal(object)  # current file_path or None

    def __init__(
        self,
        view_model: BatchListViewModel,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self._vm = view_model
        self._tabs: dict[str, ReorderableListWidget] = {}
        self._pivot_items: dict[str, QWidget] = {}
        self._pivot_compact_mode = False
        self._suspend_selection_sync = False
        self._suspend_tab_selection_sync = False
        self._pending_entry_widget_attachments: deque[tuple[QListWidget, QListWidgetItem, BatchFileEntry]] = deque()
        self._entry_widget_attach_timer = QTimer(self)
        self._entry_widget_attach_timer.setSingleShot(True)
        self._entry_widget_attach_timer.timeout.connect(self._attach_pending_entry_widgets)

        self.setObjectName("batchListSurface")
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)

        self._build_ui()
        self._wire_view_model()
        self._sync_existing_entries()

    @property
    def view_model(self) -> BatchListViewModel:
        return self._vm

    # ── UI construction ────────────────────────────────────────────────

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(_SPACING_XS, _SPACING_XS, _SPACING_XS, _SPACING_XS)
        layout.setSpacing(_SPACING_SM)

        # ── Summary section ─────────────────────────────────────────
        self._build_summary_section()
        layout.addWidget(self.summary_section)

        # ── Tab section ─────────────────────────────────────────────
        self.tabs_frame = QFrame(self)
        self.tabs_frame.setObjectName("batchListTabsFrame")
        tabs_layout = QVBoxLayout(self.tabs_frame)
        tabs_layout.setContentsMargins(_SPACING_XS, _SPACING_XS, _SPACING_XS, _SPACING_XS)
        tabs_layout.setSpacing(_SPACING_XS)

        try:
            from qfluentwidgets import Pivot as FluentPivot
        except ImportError:
            FluentPivot = None  # fallback: use a plain widget

        if FluentPivot is not None:
            self.category_pivot = FluentPivot(self.tabs_frame)
            self.category_pivot.installEventFilter(self)
        else:
            # Fallback: use a horizontal layout with buttons
            pivot_container = QWidget(self.tabs_frame)
            QHBoxLayout(pivot_container)
            self.category_pivot = pivot_container
        self.category_pivot.setObjectName("batchCategoryPivot")
        self.category_pivot.setFocusPolicy(Qt.FocusPolicy.TabFocus)

        self.category_stack = QStackedWidget(self.tabs_frame)
        self.category_stack.setObjectName("batchCategoryStack")

        tabs_layout.addWidget(self.category_pivot, 0)
        tabs_layout.addWidget(self.category_stack, 1)

        layout.addWidget(self.summary_section)
        layout.addWidget(self.tabs_frame, 1)

        # Build 6 category tabs
        for category in _CATEGORY_ORDER:
            list_widget = ReorderableListWidget(self.category_stack)
            list_widget.setObjectName("batchListWidget")
            list_widget.setProperty("batchCategory", category)
            list_widget.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
            list_widget.setAlternatingRowColors(True)
            list_widget.setVerticalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
            list_widget.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
            list_widget.setFrameShape(QFrame.Shape.NoFrame)
            list_widget.setSpacing(_SPACING_XS)
            list_widget.itemSelectionChanged.connect(self._handle_selection_changed)
            list_widget.currentItemChanged.connect(self._handle_current_item_changed)
            list_widget.item_reordered.connect(self._handle_manual_reorder)
            list_widget.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
            list_widget.customContextMenuRequested.connect(
                lambda pos, cat=category: self._show_item_context_menu(cat, pos)
            )
            self._tabs[category] = list_widget
            self.category_stack.addWidget(list_widget)

            if isinstance(self.category_pivot, FluentPivot) if FluentPivot else False:
                label = self._category_tab_label(category, 0)
                self.category_pivot.addItem(  # pyright: ignore[reportAttributeAccessIssue]
                    category,
                    label,
                    onClick=lambda _checked=False, c=category: self._activate_tab(c),
                )
                pivot_item = getattr(self.category_pivot, "items", {}).get(category)
                if pivot_item is not None:
                    self._pivot_items[category] = pivot_item

        # Activate the ViewModel's current tab when the widget is created from
        # pre-populated state, such as app startup with initial files.
        initial_category = self._vm.current_category
        if initial_category not in self._tabs:
            initial_category = _CATEGORY_ORDER[0]
        self._activate_tab(initial_category)
        self._apply_focus_navigation()

    def _build_summary_section(self) -> None:
        """Build summary bar with filter, summary, reorder, sort buttons."""
        self.summary_section = QFrame(self)
        self.summary_section.setObjectName("batchListSummarySection")
        summary_layout = QVBoxLayout(self.summary_section)
        summary_layout.setContentsMargins(0, 0, 0, 0)
        summary_layout.setSpacing(_SPACING_XS)

        summary_header = QFrame(self.summary_section)
        summary_header.setObjectName("batchListSummaryHeader")
        self._summary_header = summary_header
        self._summary_header_compact = False
        header_row = QGridLayout(summary_header)
        self._summary_header_layout = header_row
        header_row.setContentsMargins(0, 0, 0, 0)
        header_row.setSpacing(_SPACING_XS)

        self.filter_button = QPushButton(_t("components.file_drop.batch_list.filter_button", "Filter: All"))
        self.filter_button.setObjectName("batchFilterButton")
        self.filter_button.clicked.connect(self._show_filter_menu)
        header_row.addWidget(self.filter_button, 0, 0, alignment=Qt.AlignmentFlag.AlignLeft)
        header_row.setColumnStretch(0, 1)

        reorder_frame = QFrame(summary_header)
        reorder_row = QHBoxLayout(reorder_frame)
        reorder_row.setContentsMargins(0, 0, 0, 0)
        reorder_row.setSpacing(_SPACING_XS)

        self.move_up_button = QPushButton(_t("components.file_drop.batch_list.action_move_up", "Move Up"))
        self.move_down_button = QPushButton(_t("components.file_drop.batch_list.action_move_down", "Move Down"))
        self.move_up_button.setObjectName("batchReorderButton")
        self.move_down_button.setObjectName("batchReorderButton")
        move_up_hint = f"{self.move_up_button.text()} (Ctrl+↑)"
        move_down_hint = f"{self.move_down_button.text()} (Ctrl+↓)"
        self.move_up_button.setToolTip(move_up_hint)
        self.move_down_button.setToolTip(move_down_hint)
        self.move_up_button.setAccessibleName(self.move_up_button.text())
        self.move_down_button.setAccessibleName(self.move_down_button.text())
        self.move_up_button.setAccessibleDescription(move_up_hint)
        self.move_down_button.setAccessibleDescription(move_down_hint)
        self.move_up_button.clicked.connect(self._move_current_item_up)
        self.move_down_button.clicked.connect(self._move_current_item_down)
        reorder_row.addWidget(self.move_up_button)
        reorder_row.addWidget(self.move_down_button)

        self.sort_button = QPushButton(_t("components.file_drop.batch_list.sort_button", "Sort: Custom"))
        self.sort_button.setObjectName("batchSortButton")
        self.sort_button.clicked.connect(self._show_sort_menu)
        reorder_row.addWidget(self.sort_button)

        self._summary_reorder_frame = reorder_frame
        header_row.addWidget(reorder_frame, 0, 1, alignment=Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignTop)
        summary_header.installEventFilter(self)
        summary_layout.addWidget(summary_header)

        self.summary_label = QLabel(_t("components.file_drop.batch_list.no_batch_files", "No batch files"))
        self.summary_label.setObjectName("batchListSummaryLabel")
        self.summary_label.setWordWrap(True)
        summary_layout.addWidget(self.summary_label)

    # ── ViewModel wiring ───────────────────────────────────────────────

    def _wire_view_model(self) -> None:
        vm = self._vm
        vm.files_added.connect(self._on_files_added)
        vm.files_removed.connect(self._on_file_removed)
        vm.files_cleared.connect(self._on_files_cleared)
        vm.status_changed.connect(self._on_status_changed)
        vm.filter_changed.connect(self._on_filter_changed)
        vm.sort_changed.connect(self._on_sort_changed)
        vm.current_category_changed.connect(self._on_category_changed)
        vm.entry_count_changed.connect(self._on_entry_count_changed)
        vm.pulse_requested.connect(self._on_pulse_requested)

    def _sync_existing_entries(self) -> None:
        """Render ViewModel entries that were added before widget construction."""
        existing = self._vm.get_files()
        if existing:
            self._on_files_added(existing, [])
        self._on_filter_changed(self._vm.active_filter)
        self._on_sort_changed(self._vm.sort_key, self._vm.sort_ascending)
        self._refresh_titles()
        self._update_reorder_buttons()
        self._refresh_filter_button()

    def _on_files_added(self, added: list[str], _failed: list[tuple[str, str]]) -> None:
        for file_path in added:
            entry = self._vm.get_file_entry(file_path)
            if entry is None:
                continue
            category = self._vm.get_file_display_category(file_path)
            if category not in self._tabs:
                category = "other"
            list_widget = self._tabs[category]
            item = QListWidgetItem("")
            item.setData(Qt.ItemDataRole.UserRole, file_path)
            item.setData(Qt.ItemDataRole.UserRole + 1, entry)
            list_widget.addItem(item)
            self._pending_entry_widget_attachments.append((list_widget, item, entry))
        if self._pending_entry_widget_attachments and not self._entry_widget_attach_timer.isActive():
            self._attach_pending_entry_widgets()
        self._refresh_titles()
        self._update_reorder_buttons()
        self._refresh_filter_button()

        if _failed:
            from .batch_dialogs import show_batch_add_failed_dialog

            show_batch_add_failed_dialog(self, _failed)

    def _on_file_removed(self, file_path: str) -> None:
        for list_widget in self._tabs.values():
            for index in range(list_widget.count()):
                item = list_widget.item(index)
                data = item.data(Qt.ItemDataRole.UserRole)
                if data == file_path:
                    widget = list_widget.itemWidget(item)
                    if widget is not None:
                        widget.deleteLater()
                    list_widget.takeItem(index)
                    self._refresh_sequence_numbers(list_widget)
                    self._refresh_titles()
                    self._update_reorder_buttons()
                    return

    def _on_files_cleared(self) -> None:
        self._entry_widget_attach_timer.stop()
        self._pending_entry_widget_attachments.clear()
        for list_widget in self._tabs.values():
            list_widget.clear()
        self._refresh_titles()
        self._update_reorder_buttons()
        self._refresh_filter_button()

    def _on_status_changed(self, file_path: str, new_status: str) -> None:
        entry = self._vm.get_file_entry(file_path)
        if entry is None:
            return
        # Find the item and update its widget
        for list_widget in self._tabs.values():
            for index in range(list_widget.count()):
                item = list_widget.item(index)
                data = item.data(Qt.ItemDataRole.UserRole)
                if data == file_path:
                    widget = list_widget.itemWidget(item)
                    if isinstance(widget, BatchEntryItemWidget):
                        widget._apply_entry(entry)
                        widget.bind_list_item(list_widget, item)
                    else:
                        self._attach_entry_widget(list_widget, item, entry)
                    self._update_widget_state(list_widget, item)
                    self._refresh_titles()
                    return

    def _on_filter_changed(self, _filter_key: str) -> None:
        for list_widget in self._tabs.values():
            for index in range(list_widget.count()):
                item = list_widget.item(index)
                entry = item.data(Qt.ItemDataRole.UserRole + 1)
                from ..view_models.batch_list_vm import BatchFileEntry as BFE

                if isinstance(entry, BFE):
                    visible = self._vm._entry_matches_filter(entry)
                    item.setHidden(not visible)
                    widget = list_widget.itemWidget(item)
                    if widget is not None:
                        widget.setVisible(visible)
            self._ensure_visible_current_item(list_widget)
        self._refresh_titles()
        self._update_reorder_buttons()
        self._refresh_all_item_states()
        self._refresh_filter_button()

    def _on_sort_changed(self, _sort_key: str, _ascending: bool) -> None:
        for category in _CATEGORY_ORDER:
            self._apply_sort_for_category(category)
        self._refresh_sort_button()
        self._refresh_titles()

    def _on_category_changed(self, category: str) -> None:
        if category not in self._tabs:
            return
        self._apply_active_tab(category)
        if not self._suspend_tab_selection_sync:
            self._handle_selection_changed()

    def _on_entry_count_changed(self, _count: int) -> None:
        self._refresh_summary()
        self._refresh_pivot_labels()

    def _on_pulse_requested(self, file_path: str) -> None:
        for list_widget in self._tabs.values():
            for index in range(list_widget.count()):
                item = list_widget.item(index)
                if item.data(Qt.ItemDataRole.UserRole) == file_path:
                    widget = list_widget.itemWidget(item)
                    if isinstance(widget, BatchEntryItemWidget):
                        widget.pulse_processing_transition()
                    return

    # ── Tab navigation ─────────────────────────────────────────────────

    def _activate_tab(self, category: str) -> None:
        if category not in self._tabs:
            return
        if not self._vm.activate_tab(category):
            self._apply_active_tab(category)

    def _apply_active_tab(self, category: str) -> None:
        """Render the active VM tab without mutating ViewModel state."""
        list_widget = self._tabs[category]
        self.category_stack.setCurrentWidget(list_widget)
        with contextlib.suppress(Exception):
            if hasattr(self.category_pivot, "setCurrentItem"):
                self.category_pivot.setCurrentItem(category)  # pyright: ignore[reportAttributeAccessIssue]
        self._refresh_pivot_labels()
        self._refresh_pivot_states()

    def _category_tab_label(self, category: str, count: int, *, include_count: bool = True) -> str:
        """Build the tab label: 'Category · N'."""
        label = _t(f"file_types.{category}", category)
        if include_count and count > 0:
            label = f"{label} · {count}"
        return label

    # ── Event handling ─────────────────────────────────────────────────

    def focusInEvent(self, event) -> None:
        super().focusInEvent(event)
        self._focus_current_list()

    def _focus_current_list(self) -> None:
        category = self._vm.current_category
        if category not in self._tabs:
            return
        list_widget = self._tabs[category]
        if list_widget.count() > 0:
            list_widget.setFocus(Qt.FocusReason.OtherFocusReason)

    def resizeEvent(self, event) -> None:
        super().resizeEvent(event)
        self._sync_summary_header_layout()
        self._refresh_pivot_labels()

    def eventFilter(self, watched, event) -> bool:
        if watched is self._summary_header and event.type() == QEvent.Type.LayoutRequest:
            self._sync_summary_header_layout()
        if (
            watched is self.category_pivot
            and event.type() == QEvent.Type.KeyPress
            and self._handle_pivot_key_press(event)
        ):
            return True
        return super().eventFilter(watched, event)

    def _sync_summary_header_layout(self) -> None:
        """Wrap summary controls instead of squeezing their labels at large type sizes."""
        available_width = self._summary_header.contentsRect().width()
        if available_width <= 0:
            return
        required_width = (
            self.filter_button.sizeHint().width()
            + self._summary_reorder_frame.sizeHint().width()
            + self._summary_header_layout.horizontalSpacing()
        )
        compact = required_width + (_SPACING_SM * 2) > available_width
        if compact == self._summary_header_compact:
            return
        self._summary_header_compact = compact
        self._summary_header_layout.removeWidget(self._summary_reorder_frame)
        row = 1 if compact else 0
        column = 0 if compact else 1
        alignment = Qt.AlignmentFlag.AlignLeft if compact else Qt.AlignmentFlag.AlignRight
        self._summary_header_layout.addWidget(
            self._summary_reorder_frame,
            row,
            column,
            alignment=alignment | Qt.AlignmentFlag.AlignTop,
        )
        self._summary_header.updateGeometry()

    def _handle_pivot_key_press(self, event: QKeyEvent) -> bool:
        if event.modifiers() != Qt.KeyboardModifier.NoModifier:
            return False
        if event.key() not in {Qt.Key.Key_Left, Qt.Key.Key_Right}:
            return False
        current = self._vm.current_category
        try:
            idx = _CATEGORY_ORDER.index(current)
        except ValueError:
            idx = 0
        offset = -1 if event.key() == Qt.Key.Key_Left else 1
        next_idx = max(0, min(len(_CATEGORY_ORDER) - 1, idx + offset))
        if next_idx == idx:
            return False
        self._activate_tab(_CATEGORY_ORDER[next_idx])
        return True

    def _apply_focus_navigation(self) -> None:
        """Set up Tab order: category_pivot -> list -> up -> down -> sort -> pivot."""
        for list_widget in self._tabs.values():
            self.setTabOrder(self.category_pivot, list_widget)
        self.setTabOrder(self._tabs[_CATEGORY_ORDER[-1]], self.move_up_button)
        self.setTabOrder(self.move_up_button, self.move_down_button)
        self.setTabOrder(self.move_down_button, self.sort_button)
        self.setTabOrder(self.sort_button, self.category_pivot)

    # ── Selection management ───────────────────────────────────────────

    def _handle_selection_changed(self) -> None:
        self._refresh_all_item_states()
        self._refresh_summary()
        self._update_reorder_buttons()
        current = self.get_current_file()
        self.selection_changed.emit(current)

    def _handle_current_item_changed(self, current: QListWidgetItem | None, _previous: QListWidgetItem | None) -> None:
        if current is None or self._suspend_selection_sync:
            return
        self._handle_selection_changed()

    def get_current_file(self) -> str | None:
        category = self._vm.current_category
        if category not in self._tabs:
            return None
        list_widget = self._tabs[category]
        item = list_widget.currentItem()
        if item is not None and item.isHidden():
            item = None
        if item is None:
            selected = [candidate for candidate in list_widget.selectedItems() if not candidate.isHidden()]
            if not selected:
                item = self._first_visible_item(list_widget)
                if item is None:
                    return None
            else:
                item = selected[0]
        file_path = item.data(Qt.ItemDataRole.UserRole)
        return file_path if isinstance(file_path, str) else None

    def get_selected_files(self, category: str | None = None) -> list[str]:
        """Get selected file paths for a category."""
        cat = category or self._vm.current_category
        if cat not in self._tabs:
            return []
        list_widget = self._tabs[cat]
        selected: list[str] = []
        for item in list_widget.selectedItems():
            if item.isHidden():
                continue
            fp = item.data(Qt.ItemDataRole.UserRole)
            if isinstance(fp, str):
                selected.append(fp)
        if selected:
            return selected
        current = self.get_current_file()
        return [current] if current else []

    def _refresh_all_item_states(self) -> None:
        for _category, list_widget in self._tabs.items():
            for index in range(list_widget.count()):
                item = list_widget.item(index)
                self._update_widget_state(list_widget, item)

    def _update_widget_state(self, list_widget: QListWidget, item: QListWidgetItem) -> None:
        widget = list_widget.itemWidget(item)
        if not isinstance(widget, BatchEntryItemWidget):
            return
        is_current_tab = self._tabs.get(self._vm.current_category) is list_widget
        widget.set_interaction_state(
            selected=item.isSelected(),
            current=is_current_tab and list_widget.currentItem() is item,
        )

    @staticmethod
    def _first_visible_item(list_widget: QListWidget) -> QListWidgetItem | None:
        for index in range(list_widget.count()):
            item = list_widget.item(index)
            if not item.isHidden():
                return item
        return None

    def _ensure_visible_current_item(self, list_widget: QListWidget) -> None:
        current = list_widget.currentItem()
        if current is not None and not current.isHidden():
            return
        fallback = self._first_visible_item(list_widget)
        self._suspend_selection_sync = True
        try:
            list_widget.clearSelection()
            if fallback is None:
                list_widget.setCurrentRow(-1)
                return
            fallback.setSelected(True)
            list_widget.setCurrentItem(fallback)
        finally:
            self._suspend_selection_sync = False

    # ── Entry widget management ────────────────────────────────────────

    def _attach_pending_entry_widgets(self) -> None:
        """Materialize one bounded chunk of queued entry cards."""
        self._entry_widget_attach_timer.stop()
        attached = 0
        while self._pending_entry_widget_attachments and attached < _BATCH_ENTRY_WIDGET_ATTACH_CHUNK_SIZE:
            list_widget, item, entry = self._pending_entry_widget_attachments.popleft()
            if item.listWidget() is not list_widget or list_widget.itemWidget(item) is not None:
                continue
            self._attach_entry_widget(list_widget, item, entry)
            attached += 1
        if self._pending_entry_widget_attachments:
            self._entry_widget_attach_timer.start(0)

    def _attach_entry_widget(self, list_widget: QListWidget, item: QListWidgetItem, entry: BatchFileEntry) -> None:
        widget = BatchEntryItemWidget(entry, list_widget)
        widget.action_requested.connect(self._handle_entry_action)
        list_widget.setItemWidget(item, widget)
        widget.bind_list_item(list_widget, item)
        widget.set_sequence_number(list_widget.row(item) + 1)
        self._update_widget_state(list_widget, item)

    @staticmethod
    def _refresh_sequence_numbers(list_widget: QListWidget) -> None:
        """Rebind visible order markers after remove, sort, or manual reorder."""
        for index in range(list_widget.count()):
            item = list_widget.item(index)
            widget = list_widget.itemWidget(item)
            if isinstance(widget, BatchEntryItemWidget):
                widget.set_sequence_number(index + 1)

    def _handle_entry_action(self, action_key: str, file_path: str) -> None:
        """Handle an action from an entry widget."""
        if action_key == "retry_failed":
            self._vm.reset_failed_files([file_path])
            return
        if action_key == "remove_entry":
            self._vm.remove_file(file_path)
            return
        # Forward to external handlers
        self.entry_action_requested.emit(action_key, file_path)

    # ── Refresh UI ─────────────────────────────────────────────────────

    def _refresh_titles(self) -> None:
        self._refresh_pivot_labels()
        self._refresh_pivot_states()
        self._refresh_summary()

    def _refresh_pivot_labels(self) -> None:
        if not self._pivot_items:
            return
        self._pivot_compact_mode = self.width() <= _BATCH_CATEGORY_PIVOT_NARROW_THRESHOLD
        for category, item in self._pivot_items.items():
            count = self._vm.get_visible_count_for_category(category)
            label = self._category_tab_label(
                category,
                count,
                include_count=not self._pivot_compact_mode,
            )
            if hasattr(item, "setText"):
                item.setText(label)  # pyright: ignore[reportAttributeAccessIssue]

    def _refresh_pivot_states(self) -> None:
        if not self._pivot_items:
            return
        current = self._vm.current_category
        for category, item in self._pivot_items.items():
            count = self._vm.get_visible_count_for_category(category)
            is_current = category == current
            has_content = count > 0
            item.setProperty("isCurrentRoute", is_current)
            item.setProperty("hasContent", has_content)
            style = item.style()
            if style is not None:
                style.unpolish(item)
                style.polish(item)

    def _refresh_summary(self) -> None:
        total = self._vm.entry_count
        if total <= 0:
            text = _t("components.file_drop.batch_list.no_batch_files", "No batch files")
            self.summary_label.setText(text)
            self.summary_section.setProperty("hasFiles", False)
        else:
            current = self.get_current_file()
            if current:
                text = _t(
                    "components.file_drop.batch_list.total_files_selected",
                    f"{total} files loaded, current: {{current_name}}",
                    total=total,
                    current_name=Path(current).name,
                )
            else:
                text = _t(
                    "components.file_drop.batch_list.total_files",
                    f"{total} files loaded",
                    total=total,
                )
            self.summary_label.setText(text)
            self.summary_section.setProperty("hasFiles", True)
        self.summary_section.style().unpolish(self.summary_section)
        self.summary_section.style().polish(self.summary_section)

    def _refresh_filter_button(self) -> None:
        filter_key = self._vm.active_filter
        from ..view_models.batch_list_vm import FILTER_OPTIONS

        label = "All"
        for fkey, fname, _statuses in FILTER_OPTIONS:
            if fkey == filter_key:
                label = _filter_option_label(fkey, fname)
                break
        text = _t(
            "components.file_drop.batch_list.filter_button_with_label",
            f"Filter: {label}",
            label=label,
        )
        self.filter_button.setText(text)

    def _refresh_sort_button(self) -> None:
        label = _sort_option_label(self._vm.sort_key)
        text = _t(
            "components.file_drop.batch_list.sort_button_with_label",
            f"Sort: {label}",
            label=label,
        )
        self.sort_button.setText(text)
        self.sort_button.setToolTip(text)

    def _update_reorder_buttons(self) -> None:
        category = self._vm.current_category
        if category not in self._tabs:
            self.move_up_button.setEnabled(False)
            self.move_down_button.setEnabled(False)
            return
        list_widget = self._tabs[category]
        visible_items = self._visible_items(list_widget)
        current_item = list_widget.currentItem()
        current_index = visible_items.index(current_item) if current_item in visible_items else -1
        self.move_up_button.setEnabled(current_index > 0)
        self.move_down_button.setEnabled(0 <= current_index < len(visible_items) - 1)

    @staticmethod
    def _visible_item_count(list_widget: QListWidget) -> int:
        return sum(1 for i in range(list_widget.count()) if not list_widget.item(i).isHidden())

    @staticmethod
    def _visible_items(list_widget: QListWidget) -> list[QListWidgetItem]:
        return [list_widget.item(i) for i in range(list_widget.count()) if not list_widget.item(i).isHidden()]

    # ── Filter menu ────────────────────────────────────────────────────

    def _show_filter_menu(self) -> None:
        from ..view_models.batch_list_vm import FILTER_OPTIONS

        menu = QMenu(self)
        group = QActionGroup(menu)
        group.setExclusive(True)
        for fkey, fname, _statuses in FILTER_OPTIONS:
            action = menu.addAction(_filter_option_label(fkey, fname))
            action.setCheckable(True)
            action.setChecked(fkey == self._vm.active_filter)
            action.setData(fkey)
            group.addAction(action)
        chosen = menu.exec(self.filter_button.mapToGlobal(self.filter_button.rect().bottomLeft()))
        if chosen is not None:
            self._vm.set_status_filter(str(chosen.data()))

    # ── Sort menu ──────────────────────────────────────────────────────

    def _show_sort_menu(self) -> None:
        menu = QMenu(self)
        sort_group = QActionGroup(menu)
        sort_group.setExclusive(True)
        sort_options = [
            ("custom", _sort_option_label("custom")),
            ("name", _sort_option_label("name")),
            ("type", _sort_option_label("type")),
            ("size", _sort_option_label("size")),
            ("mtime", _sort_option_label("mtime")),
        ]
        for key, label in sort_options:
            action = menu.addAction(label)
            action.setCheckable(True)
            action.setChecked(key == self._vm.sort_key)
            action.setData(key)
            sort_group.addAction(action)

        menu.addSeparator()

        dir_group = QActionGroup(menu)
        dir_group.setExclusive(True)
        asc_action = menu.addAction(_t("components.file_drop.batch_list.sort_ascending", "Ascending"))
        asc_action.setCheckable(True)
        asc_action.setChecked(self._vm.sort_ascending)
        asc_action.setData(True)
        dir_group.addAction(asc_action)

        desc_action = menu.addAction(_t("components.file_drop.batch_list.sort_descending", "Descending"))
        desc_action.setCheckable(True)
        desc_action.setChecked(not self._vm.sort_ascending)
        desc_action.setData(False)
        dir_group.addAction(desc_action)

        chosen = menu.exec(self.sort_button.mapToGlobal(self.sort_button.rect().bottomLeft()))
        if chosen is None:
            return
        if chosen in dir_group.actions():
            self._vm.set_sort_state(self._vm.sort_key, bool(chosen.data()))
        elif chosen in sort_group.actions():
            self._vm.set_sort_state(str(chosen.data()), self._vm.sort_ascending)

    # ── Sort application ───────────────────────────────────────────────

    def _apply_sort_for_category(self, category: str) -> None:
        list_widget = self._tabs.get(category)
        if list_widget is None or list_widget.count() <= 1:
            return
        ordered = self._vm._get_ordered_paths_for_category(category)
        if not ordered:
            return

        # Collect items and widgets
        items_by_path: dict[str, QListWidgetItem] = {}
        widgets_by_path: dict[str, QWidget | None] = {}
        current_item = list_widget.currentItem()
        current_path = ""
        if current_item is not None:
            fp = current_item.data(Qt.ItemDataRole.UserRole)
            if isinstance(fp, str):
                current_path = fp

        for index in range(list_widget.count()):
            item = list_widget.item(index)
            fp = item.data(Qt.ItemDataRole.UserRole)
            if not isinstance(fp, str):
                continue
            items_by_path[fp] = item
            widget = list_widget.itemWidget(item)
            if widget is not None:
                list_widget.removeItemWidget(item)
                widget.setParent(None)
            widgets_by_path[fp] = widget

        for _ in range(list_widget.count()):
            list_widget.takeItem(0)

        for idx, file_path in enumerate(ordered):
            item = items_by_path.get(file_path)
            if item is None:
                continue
            list_widget.insertItem(idx, item)
            widget = widgets_by_path.get(file_path)
            if widget is not None:
                list_widget.setItemWidget(item, widget)

        if current_path and current_path in items_by_path:
            list_widget.setCurrentItem(items_by_path[current_path])
        self._refresh_sequence_numbers(list_widget)

    # ── Move / reorder ─────────────────────────────────────────────────

    def _move_current_item_up(self) -> bool:
        return self._move_current_item(-1)

    def _move_current_item_down(self) -> bool:
        return self._move_current_item(1)

    def _move_current_item(self, offset: int) -> bool:
        category = self._vm.current_category
        if category not in self._tabs:
            self._update_reorder_buttons()
            return False
        list_widget = self._tabs[category]
        moved = list_widget.move_current_item_by(offset)
        if moved:
            # Record new order
            ordered = []
            for i in range(list_widget.count()):
                item = list_widget.item(i)
                fp = item.data(Qt.ItemDataRole.UserRole)
                if isinstance(fp, str):
                    ordered.append(fp)
            self._vm.reorder_manual(category, ordered)
            self._focus_current_list()
        self._update_reorder_buttons()
        return moved

    def _handle_manual_reorder(self) -> None:
        """Called when Ctrl+Up/Down reorders items."""
        sender = self.sender()
        if isinstance(sender, QListWidget):
            category = str(sender.property("batchCategory") or "")
            if category:
                ordered = []
                for i in range(sender.count()):
                    item = sender.item(i)
                    fp = item.data(Qt.ItemDataRole.UserRole)
                    if isinstance(fp, str):
                        ordered.append(fp)
                self._vm.reorder_manual(category, ordered)
            self._refresh_sequence_numbers(sender)
        self._handle_selection_changed()

    # ── Context menu ───────────────────────────────────────────────────

    def _show_item_context_menu(self, category: str, position) -> None:
        list_widget = self._tabs[category]
        item = list_widget.itemAt(position)
        if item is None:
            return
        fp = item.data(Qt.ItemDataRole.UserRole)
        file_path = fp if isinstance(fp, str) else None
        if file_path and item not in list_widget.selectedItems():
            list_widget.setCurrentItem(item)
            item.setSelected(True)

        menu = QMenu(self)

        # Retry actions
        selected_failed, category_failed = self._vm.build_retry_targets(category, file_path)
        retry_selected = menu.addAction(
            _t("components.file_drop.batch_list.retry_selected_failed", "Retry Selected Failed")
        )
        retry_selected.setEnabled(bool(selected_failed))
        retry_selected.triggered.connect(lambda: self._vm.reset_failed_files(selected_failed))

        retry_all = menu.addAction(_t("components.file_drop.batch_list.retry_all_failed", "Retry All Failed"))
        retry_all.setEnabled(bool(category_failed))
        retry_all.triggered.connect(lambda: self._vm.reset_failed_files(category_failed))

        # Error details for failed entries
        if file_path:
            entry = self._vm.get_file_entry(file_path)
            if entry is not None and entry.status == "failed" and entry.error_message:
                menu.addSeparator()
                _action_view = menu.addAction(
                    _t("components.file_drop.batch_list.action_view_error", "Show Error Details"),
                )
                _action_view.triggered.connect(
                    lambda: self.entry_action_requested.emit("show_error_details", file_path),
                )
                _action_copy = menu.addAction(
                    _t("editors.common.copy", "Copy Error"),
                )
                _action_copy.triggered.connect(
                    lambda: self.entry_action_requested.emit("copy_error_details", file_path),
                )

        # Multi-selection actions
        selected = self.get_selected_files(category)
        if len(selected) > 1:
            menu.addSeparator()
            _action_remove = menu.addAction(
                _t(
                    "components.file_drop.batch_list.action_remove_selected",
                    "Remove Selected ({count})",
                    count=len(selected),
                ),
            )
            _action_remove.triggered.connect(lambda paths=selected: self._remove_selected(paths))
            _action_open = menu.addAction(
                _t(
                    "components.file_drop.batch_list.action_open_selected_locations",
                    "Open Selected Locations ({count})",
                    count=len(selected),
                ),
            )
            _action_open.triggered.connect(lambda paths=selected: self._open_selected_locations(paths))

        menu.exec(list_widget.mapToGlobal(position))

    def _remove_selected(self, paths: list[str]) -> None:
        for path in paths:
            self._vm.remove_file(path)

    def _open_selected_locations(self, paths: list[str]) -> None:
        for path in paths:
            self.entry_action_requested.emit("open_source_location", path)

    # ── Public helpers ───────────────────────────────────────────────────

    def select_file(self, file_path: str) -> bool:
        """Activate and select a file entry by path.

        Returns ``True`` when the entry exists and selection changed.
        """
        target = _path_compare_key(str(file_path))
        for category, list_widget in self._tabs.items():
            for index in range(list_widget.count()):
                item = list_widget.item(index)
                stored = item.data(Qt.ItemDataRole.UserRole)
                if not isinstance(stored, str) or _path_compare_key(stored) != target:
                    continue
                self._suspend_tab_selection_sync = True
                try:
                    self._activate_tab(category)
                finally:
                    self._suspend_tab_selection_sync = False
                self._suspend_selection_sync = True
                try:
                    list_widget.clearSelection()
                    item.setSelected(True)
                    list_widget.setCurrentItem(item)
                    list_widget.scrollToItem(item)
                finally:
                    self._suspend_selection_sync = False
                self._handle_selection_changed()
                return True
        return False

    def copy_error_details(self, file_path: str) -> bool:
        """Copy the error details for a specific entry to the clipboard."""
        target = str(file_path)
        for list_widget in self._tabs.values():
            for index in range(list_widget.count()):
                item = list_widget.item(index)
                if item.data(Qt.ItemDataRole.UserRole) != target:
                    continue
                widget = list_widget.itemWidget(item)
                if isinstance(widget, BatchEntryItemWidget):
                    return widget.copy_error_details()
                return False
        return False

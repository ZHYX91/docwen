"""InfoArea widget — status bar with history, transient, and task summary.

Renders the InfoArea UI based on ``InfoAreaViewModel`` state.
Does NOT call runtime/plugins directly — all actions go through the ViewModel.

Widget structure:
  content_card (SimpleCardWidget)
    +-- scroll (QScrollArea, horizontal scrollbar always off)
    |     +-- msg_container (history message rows)
    +-- divider (QFrame HLine)
    +-- status_section
          +-- status_meta_label (overview source + activity dots)
          +-- status_summary_label (InteractiveTextLabel, clickable)
          +-- status_guide_row (task completion guide buttons, hidden by default)
"""

from __future__ import annotations

import logging
from typing import TYPE_CHECKING
from typing import cast as _cast

from PySide6.QtCore import QEvent, QSize, Qt, QTimer
from PySide6.QtGui import QAction
from PySide6.QtWidgets import (
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QMenu,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QToolButton,
    QVBoxLayout,
    QWidget,
)

from docwen_gui.i18n import t as _t

if TYPE_CHECKING:
    from ..view_models.info_area_vm import InfoAreaViewModel

logger = logging.getLogger(__name__)

# ── Design constants ──────────────────────────────────────────────────────
_SPACING_XS = 4
_SPACING_SM = 8
_SPACING_MD = 12
_TIMESTAMP_WIDTH = 54  # logical px, will be DPI-scaled
_LOCATION_BUTTON_SIZE = 26
_LOCATION_ICON_SIZE = 16
_SCROLL_DELAY_MS = 50


def _scale(value: int) -> int:
    """DPI-scale a logical pixel value."""
    try:
        from PySide6.QtWidgets import QApplication

        app = QApplication.instance()
        if app is not None:
            screen = _cast(QApplication, app).primaryScreen()
            if screen is not None:
                dpi = screen.logicalDotsPerInch()
                return round(value * dpi / 96.0)
    except (AttributeError, RuntimeError):
        pass
    return value


# ── Helper widgets ────────────────────────────────────────────────────────


class _StatusLocationButton(QToolButton):
    """Lightweight location button with folder icon.

    Left-click: emits path_clicked signal.
    Right-click: shows context menu with "Copy path" action.
    """

    def __init__(
        self,
        file_path: str,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self._file_path = file_path

        self.setObjectName("statusLocationButton")
        self.setToolTip(_t("info_area.open_location", "Open {path}", path=file_path))
        self.setAccessibleName(self.toolTip())
        self.setAccessibleDescription(file_path)
        self.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonIconOnly)

        # Use standard folder icon — no deep import of old gui icon_utils.
        self.setIcon(self.style().standardIcon(self.style().StandardPixmap.SP_DirOpenIcon))
        icon_size = _scale(_LOCATION_ICON_SIZE)
        self.setIconSize(QSize(icon_size, icon_size))
        btn_size = _scale(_LOCATION_BUTTON_SIZE)
        self.setFixedSize(btn_size, btn_size)
        self.setAutoRaise(True)
        self.setCursor(Qt.CursorShape.PointingHandCursor)

    def contextMenuEvent(self, event) -> None:  # type: ignore[override]
        """Show right-click context menu with copy path action."""
        menu = self._create_context_menu()
        menu.exec(event.globalPos())
        event.accept()

    def _create_context_menu(self) -> QMenu:
        menu = QMenu(self)
        copy_action = QAction(_t("info_area.copy_path", "Copy Path"), menu)
        copy_action.triggered.connect(self._copy_path_to_clipboard)
        menu.addAction(copy_action)
        return menu

    def _copy_path_to_clipboard(self) -> None:
        """Copy the file path to clipboard."""
        try:
            from PySide6.QtWidgets import QApplication

            QApplication.clipboard().setText(self._file_path)
        except Exception:
            pass


# ── Main widget ───────────────────────────────────────────────────────────


class InfoArea(QWidget):
    """Scrollable status bar component.

    Renders history messages, transient status, task summary, and
    task completion guide buttons.  All state is driven by the
    ``InfoAreaViewModel`` — this widget is a pure renderer.
    """

    def __init__(
        self,
        view_model: InfoAreaViewModel,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self._vm = view_model
        self.setObjectName("infoArea")

        # Widget refs
        self._content_card: QWidget = _cast(QWidget, None)
        self._scroll: QScrollArea = _cast(QScrollArea, None)
        self._msg_container: QWidget = _cast(QWidget, None)
        self._msg_layout: QVBoxLayout = _cast(QVBoxLayout, None)
        self._empty_history_state: QWidget = _cast(QWidget, None)
        self._history_row_widgets: list[QWidget] = []
        self._status_meta_label: QLabel = _cast(QLabel, None)
        self._status_summary_label: QLabel = _cast(QLabel, None)
        self._status_guide_row: QWidget = _cast(QWidget, None)
        self._status_guide_actions_widget: QWidget = _cast(QWidget, None)
        self._status_guide_actions_layout: QGridLayout = _cast(QGridLayout, None)
        self._status_guide_buttons: list[QPushButton] = []
        self._guide_layout_timer = QTimer(self)
        self._guide_layout_timer.setSingleShot(True)
        self._guide_layout_timer.timeout.connect(self._sync_guide_button_layout)

        self._build_ui()
        self._wire_vm()

    # ── UI Construction ──────────────────────────────────────────────────

    def _build_ui(self) -> None:
        """Build the InfoArea widget skeleton."""
        root_layout = QVBoxLayout(self)
        root_layout.setContentsMargins(0, _SPACING_XS, 0, _SPACING_XS)
        root_layout.setSpacing(_SPACING_SM)

        # Content card
        self._content_card = QWidget(self)
        self._content_card.setObjectName("infoAreaContentCard")
        card_layout = QVBoxLayout(self._content_card)
        card_layout.setContentsMargins(_SPACING_SM, _SPACING_SM, _SPACING_SM, _SPACING_SM)
        card_layout.setSpacing(_SPACING_SM)

        # Scroll area for history messages
        self._scroll = QScrollArea()
        self._scroll.setObjectName("infoHistoryScrollArea")
        self._scroll.setWidgetResizable(True)
        self._scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self._scroll.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Ignored)

        self._msg_container = QWidget()
        self._msg_layout = QVBoxLayout(self._msg_container)
        self._msg_layout.setContentsMargins(0, 0, 0, 0)
        self._msg_layout.setSpacing(_SPACING_SM)
        self._msg_layout.setAlignment(Qt.AlignmentFlag.AlignTop)

        self._empty_history_state = QWidget(self._msg_container)
        self._empty_history_state.setObjectName("infoHistoryEmptyState")
        empty_layout = QVBoxLayout(self._empty_history_state)
        empty_layout.setContentsMargins(_SPACING_MD, _SPACING_MD, _SPACING_MD, _SPACING_MD)
        empty_layout.setSpacing(_SPACING_XS)
        empty_layout.addStretch(1)

        empty_title = QLabel(_t("info_area.history_title", "Recent messages"), self._empty_history_state)
        empty_title.setObjectName("infoHistoryEmptyTitle")
        empty_title.setAlignment(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter)
        empty_layout.addWidget(empty_title)

        empty_caption = QLabel(_t("common.ready", "Ready"), self._empty_history_state)
        empty_caption.setObjectName("infoHistoryEmptyCaption")
        empty_caption.setAlignment(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter)
        empty_layout.addWidget(empty_caption)
        empty_layout.addStretch(1)

        self._msg_layout.addWidget(self._empty_history_state, stretch=1)

        self._scroll.setWidget(self._msg_container)
        card_layout.addWidget(self._scroll, stretch=1)

        # Divider
        divider = QFrame()
        divider.setObjectName("infoSectionDivider")
        divider.setFrameShape(QFrame.Shape.HLine)
        card_layout.addWidget(divider)

        # Status section
        status_section = QWidget()
        status_section.setProperty("infoStatusSource", "idle")
        status_section.setProperty("infoStatusTone", "secondary")
        overview_layout = QVBoxLayout(status_section)
        overview_layout.setContentsMargins(0, 0, 0, 0)
        overview_layout.setSpacing(_SPACING_XS)

        # Meta label (e.g. "Task active...", "History (3)")
        self._status_meta_label = QLabel("", status_section)
        self._status_meta_label.setObjectName("infoStatusMeta")
        overview_layout.addWidget(self._status_meta_label)

        # Summary label (main status text, potentially interactive)
        self._status_summary_label = QLabel("", status_section)
        self._status_summary_label.setObjectName("infoStatusSummary")
        self._status_summary_label.setWordWrap(True)
        overview_layout.addWidget(self._status_summary_label)

        # Guide row (hidden by default)
        self._status_guide_row = QWidget(status_section)
        self._status_guide_row.setObjectName("infoStatusGuideRow")
        guide_outer_layout = QHBoxLayout(self._status_guide_row)
        guide_outer_layout.setContentsMargins(0, 0, 0, 0)
        guide_outer_layout.setSpacing(_SPACING_XS)

        guide_actions_widget = QWidget(self._status_guide_row)
        guide_actions_widget.setObjectName("infoStatusGuideActions")
        self._status_guide_actions_widget = guide_actions_widget
        guide_actions_widget.installEventFilter(self)
        self._status_guide_actions_layout = QGridLayout(guide_actions_widget)
        self._status_guide_actions_layout.setContentsMargins(0, 0, 0, 0)
        self._status_guide_actions_layout.setSpacing(_SPACING_XS)
        guide_outer_layout.addWidget(guide_actions_widget, stretch=1)

        self._status_guide_row.setVisible(False)
        overview_layout.addWidget(self._status_guide_row)
        card_layout.addWidget(status_section)

        root_layout.addWidget(self._content_card)

    # ── ViewModel Wiring ─────────────────────────────────────────────────

    def _wire_vm(self) -> None:
        """Connect ViewModel signals to widget updates."""
        self._vm.state_changed.connect(self._sync_from_vm)
        self._sync_from_vm()

    def _sync_from_vm(self) -> None:
        """Rebuild the widget from current ViewModel state."""
        self._rebuild_history()
        self._update_status_section()
        self._update_guide_row()

    # ── History rendering ────────────────────────────────────────────────

    def _rebuild_history(self) -> None:
        """Rebuild all history message rows from ViewModel state."""
        if self._msg_layout is None:
            return

        for row in self._history_row_widgets:
            self._msg_layout.removeWidget(row)
            row.hide()
            row.setParent(None)
            row.deleteLater()
        self._history_row_widgets.clear()

        has_history = bool(self._vm.history_rows)
        self._empty_history_state.setVisible(not has_history)

        # Add rows from ViewModel
        for row_data in self._vm.history_rows:
            row_widget = self._build_history_row(row_data)
            self._msg_layout.addWidget(row_widget)
            self._history_row_widgets.append(row_widget)

        # Schedule scroll to bottom
        QTimer.singleShot(_SCROLL_DELAY_MS, self._scroll_to_bottom)

    def _build_history_row(self, row_data) -> QWidget:
        """Build a single history message row widget."""

        has_location = bool(row_data.show_location and row_data.file_path)
        row = QWidget()
        row.setObjectName("infoHistoryRow")
        row.setProperty("infoStatusTone", row_data.message_type)
        row.setProperty(
            "hasLocationAction",
            "true" if has_location else "false",
        )
        row.setProperty("hasNavigationTarget", "false")
        row.setProperty("hasOperationId", "false")
        row_layout = QHBoxLayout(row)
        row_layout.setContentsMargins(0, 2, 0, 2)
        row_layout.setSpacing(_SPACING_XS)

        # Content wrapper
        content = QWidget(row)
        content_layout = QVBoxLayout(content)
        content_layout.setContentsMargins(0, 0, 0, 0)
        content_layout.setSpacing(2)

        # First line: timestamp + message
        first_line = QWidget(content)
        first_line.setObjectName("infoHistoryMeta")
        first_line.setProperty("hasBadge", "false")
        first_line_layout = QHBoxLayout(first_line)
        first_line_layout.setContentsMargins(0, 0, 0, 0)
        first_line_layout.setSpacing(_SPACING_XS)

        # Timestamp label (HH:MM:SS, fixed width)
        timestamp_width = _scale(_TIMESTAMP_WIDTH)
        timestamp_label = QLabel(row_data.timestamp, first_line)
        timestamp_label.setObjectName("statusTimestamp")
        timestamp_label.setMinimumWidth(timestamp_width)
        timestamp_label.setMaximumWidth(timestamp_width)
        first_line_layout.addWidget(timestamp_label, alignment=Qt.AlignmentFlag.AlignTop)

        # Message text (selectable)
        msg_label = QLabel(row_data.message, first_line)
        msg_label.setObjectName("infoHistoryText")
        msg_label.setProperty("infoStatusTone", row_data.message_type)
        msg_label.setWordWrap(True)
        msg_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        msg_label.setToolTip(row_data.message)
        row.setToolTip(row_data.message)
        first_line_layout.addWidget(msg_label, stretch=1)
        content_layout.addWidget(first_line)

        row_layout.addWidget(content, stretch=1)

        # Location button
        if row_data.show_location and row_data.file_path:
            loc_btn = _StatusLocationButton(row_data.file_path, row)
            loc_btn.clicked.connect(lambda checked=False, p=row_data.file_path: self._vm.request_location(p))
            row_layout.addWidget(loc_btn, alignment=Qt.AlignmentFlag.AlignTop)

        return row

    # ── Status section ───────────────────────────────────────────────────

    def _update_status_section(self) -> None:
        """Update the status meta, summary, and styling from ViewModel."""
        vm = self._vm

        if self._status_meta_label is not None:
            if vm.activity_enabled:
                self._status_meta_label.setText(vm.activity_meta_text)
            else:
                self._status_meta_label.setText(vm.status_meta_text)
            self._status_meta_label.setToolTip(vm.status_meta_text)

        if self._status_summary_label is not None:
            self._status_summary_label.setText(vm.status_summary_text)
            tooltip = vm.status_summary_text
            if vm.status_action_target:
                tooltip = f"{tooltip}\n{vm.status_action_target}"
            self._status_summary_label.setToolTip(tooltip)

            # Interactive styling
            has_target = bool(vm.status_action_target)
            self._status_summary_label.setProperty("interactiveText", has_target)
            if has_target:
                self._status_summary_label.setCursor(Qt.CursorShape.PointingHandCursor)
                self._status_summary_label.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
            else:
                self._status_summary_label.setCursor(Qt.CursorShape.ArrowCursor)
                self._status_summary_label.setFocusPolicy(Qt.FocusPolicy.NoFocus)

        # Update status section properties
        self._status_summary_label.parent() if self._status_summary_label else None
        # Walk up to find the status section (parent of summary label)
        if self._status_summary_label is not None:
            p = self._status_summary_label.parent()
            if p is not None:
                p.setProperty("infoStatusSource", vm.status_source)
                p.setProperty("infoStatusTone", vm.status_tone)
                self._refresh_widget_style(_cast(QWidget, p))

    # ── Guide row ────────────────────────────────────────────────────────

    def _update_guide_row(self) -> None:
        """Update the guide button row from ViewModel state."""
        if self._status_guide_row is None:
            return
        if self._status_guide_actions_layout is None:
            return

        # Clear existing buttons
        self._clear_layout(self._status_guide_actions_layout)
        self._status_guide_buttons.clear()

        vm = self._vm
        if not vm.guide_visible:
            self._status_guide_row.setVisible(False)
            return

        guide_actions = vm.guide_actions
        if not guide_actions:
            self._status_guide_row.setVisible(False)
            return

        for index, action in enumerate(guide_actions):
            action_key = action.get("action_key", "")
            target_path = action.get("target_path", "")
            label = vm.guide_action_label(action_key)

            button = QPushButton(label, self._status_guide_row)
            button.setObjectName("infoStatusGuideButton")
            button.setProperty(
                "guideActionPriority",
                "primary" if index == 0 else "secondary",
            )
            button.setAccessibleName(label)
            button.setAccessibleDescription(target_path or label)
            button.setMinimumHeight(_scale(32))

            # Capture action_key and target_path for the lambda
            button.clicked.connect(lambda checked=False, ak=action_key, tp=target_path: vm.request_guide_action(ak, tp))
            self._status_guide_buttons.append(button)

        self._status_guide_row.setVisible(True)
        self._sync_guide_button_layout()

    def eventFilter(self, watched, event) -> bool:
        if watched is self._status_guide_actions_widget and event.type() in {
            QEvent.Type.Resize,
            QEvent.Type.LayoutRequest,
        }:
            self._guide_layout_timer.start(0)
        return super().eventFilter(watched, event)

    def changeEvent(self, event) -> None:
        super().changeEvent(event)
        if event.type() in {QEvent.Type.FontChange, QEvent.Type.StyleChange}:
            self._guide_layout_timer.start(0)

    def _sync_guide_button_layout(self) -> None:
        """Wrap guide actions before their translated labels are squeezed."""
        buttons = [button for button in self._status_guide_buttons if button.parent() is not None]
        if not buttons:
            return

        layout = self._status_guide_actions_layout
        available_width = self._status_guide_actions_widget.contentsRect().width()
        if available_width <= 0:
            return
        spacing = max(0, layout.horizontalSpacing())
        widths = [button.sizeHint().width() for button in buttons]
        one_row_width = sum(widths) + spacing * (len(buttons) - 1)
        # Keep a small reserve for style/palette metrics that are not reflected
        # consistently in QPushButton.sizeHint() across native Windows themes.
        if one_row_width + (_SPACING_SM * 2) <= available_width:
            columns = len(buttons)
        elif len(buttons) > 1 and (max(widths) * 2 + spacing) <= available_width:
            columns = 2
        else:
            columns = 1

        for button in buttons:
            layout.removeWidget(button)
        for index, button in enumerate(buttons):
            layout.addWidget(button, index // columns, index % columns)
        for column in range(columns):
            layout.setColumnStretch(column, 1)
        self._status_guide_actions_widget.updateGeometry()

    # ── Helpers ──────────────────────────────────────────────────────────

    def _scroll_to_bottom(self) -> None:
        """Scroll the history area to the bottom."""
        if self._scroll is None:
            return
        sb = self._scroll.verticalScrollBar()
        sb.setValue(sb.maximum())

    @staticmethod
    def _clear_layout(layout: QGridLayout | QHBoxLayout | QVBoxLayout) -> None:
        """Remove and delete all widgets from a layout."""
        while layout.count() > 0:
            item = layout.takeAt(0)
            if item is None:
                continue
            widget = item.widget()
            if widget is not None:
                widget.hide()
                widget.setParent(None)
                widget.deleteLater()
            else:
                # Could be a layout item or spacer
                sub_layout = item.layout()
                if sub_layout is not None:
                    InfoArea._clear_layout(sub_layout)  # pyright: ignore[reportArgumentType]

    @staticmethod
    def _refresh_widget_style(widget: QWidget) -> None:
        """Force a widget to re-read its stylesheet properties."""
        style = widget.style()
        if style is not None:
            style.unpolish(widget)
            style.polish(widget)

    # ── Event handlers ───────────────────────────────────────────────────

    def mouseReleaseEvent(self, event) -> None:
        """Handle click on the status summary area."""
        if event.button() == Qt.MouseButton.LeftButton and self._status_summary_label is not None:
            # Check if click is on or near the summary label
            local_pos = self._status_summary_label.mapFromParent(event.position().toPoint())
            if self._status_summary_label.rect().contains(local_pos) and self._vm.status_action_target:
                self._vm.request_navigation()
                event.accept()
                return
        super().mouseReleaseEvent(event)

    # ── Public API ───────────────────────────────────────────────────────

    @property
    def view_model(self) -> InfoAreaViewModel:
        """The ViewModel driving this info area."""
        return self._vm

    @property
    def status_summary_text(self) -> str:
        """Current status summary text (for testing)."""
        return self._vm.status_summary_text

    @property
    def status_meta_text(self) -> str:
        """Current status meta text (for testing)."""
        return self._vm.status_meta_text

    @property
    def status_source(self) -> str:
        """Current status source (for testing)."""
        return self._vm.status_source

    @property
    def status_tone(self) -> str:
        """Current status tone (for testing)."""
        return self._vm.status_tone

    @property
    def is_guide_row_visible(self) -> bool:
        """Whether the guide row is visible (for testing)."""
        return self._status_guide_row is not None and not self._status_guide_row.isHidden()

    @property
    def message_count(self) -> int:
        """Number of history messages (for testing)."""
        return self._vm.message_count

    @property
    def message_types(self) -> list[str]:
        """List of message types (for testing)."""
        return self._vm.message_types

    def get_history_row_widget(self, index: int) -> QWidget | None:
        """Return the widget for a history row at the given index (for testing)."""
        if self._msg_layout is None:
            return None
        if index < 0 or index >= self._vm.message_count:
            return None
        return self._history_row_widgets[index]

    def find_guide_buttons(self) -> list[QPushButton]:
        """Find all guide buttons (for testing)."""
        if self._status_guide_row is None:
            return []
        return [
            btn
            for btn in self._status_guide_row.findChildren(QPushButton)
            if btn.objectName() == "infoStatusGuideButton"
        ]

    def find_location_buttons(self) -> list[QToolButton]:
        """Find all location buttons (for testing)."""
        return [btn for btn in self.findChildren(QToolButton) if btn.objectName() == "statusLocationButton"]


__all__ = ["InfoArea", "_StatusLocationButton"]

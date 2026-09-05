"""InfoArea widget — status bar with history, transient, and task summary.

Renders the InfoArea UI based on ``InfoAreaViewModel`` state.
Does NOT call runtime/plugins directly — all actions go through the ViewModel.

Current task feedback stays above an optional, bounded activity history.
"""

from __future__ import annotations

import logging
from typing import TYPE_CHECKING
from typing import cast as _cast

from PySide6.QtCore import QEvent, QSize, Qt, QTimer
from PySide6.QtGui import QAction, QFontDatabase
from PySide6.QtWidgets import (
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QMenu,
    QProgressBar,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QToolButton,
    QVBoxLayout,
    QWidget,
)

from docwen_gui.i18n import t as _t

from .panel_card import PanelCard, TaskActivityList

if TYPE_CHECKING:
    from ..view_models.info_area_vm import HistoryRowData, InfoAreaViewModel

logger = logging.getLogger(__name__)

# ── Design constants ──────────────────────────────────────────────────────
_SPACING_XS = 4
_SPACING_SM = 8
_LOCATION_BUTTON_SIZE = 26
_LOCATION_ICON_SIZE = 16
_SCROLL_DELAY_MS = 50


class _SummaryButton(QPushButton):
    """A native keyboard-accessible button with a wrapping caption."""

    def __init__(self, parent: QWidget) -> None:
        super().__init__(parent)
        self.caption = QLabel(self)
        self.caption.setWordWrap(True)
        self.caption.setTextFormat(Qt.TextFormat.PlainText)
        self.caption.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)
        layout = QVBoxLayout(self)
        self._caption_layout = layout
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self.caption)
        self.setFlat(True)
        self.setAutoDefault(False)
        policy = QSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Preferred)
        policy.setHeightForWidth(True)
        self.setSizePolicy(policy)

    def setText(self, text: str) -> None:
        self.caption.setText(text)
        self.setAccessibleName(text)
        self._caption_layout.invalidate()
        self.updateGeometry()

    def text(self) -> str:
        return self.caption.text()

    def sizeHint(self) -> QSize:
        return self._caption_layout.sizeHint()

    def minimumSizeHint(self) -> QSize:
        return self._caption_layout.minimumSize()

    def heightForWidth(self, width: int) -> int:
        return self.caption.heightForWidth(width)

    def keyPressEvent(self, event) -> None:
        if event.key() in {Qt.Key.Key_Return, Qt.Key.Key_Enter}:
            self.click()
            event.accept()
            return
        super().keyPressEvent(event)


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
        self._history_row_widgets: list[QWidget] = []
        self._rendered_history_rows: list[HistoryRowData] = []
        self._history_expanded = False
        self._rendered_guide_actions: list[dict[str, str]] | None = None
        self._status_meta_label: QLabel = _cast(QLabel, None)
        self._status_summary_label: _SummaryButton = _cast(_SummaryButton, None)
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
        self._content_card = PanelCard(parent=self)
        self._content_card.setObjectName("infoAreaContentCard")
        card_layout = self._content_card.content_layout
        card_layout.setSpacing(_SPACING_SM)

        # Scroll area for history messages
        self._scroll = QScrollArea()
        self._scroll.setObjectName("infoHistoryScrollArea")
        self._scroll.setWidgetResizable(True)
        self._scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self._scroll.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)

        self._msg_container = TaskActivityList()
        self._msg_layout = self._msg_container.content_layout

        self._scroll.setWidget(self._msg_container)

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
        self._status_summary_label = _SummaryButton(status_section)
        self._status_summary_label.setObjectName("infoStatusSummary")
        self._status_summary_label.clicked.connect(self._vm.request_navigation)
        overview_layout.addWidget(self._status_summary_label)
        self._progress = QProgressBar(status_section)
        self._progress.setObjectName("infoTaskProgress")
        self._progress.setTextVisible(False)
        self._progress.setFixedHeight(5)
        self._progress.hide()
        overview_layout.addWidget(self._progress)

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

        self._notice = QLabel(self)
        self._notice.setObjectName("infoNotification")
        self._notice.setWordWrap(True)
        self._notice.setTextFormat(Qt.TextFormat.PlainText)
        self._notice.hide()
        card_layout.addWidget(self._notice)
        self._history_toolbar = QWidget(self)
        self._history_toolbar.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Fixed)
        toolbar = QHBoxLayout(self._history_toolbar)
        toolbar.setContentsMargins(0, 0, 0, 0)
        self._history_toggle = QToolButton(self._history_toolbar)
        self._history_toggle.setCheckable(True)
        self._history_toggle.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextBesideIcon)
        self._history_toggle.setAutoRaise(True)
        self._history_toggle.toggled.connect(self._set_history_expanded)
        toolbar.addWidget(self._history_toggle)
        toolbar.addStretch(1)
        clear = QToolButton(self._history_toolbar)
        clear.setText(_t("components.file_drop.clear_button", "Clear"))
        clear.setToolTip(_t("info_area.clear_history", "Clear activity records"))
        clear.setAccessibleName(clear.toolTip())
        clear.clicked.connect(self._vm.clear_history)
        toolbar.addWidget(clear)
        card_layout.addWidget(self._history_toolbar)
        card_layout.addWidget(self._scroll)
        self._scroll.setFixedHeight(180)
        self._scroll.hide()

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

    def _set_history_expanded(self, expanded: bool) -> None:
        self._history_expanded = expanded
        self._history_toggle.setArrowType(Qt.ArrowType.DownArrow if expanded else Qt.ArrowType.RightArrow)
        self._scroll.setVisible(expanded and bool(self._vm.history_rows))
        self.updateGeometry()

    # ── History rendering ────────────────────────────────────────────────

    def _rebuild_history(self) -> None:
        """Rebuild all history message rows from ViewModel state."""
        if self._msg_layout is None:
            return

        rows = self._vm.history_rows
        if rows == self._rendered_history_rows:
            self._history_toolbar.setVisible(bool(rows))
            return
        scrollbar = self._scroll.verticalScrollBar()
        follow_tail = scrollbar.value() >= self._history_tail_position() - 2
        old_position = scrollbar.value()
        available = list(zip(self._rendered_history_rows, self._history_row_widgets, strict=True))
        widgets = []
        for data in rows:
            match = next((index for index, (old, _) in enumerate(available) if old is data), None)
            if match is None:
                match = next(
                    (
                        index
                        for index, (old, _) in enumerate(available)
                        if old.message == data.message
                        and old.operation_id == data.operation_id
                        and old.file_path == data.file_path
                        and old.message_type == data.message_type
                        and old.navigate_file_path == data.navigate_file_path
                        and old.show_location == data.show_location
                    ),
                    None,
                )
            if match is None:
                widget = self._build_history_row(data)
            else:
                old, widget = available.pop(match)
                if old != data:
                    timestamp = widget.findChild(QLabel, "statusTimestamp")
                    badge = widget.findChild(QLabel, "infoHistoryRepeatBadge")
                    if timestamp is not None:
                        timestamp.setText(data.timestamp)
                    if badge is not None:
                        badge.setText(
                            _t("info_area.history_repeated", "Repeated {count} times", count=data.repeat_count)
                        )
                        badge.setToolTip(badge.text())
                        badge.setVisible(data.repeat_count > 1)
            widgets.append(widget)
        for _, widget in available:
            self._msg_layout.removeWidget(widget)
            widget.hide()
            widget.setParent(None)
            widget.deleteLater()
        for widget in widgets:
            self._msg_layout.removeWidget(widget)
            self._msg_layout.addWidget(widget)
        self._history_row_widgets = widgets
        self._rendered_history_rows = rows
        self._history_toolbar.setVisible(bool(rows))
        self._history_toggle.setText(_t("info_area.activity_count", "Activity ({count})", count=len(rows)))
        self._set_history_expanded(self._history_expanded)
        if follow_tail:
            QTimer.singleShot(_SCROLL_DELAY_MS, self._scroll_to_bottom)
        else:
            scrollbar.setValue(old_position)

    def _build_history_row(self, row_data: HistoryRowData) -> QWidget:
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
        row_layout.setContentsMargins(_SPACING_SM, _SPACING_XS, _SPACING_XS, _SPACING_XS)
        row_layout.setSpacing(_SPACING_SM)

        tone_marker = QFrame(row)
        tone_marker.setObjectName("infoHistoryToneMarker")
        tone_marker.setProperty("infoStatusTone", row_data.message_type)
        tone_marker.setFixedWidth(_scale(3))
        row_layout.addWidget(tone_marker)

        # Content wrapper
        content = QWidget(row)
        content_layout = QVBoxLayout(content)
        content_layout.setContentsMargins(0, 0, 0, 0)
        content_layout.setSpacing(2)

        # First line: timestamp + message
        first_line = QWidget(content)
        first_line.setObjectName("infoHistoryMeta")
        first_line.setProperty("hasBadge", "true" if row_data.repeat_count > 1 else "false")
        first_line_layout = QHBoxLayout(first_line)
        first_line_layout.setContentsMargins(0, 0, 0, 0)
        first_line_layout.setSpacing(_SPACING_XS)

        # Size from actual font metrics so every second remains visible.
        timestamp_label = QLabel(row_data.timestamp, first_line)
        timestamp_label.setObjectName("statusTimestamp")
        timestamp_label.setFont(QFontDatabase.systemFont(QFontDatabase.SystemFont.FixedFont))
        # Global typography is stylesheet-driven.  Polish before measuring so
        # an xlarge preset cannot enlarge the glyphs after the fixed width was
        # chosen and clip the final second digit.
        timestamp_label.ensurePolished()
        timestamp_width = timestamp_label.fontMetrics().horizontalAdvance("00:00:00")
        timestamp_label.setFixedWidth(timestamp_width + _scale(8))
        first_line_layout.addWidget(timestamp_label, alignment=Qt.AlignmentFlag.AlignTop)

        # Message text (selectable)
        msg_label = QLabel(row_data.message, first_line)
        msg_label.setTextFormat(Qt.TextFormat.PlainText)
        msg_label.setObjectName("infoHistoryText")
        msg_label.setProperty("infoStatusTone", row_data.message_type)
        msg_label.setWordWrap(True)
        msg_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        msg_label.setToolTip(row_data.message)
        row.setToolTip(row_data.message)
        first_line_layout.addStretch(1)
        repeat_label = QLabel(
            _t(
                "info_area.history_repeated",
                "Repeated {count} times",
                count=row_data.repeat_count,
            ),
            first_line,
        )
        repeat_label.setObjectName("infoHistoryRepeatBadge")
        repeat_label.setToolTip(repeat_label.text())
        repeat_label.setVisible(row_data.repeat_count > 1)
        first_line_layout.addWidget(repeat_label, alignment=Qt.AlignmentFlag.AlignTop)
        content_layout.addWidget(first_line)
        content_layout.addWidget(msg_label)

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
            self._status_summary_label.setEnabled(has_target)
            self._status_summary_label.setAccessibleDescription(vm.status_action_target)
            self._status_summary_label.setProperty("interactiveText", has_target)
            if has_target:
                self._status_summary_label.setCursor(Qt.CursorShape.PointingHandCursor)
                self._status_summary_label.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
            else:
                self._status_summary_label.setCursor(Qt.CursorShape.ArrowCursor)
                self._status_summary_label.setFocusPolicy(Qt.FocusPolicy.NoFocus)

        self._status_meta_label.setVisible(bool(vm.status_meta_text))
        self._notice.setText(vm.notification_text)
        self._notice.setVisible(bool(vm.notification_text))
        task = vm.task_summary
        self._progress.setVisible(task.state in {"active", "cancelling"})
        if task.percent is None:
            self._progress.setRange(0, 0)
        else:
            self._progress.setRange(0, 100)
            self._progress.setValue(round(task.percent))

        # Update status section properties
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
        actions = self._vm.guide_actions if self._vm.guide_visible else []
        if actions == self._rendered_guide_actions:
            return
        self._rendered_guide_actions = actions
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

    def _history_tail_position(self) -> int:
        """The last message, excluding spare space from Qt's size hints."""
        if not self._history_row_widgets:
            return 0
        content_bottom = self._history_row_widgets[-1].geometry().bottom() + 1
        return max(
            0, min(self._scroll.verticalScrollBar().maximum(), content_bottom - self._scroll.viewport().height())
        )

    def _scroll_to_bottom(self) -> None:
        """Follow actual messages even when wrapped size hints reserve extra space."""
        if self._scroll is None:
            return
        sb = self._scroll.verticalScrollBar()
        sb.setValue(self._history_tail_position())

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

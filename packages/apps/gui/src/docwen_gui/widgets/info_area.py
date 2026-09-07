"""InfoArea widget — status bar with history, transient, and task summary.

Renders the InfoArea UI based on ``InfoAreaViewModel`` state.
Does NOT call runtime/plugins directly — all actions go through the ViewModel.

Current task feedback links to one searchable activity window.
"""

from __future__ import annotations

import logging
from typing import TYPE_CHECKING
from typing import cast as _cast

from PySide6.QtCore import QEvent, QSize, Qt, QTimer
from PySide6.QtWidgets import (
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QProgressBar,
    QPushButton,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)

from docwen_gui.i18n import t as _t
from docwen_gui.resources import load_svg_icon
from docwen_gui.styles.design_tokens import Sizing, Spacing

from .elided_label import MiddleElidedLabel
from .output_file_row import OutputFileRow
from .panel_card import PanelCard

if TYPE_CHECKING:
    from ..view_models.info_area_vm import InfoAreaViewModel

logger = logging.getLogger(__name__)

# ── Design constants ──────────────────────────────────────────────────────
_SPACING_XS = Spacing.XS
_SPACING_SM = Spacing.SM


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


class _ActivityButton(QPushButton):
    def keyPressEvent(self, event) -> None:
        if event.key() in (Qt.Key.Key_Return, Qt.Key.Key_Enter):
            if not event.isAutoRepeat():
                self.click()
            event.accept()
            return
        super().keyPressEvent(event)


class InfoArea(QWidget):
    """Scrollable status bar component.

    Renders transient status, task summary, activity navigation, and
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
        self._content_card: PanelCard = _cast(PanelCard, None)
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

        # Status section
        status_section = QWidget()
        status_section.setProperty("infoStatusSource", "idle")
        status_section.setProperty("infoStatusTone", "secondary")
        overview_layout = QVBoxLayout(status_section)
        overview_layout.setContentsMargins(0, 0, 0, 0)
        overview_layout.setSpacing(_SPACING_XS)

        # Meta label (e.g. "Task active...", "History (3)")
        self._status_meta_label = self._content_card.header

        # Summary label (main status text, potentially interactive)
        self._output_row = OutputFileRow(status_section)
        self._output_row.location_requested.connect(self._vm.location_requested.emit)
        self._output_row.details_requested.connect(lambda: self._vm.request_guide_action("view_outputs", ""))
        overview_layout.addWidget(self._output_row)
        self._status_summary_label = _SummaryButton(status_section)
        self._status_summary_label.setObjectName("infoStatusSummary")
        self._status_summary_label.clicked.connect(self._vm.request_navigation)
        overview_layout.addWidget(self._status_summary_label)
        self._output_destination_label = MiddleElidedLabel("", status_section)
        self._output_destination_label.setObjectName("infoOutputDestination")
        self._output_destination_label.hide()
        overview_layout.addWidget(self._output_destination_label)
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
        self._activity_button = _ActivityButton(self)
        self._activity_button.setObjectName("infoActivityButton")
        self._activity_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self._activity_button.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        self._activity_button.setAutoDefault(False)
        self._activity_button.setSizePolicy(QSizePolicy.Policy.Maximum, QSizePolicy.Policy.Fixed)
        self._activity_button.setIconSize(QSize(18, 18))
        self._activity_button.clicked.connect(self._vm.activity_requested.emit)
        card_layout.addWidget(self._activity_button, alignment=Qt.AlignmentFlag.AlignLeft)

        root_layout.addWidget(self._content_card)

    # ── ViewModel Wiring ─────────────────────────────────────────────────

    def _wire_vm(self) -> None:
        """Connect ViewModel signals to widget updates."""
        self._vm.state_changed.connect(self._sync_from_vm)
        self._sync_from_vm()

    def _sync_from_vm(self) -> None:
        """Rebuild the widget from current ViewModel state."""
        self._update_activity_button()
        self._update_status_section()
        self._update_guide_row()

    def _update_activity_button(self) -> None:
        count, failed = self._vm.activity_counts
        text = _t("info_area.activity_count", count=count)
        self._activity_button.setText(text)
        detail = _t("activity.failed_count", count=failed) if failed else text
        self._activity_button.setToolTip(detail)
        self._activity_button.setAccessibleName(f"{text} — {detail}" if failed else text)
        icon = load_svg_icon("error.svg" if failed else "logging.svg")
        if icon is not None:
            self._activity_button.setIcon(icon)
        self._activity_button.setProperty("hasFailures", failed > 0)
        self._refresh_widget_style(self._activity_button)
        self._activity_button.setVisible(count > 0)

    # ── Status section ───────────────────────────────────────────────────

    def _update_status_section(self) -> None:
        """Update the status meta, summary, and styling from ViewModel."""
        vm = self._vm

        idle = vm.status_source == "idle" and not vm.activity_enabled
        if self._status_meta_label is not None:
            if vm.activity_enabled:
                self._status_meta_label.setText(vm.activity_meta_text)
            else:
                icon = {"success": "✓", "danger": "✕", "warning": "⚠"}.get(vm.status_tone, "")
                title = vm.status_summary_text if idle else vm.status_meta_text
                self._status_meta_label.setText(f"{icon} {title}".strip())
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

        self._status_meta_label.setVisible(bool(self._status_meta_label.text()))
        self._status_summary_label.setVisible(not idle and bool(vm.status_summary_text))
        self._output_row.set_output(vm.task_summary.output_path, len(vm.task_summary.output_paths))
        self._output_row.setVisible(vm.show_output_file)
        self._output_destination_label.set_full_text(vm.output_destination_hint)
        self._output_destination_label.setVisible(bool(vm.output_destination_hint))
        self._content_card.setTone(
            vm.status_tone
            if vm.status_tone in {"primary", "success", "warning", "danger", "info", "secondary"}
            else "secondary"
        )
        self._notice.setText(vm.notification_text)
        self._notice.setVisible(bool(vm.notification_text))
        self._content_card.setContentVisible(
            not idle or bool(vm.output_destination_hint) or bool(vm.notification_text) or vm.activity_counts[0] > 0
        )
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
            button.setMinimumHeight(Sizing.CONTROL_HEIGHT)
            self._refresh_widget_style(button)

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
        for column in range(layout.columnCount()):
            layout.setColumnStretch(column, 0)
            layout.setColumnMinimumWidth(column, 0)
        for index, button in enumerate(buttons):
            layout.addWidget(button, index // columns, index % columns)
        for column in range(columns):
            layout.setColumnStretch(column, 1)
        self._status_guide_actions_widget.updateGeometry()

    # ── Helpers ──────────────────────────────────────────────────────────

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

    def find_guide_buttons(self) -> list[QPushButton]:
        """Find all guide buttons (for testing)."""
        if self._status_guide_row is None:
            return []
        return [
            btn
            for btn in self._status_guide_row.findChildren(QPushButton)
            if btn.objectName() == "infoStatusGuideButton"
        ]


__all__ = ["InfoArea"]

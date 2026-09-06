"""Shared visual primitives for the three-column workspace.

The primitives in this module deliberately own structure rather than colours.
Their semantic properties are consumed by the shared panel-card stylesheet so
all workflow panels use the same hierarchy without nesting framed group boxes.
"""

from __future__ import annotations

from collections.abc import Sequence
from typing import Protocol

from PySide6.QtCore import QEvent, QRect, Qt, QTimer, Signal
from PySide6.QtGui import QStandardItemModel
from PySide6.QtWidgets import (
    QBoxLayout,
    QComboBox,
    QFrame,
    QHBoxLayout,
    QLabel,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)


class FormatChoiceLike(Protocol):
    """Structural type accepted by :class:`FormatSelector`."""

    @property
    def display_name(self) -> str: ...

    @property
    def enabled(self) -> bool: ...

    @property
    def disabled_reason(self) -> str: ...

    @property
    def help_text(self) -> str: ...


class SectionHeader(QLabel):
    """A title inside a card; never painted across a frame border."""

    def __init__(self, title: str = "", parent: QWidget | None = None, *, level: str = "section") -> None:
        super().__init__(title, parent)
        if level not in {"card", "section"}:
            raise ValueError("section header level must be 'card' or 'section'")
        self.setProperty("headerLevel", level)
        self.setObjectName("panelCardTitle" if level == "card" else "panelSectionTitle")
        self.setWordWrap(True)
        alignment = (
            Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter
            if level == "card"
            else Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter
        )
        self.setAlignment(alignment)


class _ResponsiveFrame(QFrame):
    """Coalesce geometry, content and typography changes into one reflow."""

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._reflow_timer = QTimer(self)
        self._reflow_timer.setSingleShot(True)
        self._reflow_timer.timeout.connect(self._sync_layout)

    def event(self, event: QEvent) -> bool:
        handled = super().event(event)
        if event.type() in {
            QEvent.Type.LayoutRequest,
            QEvent.Type.FontChange,
            QEvent.Type.StyleChange,
            QEvent.Type.Show,
            QEvent.Type.Resize,
        }:
            timer = getattr(self, "_reflow_timer", None)
            if timer is not None and not timer.isActive():
                timer.start(0)
        return handled

    def _sync_layout(self) -> None:
        raise NotImplementedError


class WrappingLabel(QLabel):
    """Keep wrapped text's minimum height current after text, font or width changes."""

    def __init__(self, text: str = "", parent: QWidget | None = None) -> None:
        super().__init__(text, parent)
        self.setWordWrap(True)

    def sync_wrapped_height(self) -> None:
        margins = self.contentsMargins()
        width = max(1, self.width() - margins.left() - margins.right() - 2 * self.margin())
        height = (
            self.fontMetrics().boundingRect(QRect(0, 0, width, 100000), Qt.TextFlag.TextWordWrap, self.text()).height()
        )
        height += margins.top() + margins.bottom() + 2 * self.margin()
        if self.minimumHeight() != height:
            self.setMinimumHeight(height)
            self.updateGeometry()

    def setText(self, text: str) -> None:
        super().setText(text)
        self.sync_wrapped_height()

    def resizeEvent(self, event) -> None:
        super().resizeEvent(event)
        self.sync_wrapped_height()

    def changeEvent(self, event) -> None:
        super().changeEvent(event)
        if event.type() in {QEvent.Type.FontChange, QEvent.Type.StyleChange}:
            self.sync_wrapped_height()


class _FormLabel(QLabel):
    geometryChanged = Signal()

    def setText(self, text: str) -> None:
        super().setText(text)
        self.geometryChanged.emit()

    def changeEvent(self, event) -> None:
        super().changeEvent(event)
        if event.type() in {QEvent.Type.FontChange, QEvent.Type.StyleChange}:
            self.geometryChanged.emit()


class FormRow(_ResponsiveFrame):
    """A label/control row that stacks before translated content clips."""

    def __init__(
        self,
        label: str,
        control: QWidget,
        parent: QWidget | None = None,
        *,
        label_suffix: QWidget | None = None,
        alignment_group: Sequence[FormRow] | None = None,
        minimum_label_height: int = 0,
    ) -> None:
        super().__init__(parent)
        self.setObjectName("panelFormRow")
        self.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Preferred)
        self.content_layout = QBoxLayout(QBoxLayout.Direction.LeftToRight, self)
        self.content_layout.setContentsMargins(0, 0, 0, 0)
        self.content_layout.setSpacing(8)
        self.label = _FormLabel(label, self)
        self.label.geometryChanged.connect(self._schedule_group_reflow)
        self.label.setObjectName("panelFormLabel")
        self.label.setBuddy(control)
        self.label.setMinimumWidth(0)
        self.label.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Preferred)
        self.label_suffix = label_suffix
        self.alignment_group = alignment_group
        self.minimum_label_height = minimum_label_height
        self.label_container: QWidget = self.label
        if label_suffix is not None:
            self.label_container = QWidget(self)
            label_layout = QHBoxLayout(self.label_container)
            label_layout.setContentsMargins(0, 0, 0, 0)
            label_layout.setSpacing(6)
            label_layout.addWidget(self.label, 1)
            label_layout.addWidget(label_suffix, 0, Qt.AlignmentFlag.AlignVCenter)
        self.control = control
        self.content_layout.addWidget(self.label_container)
        self.content_layout.addWidget(control, stretch=1)

    def _schedule_group_reflow(self) -> None:
        for row in self.alignment_group or (self,):
            row._reflow_timer.start(0)

    def _sync_layout(self) -> None:
        peers = self.alignment_group or (self,)
        column_width = max(
            row.label.fontMetrics().horizontalAdvance(row.label.text())
            + (row.label_suffix.sizeHint().width() + 6 if row.label_suffix is not None else 0)
            for row in peers
        )
        suffix_width = self.label_suffix.sizeHint().width() + 6 if self.label_suffix is not None else 0
        label_width = column_width - suffix_width
        control_min_width = max(row._control_readable_width() for row in peers)
        required_width = column_width + control_min_width + 8
        horizontal = required_width <= self.contentsRect().width()
        self.label.setWordWrap(not horizontal)
        if horizontal:
            self.label.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Preferred)
            self.label.setFixedWidth(label_width)
            self.label_container.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Preferred)
            self.label_container.setFixedWidth(label_width + suffix_width)
        else:
            self.label.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Preferred)
            self.label.setMinimumWidth(0)
            self.label.setMaximumWidth(16777215)
            self.label_container.setMinimumWidth(0)
            self.label_container.setMaximumWidth(16777215)
            self.label_container.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Preferred)
        direction = QBoxLayout.Direction.LeftToRight if horizontal else QBoxLayout.Direction.TopToBottom
        if self.content_layout.direction() != direction:
            self.content_layout.setDirection(direction)
            self.content_layout.setSpacing(8 if horizontal else 4)
            self.updateGeometry()
        text_width = label_width if horizontal else max(1, self.contentsRect().width() - suffix_width)
        margins = self.label.contentsMargins()
        horizontal_padding = margins.left() + margins.right() + 2 * self.label.margin()
        vertical_padding = margins.top() + margins.bottom() + 2 * self.label.margin()
        label_height = self.label.fontMetrics().height()
        if self.label.wordWrap():
            label_height = (
                self.label.fontMetrics()
                .boundingRect(
                    QRect(0, 0, max(1, text_width - horizontal_padding), 100000),
                    Qt.TextFlag.TextWordWrap,
                    self.label.text(),
                )
                .height()
            )
        label_height += vertical_padding
        label_height = max(label_height, self.minimum_label_height)
        if self.label_suffix is not None:
            label_height = max(label_height, self.label_suffix.sizeHint().height())
        self.label_container.setFixedHeight(label_height)
        control_width = (
            max(1, self.contentsRect().width() - label_width - suffix_width - 8)
            if horizontal
            else self.contentsRect().width()
        )
        control_height = (
            self.control.heightForWidth(control_width)
            if self.control.hasHeightForWidth()
            else self.control.sizeHint().height()
        )
        control_height = max(control_height, self.control.minimumHeight(), self.control.minimumSizeHint().height())
        self.setFixedHeight(max(label_height, control_height) if horizontal else label_height + control_height + 4)

    def _control_readable_width(self) -> int:
        width = max(self.control.minimumSizeHint().width(), self.control.minimumWidth())
        if isinstance(self.control, QComboBox):
            metrics = self.control.fontMetrics()
            text_width = max(
                (metrics.horizontalAdvance(self.control.itemText(index)) for index in range(self.control.count())),
                default=0,
            )
            width = max(width, text_width + 48)
        return width


class ChoiceGroup(_ResponsiveFrame):
    """Unframed container for related choices.

    A responsive group uses a comfortably separated horizontal scan pattern
    while it fits, then stacks without clipping at narrow widths or under
    larger translated fonts.
    """

    def __init__(self, parent: QWidget | None = None, *, responsive: bool = False, spacing: int = 24) -> None:
        super().__init__(parent)
        self.setObjectName("panelChoiceGroup")
        self._responsive = responsive
        self._horizontal_spacing = spacing
        if responsive:
            self.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Preferred)
        direction = QBoxLayout.Direction.LeftToRight if responsive else QBoxLayout.Direction.TopToBottom
        self.content_layout = QBoxLayout(direction, self)
        self.content_layout.setContentsMargins(0, 0, 0, 0)
        self.content_layout.setSpacing(spacing if responsive else 8)
        if responsive:
            self.content_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

    def _sync_layout(self) -> None:
        if not self._responsive:
            return
        widgets: list[QWidget] = []
        for index in range(self.content_layout.count()):
            item = self.content_layout.itemAt(index)
            widget = item.widget() if item is not None else None
            if widget is not None:
                widgets.append(widget)
        required = sum(widget.sizeHint().width() for widget in widgets)
        required += max(0, len(widgets) - 1) * self._horizontal_spacing
        horizontal = required <= self.contentsRect().width()
        self.content_layout.setDirection(
            QBoxLayout.Direction.LeftToRight if horizontal else QBoxLayout.Direction.TopToBottom
        )
        self.content_layout.setSpacing(self._horizontal_spacing if horizontal else 8)


class ActionFooter(QFrame):
    """Unframed row reserved for the action that follows its options."""

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("panelActionFooter")
        self.content_layout = QHBoxLayout(self)
        self.content_layout.setContentsMargins(0, 0, 0, 0)
        self.content_layout.setSpacing(8)


class InlineNotice(QFrame):
    """Compact semantic notice used for warnings and availability states."""

    def __init__(self, text: str = "", parent: QWidget | None = None, *, tone: str = "info") -> None:
        super().__init__(parent)
        self.setObjectName("panelInlineNotice")
        self.setProperty("noticeTone", tone)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(8, 6, 8, 6)
        self.label = QLabel(text, self)
        self.label.setWordWrap(True)
        layout.addWidget(self.label, stretch=1)

    def setText(self, text: str) -> None:
        self.label.setText(text)


class TaskActivityList(QFrame):
    """Unframed host for the task-history rows in the feedback panel."""

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("taskActivityList")
        self.content_layout = QVBoxLayout(self)
        self.content_layout.setContentsMargins(0, 0, 0, 0)
        self.content_layout.setSpacing(8)
        self.content_layout.setAlignment(Qt.AlignmentFlag.AlignTop)


class FormatSelector(QComboBox):
    """A combo that retains unavailable targets with an explicit reason."""

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("formatSelector")
        self.setMinimumWidth(100)
        self.setSizeAdjustPolicy(QComboBox.SizeAdjustPolicy.AdjustToContents)
        self.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Fixed)
        self.currentIndexChanged.connect(self._sync_tooltip)

    def set_choices(self, choices: Sequence[FormatChoiceLike], *, selected: str | None = None) -> None:
        """Replace items while selecting the first enabled target when needed."""

        previous = self.currentText()
        self.clear()
        model = self.model()
        if not isinstance(model, QStandardItemModel):
            raise TypeError("FormatSelector requires a QStandardItemModel")

        first_enabled = -1
        requested_index = -1
        for index, choice in enumerate(choices):
            self.addItem(choice.display_name)
            item = model.item(index)
            item.setEnabled(choice.enabled)
            tooltip = choice.help_text if choice.enabled else choice.disabled_reason
            item.setToolTip(tooltip)
            self.setItemData(index, choice.enabled, Qt.ItemDataRole.UserRole)
            self.setItemData(index, tooltip, Qt.ItemDataRole.ToolTipRole)
            if choice.enabled and first_enabled < 0:
                first_enabled = index
            if choice.enabled and choice.display_name == (selected or previous):
                requested_index = index

        self.setCurrentIndex(requested_index if requested_index >= 0 else first_enabled)
        longest = max(
            (self.fontMetrics().horizontalAdvance(self.itemText(index)) for index in range(self.count())), default=0
        )
        self.setMinimumWidth(max(100, longest + 48))
        self._sync_tooltip()

    def current_choice_enabled(self) -> bool:
        """Whether the current item can be submitted."""

        return self.currentIndex() >= 0 and bool(self.currentData(Qt.ItemDataRole.UserRole))

    def _sync_tooltip(self) -> None:
        self.setToolTip(str(self.currentData(Qt.ItemDataRole.ToolTipRole) or ""))


class PanelCard(QFrame):
    """Neutral card with an internal title and dedicated content layout.

    Card titles are compact and centred. Section titles are left aligned and
    intentionally avoid another framed surface, so nested options do not turn
    into a stack of fieldsets.
    """

    def __init__(
        self,
        title: str = "",
        parent: QWidget | None = None,
        *,
        level: str = "card",
    ) -> None:
        super().__init__(parent)
        if level not in {"card", "section"}:
            raise ValueError("panel card level must be 'card' or 'section'")

        self.setProperty("panelLevel", level)
        outer = QVBoxLayout(self)
        margin_x = 12 if level == "card" else 0
        margin_top = 10 if level == "card" else 4
        margin_bottom = 12 if level == "card" else 4
        outer.setContentsMargins(margin_x, margin_top, margin_x, margin_bottom)
        outer.setSpacing(8)

        self._title_label = SectionHeader(parent=self, level=level)
        outer.addWidget(self._title_label)

        self._content = QWidget(self)
        self._content.setObjectName("panelCardContent")
        self._content_layout = QVBoxLayout(self._content)
        self._content_layout.setContentsMargins(0, 0, 0, 0)
        self._content_layout.setSpacing(8)
        outer.addWidget(self._content)

        self.setTitle(title)

    @property
    def content_layout(self) -> QVBoxLayout:
        """Return the layout owned by the card's content region."""

        return self._content_layout

    def setTitle(self, title: str) -> None:
        """Update the internal title without placing text on the border."""

        self._title_label.setText(title)
        self._title_label.setVisible(bool(title.strip()))

    def title(self) -> str:
        """Return the current internal title."""

        return self._title_label.text()


__all__ = [
    "ActionFooter",
    "ChoiceGroup",
    "FormRow",
    "FormatSelector",
    "InlineNotice",
    "PanelCard",
    "SectionHeader",
    "TaskActivityList",
]

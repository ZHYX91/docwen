"""Consistent, content-sized software priority controls for settings pages."""

from PySide6.QtCore import QEvent, QRect, QSize, Qt, QTimer
from PySide6.QtWidgets import QHBoxLayout, QLabel, QListWidget, QPushButton, QVBoxLayout, QWidget

from ...i18n import t


class SoftwarePriorityEditor(QWidget):
    """Place the category above a full-width list with aligned move buttons."""

    def __init__(self, title: str, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("settingsSoftwarePriorityEditor")
        self._resize_timer = QTimer(self)
        self._resize_timer.setSingleShot(True)
        self._resize_timer.timeout.connect(self._sync_geometry)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)
        self.title_label = QLabel(title, self)
        self.title_label.setTextFormat(Qt.TextFormat.PlainText)
        self.title_label.setWordWrap(True)
        self.title_label.setProperty("settingsRole", "fieldLabel")
        layout.addWidget(self.title_label)

        row = QHBoxLayout()
        row.setSpacing(8)
        self.list_widget = QListWidget(self)
        self.list_widget.setObjectName("settingsPriorityList")
        self.list_widget.setAccessibleName(title)
        self.list_widget.setWordWrap(True)
        self.list_widget.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.title_label.setBuddy(self.list_widget)
        row.addWidget(self.list_widget, 1, Qt.AlignmentFlag.AlignTop)
        buttons = QVBoxLayout()
        buttons.setSpacing(6)
        self.move_up_button = QPushButton(t("editors.common.move_up", "Move Up"), self)
        self.move_down_button = QPushButton(t("editors.common.move_down", "Move Down"), self)
        buttons.addWidget(self.move_up_button)
        buttons.addWidget(self.move_down_button)
        buttons.addStretch(1)
        row.addLayout(buttons)
        layout.addLayout(row)

        for widget in (self, self.list_widget, self.list_widget.viewport()):
            widget.installEventFilter(self)
        model = self.list_widget.model()
        for signal in (model.rowsInserted, model.rowsRemoved, model.modelReset, model.dataChanged):
            signal.connect(self._schedule_geometry)
        self._schedule_geometry()

    def eventFilter(self, watched, event) -> bool:
        if event.type() in {QEvent.Type.Resize, QEvent.Type.FontChange, QEvent.Type.StyleChange, QEvent.Type.Show}:
            self._schedule_geometry()
        return super().eventFilter(watched, event)

    def _schedule_geometry(self, *_args) -> None:
        if not self._resize_timer.isActive():
            self._resize_timer.start(0)

    def _sync_geometry(self) -> None:
        buttons = (self.move_up_button, self.move_down_button)
        button_width = max(72, *(button.sizeHint().width() for button in buttons))
        button_height = max(32, *(button.sizeHint().height() for button in buttons))
        for button in buttons:
            button.setFixedSize(button_width, button_height)
        width = max(1, self.list_widget.viewport().width() - 16)
        metrics = self.list_widget.fontMetrics()
        row_height = max(32, metrics.height() + 12)
        for index in range(self.list_widget.count()):
            text = self.list_widget.item(index).text()
            bounds = metrics.boundingRect(QRect(0, 0, width, 10000), Qt.TextFlag.TextWordWrap, text)
            row_height = max(row_height, bounds.height() + 12)
        size = QSize(0, row_height)
        for index in range(self.list_widget.count()):
            item = self.list_widget.item(index)
            if item.sizeHint() != size:
                item.setSizeHint(size)
        self.list_widget.setFixedHeight(
            max(1, self.list_widget.count()) * row_height + 2 * self.list_widget.frameWidth()
        )

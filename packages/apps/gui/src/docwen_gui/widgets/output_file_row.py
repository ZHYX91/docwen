"""A compact output identity, multiple-results badge, and file-location action."""

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import QHBoxLayout, QToolButton, QWidget

from docwen_gui.i18n import t
from docwen_gui.styles.design_tokens import Sizing, Spacing

from .elided_label import MiddleElidedLabel
from .location_button import LocationButton


class OutputFileRow(QWidget):
    location_requested = Signal(str)
    details_requested = Signal()

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._path = ""
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(Spacing.CONTROL_GAP)
        self.name_label = MiddleElidedLabel("", self)
        self.name_label.setObjectName("outputFileName")
        layout.addWidget(self.name_label, 1)
        self.count_button = QToolButton(self)
        self.count_button.setObjectName("outputFileCount")
        self.count_button.setMinimumHeight(Sizing.CONTROL_HEIGHT)
        self.count_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.count_button.clicked.connect(self.details_requested.emit)
        layout.addWidget(self.count_button)
        self.location_button = LocationButton(self, label=t("file_locations.output"))
        self.location_button.clicked.connect(self._locate)
        layout.addWidget(self.location_button)
        self.set_output("", 0)

    def set_output(self, path: str, count: int) -> None:
        self._path = path
        name = path.replace("\\", "/").rsplit("/", 1)[-1]
        self.name_label.set_full_text(t("info_area.task_output_file", name=name) if path else "")
        self.name_label.setToolTip(path)
        self.location_button.setAccessibleDescription(path)
        self.count_button.setText(t("file_locations.multiple", count=count))
        self.count_button.setToolTip(t("file_locations.show_outputs"))
        self.count_button.setAccessibleName(t("file_locations.show_outputs") + f" ({count})")
        self.count_button.setVisible(bool(path) and count > 1)
        self.setVisible(bool(path))

    def _locate(self) -> None:
        if self._path:
            self.location_requested.emit(self._path)

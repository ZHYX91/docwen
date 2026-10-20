"""Keyboard and accessible names for the Qt settings sidebar."""

from __future__ import annotations

from PySide6.QtCore import QEvent, QObject, QRectF, QSize, Qt, QTimer, Signal
from PySide6.QtGui import QIcon, QKeyEvent, QPainter, QPaintEvent, QPalette, QPen
from PySide6.QtWidgets import QFrame, QLabel, QPushButton, QScrollArea, QSizePolicy, QStyle, QVBoxLayout, QWidget

from docwen_gui.styles.design_tokens import Sizing, Spacing
from docwen_gui.styles.ui_scale import dp, set_metric


class SettingsSidebar(QScrollArea):
    """Grouped page buttons stay reachable without taking height from the content."""

    page_requested = Signal(str)

    def __init__(self, parent: QWidget) -> None:
        super().__init__(parent)
        self.setObjectName("settingsNavigation")
        self.setFrameShape(QFrame.Shape.NoFrame)
        self.setWidgetResizable(True)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Expanding)
        self._buttons: dict[str, QPushButton] = {}
        self._group_labels: dict[str, QLabel] = {}
        self._current_key: str | None = None
        self._content = QWidget(self)
        self._content.setObjectName("settingsNavigationContent")
        self._layout = QVBoxLayout(self._content)
        self._layout.setContentsMargins(0, 0, 0, 0)
        set_metric(self._layout, "setSpacing", Spacing.XS)
        self._layout.addStretch(1)
        self.setWidget(self._content)

    def add_group(self, key: str, title: str) -> None:
        """Add a non-interactive heading immediately before a nonempty group."""
        if self._group_labels:
            separator = QFrame(self._content)
            separator.setObjectName(f"settingsNavSeparator_{key}")
            separator.setProperty("settingsNavigationSeparator", True)
            separator.setFrameShape(QFrame.Shape.HLine)
            separator.setFocusPolicy(Qt.FocusPolicy.NoFocus)
            self._layout.insertSpacing(self._layout.count() - 1, dp(Spacing.SM))
            self._layout.insertWidget(self._layout.count() - 1, separator)
        label = QLabel(title, self._content)
        label.setObjectName(f"settingsNavGroup_{key}")
        label.setProperty("settingsRole", "navigationGroup")
        label.setTextFormat(Qt.TextFormat.PlainText)
        label.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self._group_labels[key] = label
        self._layout.insertWidget(self._layout.count() - 1, label)

    def add_page(self, key: str, title: str, icon: QIcon | None) -> QPushButton:
        button = QPushButton(title, self._content)
        button.setProperty("settingsNavigationItem", True)
        button.setCheckable(True)
        button.setAutoExclusive(True)
        button.setAutoDefault(False)
        button.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum)
        button.setToolTip(title)
        if icon is not None:
            button.setIcon(icon)
        set_metric(button, "setIconSize", QSize(18, 18))
        set_metric(button, "setMinimumHeight", Sizing.CONTROL_HEIGHT)
        button.clicked.connect(lambda _checked=False: self.page_requested.emit(key))
        self._buttons[key] = button
        self._layout.insertWidget(self._layout.count() - 1, button)
        return button

    def page_button(self, key: str) -> QPushButton | None:
        return self._buttons.get(key)

    def select_page(self, key: str) -> None:
        button = self.page_button(key)
        if button is not None:
            self._current_key = key
            button.setChecked(True)
            self.ensure_current_visible()

    def ensure_current_visible(self) -> None:
        if not self.isVisible() or self._current_key is None:
            return
        button = self.page_button(self._current_key)
        if button is not None:
            self.ensureWidgetVisible(button, 0, dp(Spacing.XS))

    def resizeEvent(self, event) -> None:
        super().resizeEvent(event)
        # The scroll range settles after its viewport and child layout resize.
        QTimer.singleShot(0, self, self.ensure_current_visible)

    def refresh_width(self) -> None:
        labels = [*self._buttons.values(), *self._group_labels.values()]
        for label in labels:
            label.ensurePolished()
        scrollbar_width = self.style().pixelMetric(QStyle.PixelMetric.PM_ScrollBarExtent)
        self.setFixedWidth(
            max(dp(168), max((label.sizeHint().width() for label in labels), default=0) + scrollbar_width)
        )


class _NavigationFocusFrame(QWidget):
    """A mouse-transparent overlay unaffected by native focus-frame masks."""

    def __init__(self, parent: QWidget) -> None:
        super().__init__(parent)
        self.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)
        self.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.hide()

    def follow(self, widget: QWidget) -> None:
        self.setParent(widget)
        self.setGeometry(widget.rect())
        self.show()
        self.raise_()

    def paintEvent(self, event: QPaintEvent) -> None:
        del event
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.setPen(QPen(self.palette().color(QPalette.ColorRole.Highlight), 2.0))
        painter.setBrush(Qt.BrushStyle.NoBrush)
        painter.drawRoundedRect(QRectF(self.rect()).adjusted(1, 1, -1, -1), 5, 5)
        painter.end()


class SettingsNavigationKeyboard(QObject):
    """Keep one sidebar tab stop, with arrows selecting a page without losing focus."""

    page_requested = Signal(str)

    def __init__(self, parent: QWidget) -> None:
        super().__init__(parent)
        self._items: dict[str, QWidget] = {}
        self._focus_frame = _NavigationFocusFrame(parent)
        self._focus_frame.setObjectName("settingsNavigationFocusFrame")

    def add(self, key: str, entry: QWidget, title: str) -> None:
        item = getattr(entry, "itemWidget", entry)
        if not isinstance(item, QWidget):
            raise TypeError("settings_navigation_item_must_be_widget")
        entry.setAccessibleName(title)
        entry.setObjectName(f"settingsNavEntry_{key}")
        item.setAccessibleName(title)
        item.setObjectName(f"settingsNavItem_{key}")
        if entry is not item:
            entry.setFocusProxy(item)
        item.installEventFilter(self)
        self._items[key] = item
        item.setFocusPolicy(Qt.FocusPolicy.ClickFocus)

    def sync(self, key: str, content: QWidget | None = None) -> None:
        for item_key, item in self._items.items():
            item.setFocusPolicy(Qt.FocusPolicy.StrongFocus if item_key == key else Qt.FocusPolicy.ClickFocus)
        current = self._items.get(key)
        if current is not None and content is not None:
            QWidget.setTabOrder(current, content)

    def eventFilter(self, watched: QObject, event: QEvent) -> bool:
        key = next((key for key, item in self._items.items() if item is watched), None)
        if key is None or not isinstance(watched, QWidget):
            return super().eventFilter(watched, event)
        if event.type() == QEvent.Type.FocusIn:
            self._focus_frame.follow(watched)
        elif event.type() == QEvent.Type.FocusOut:
            self._focus_frame.hide()
        elif event.type() == QEvent.Type.Resize and self._focus_frame.parentWidget() is watched:
            self._focus_frame.setGeometry(watched.rect())
        elif event.type() == QEvent.Type.KeyPress and isinstance(event, QKeyEvent):
            keys = list(self._items)
            position = keys.index(key)
            destinations: dict[int, str] = {
                Qt.Key.Key_Up: keys[max(0, position - 1)],
                Qt.Key.Key_Down: keys[min(len(keys) - 1, position + 1)],
                Qt.Key.Key_Home: keys[0],
                Qt.Key.Key_End: keys[-1],
                Qt.Key.Key_Return: key,
                Qt.Key.Key_Enter: key,
                Qt.Key.Key_Space: key,
            }
            destination = destinations.get(event.key())
            if destination is not None:
                self.page_requested.emit(destination)
                self._items[destination].setFocus(Qt.FocusReason.TabFocusReason)
                event.accept()
                return True
        return super().eventFilter(watched, event)

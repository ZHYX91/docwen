"""Keyboard and accessible names for the custom-painted settings sidebar."""

from __future__ import annotations

from PySide6.QtCore import QEvent, QObject, QRectF, Qt
from PySide6.QtGui import QKeyEvent, QPainter, QPaintEvent, QPalette, QPen
from PySide6.QtWidgets import QTabWidget, QWidget


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

    def __init__(self, tabs: QTabWidget, parent: QWidget) -> None:
        super().__init__(parent)
        self._tabs = tabs
        self._items: dict[int, QWidget] = {}
        self._focus_frame = _NavigationFocusFrame(parent)
        self._focus_frame.setObjectName("settingsNavigationFocusFrame")

    def add(self, index: int, key: str, entry: QWidget, title: str) -> None:
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
        self._items[index] = item
        self.sync(self._tabs.currentIndex())

    def sync(self, index: int, content: QWidget | None = None) -> None:
        for item_index, item in self._items.items():
            item.setFocusPolicy(Qt.FocusPolicy.StrongFocus if item_index == index else Qt.FocusPolicy.ClickFocus)
        current = self._items.get(index)
        if current is not None and content is not None:
            QWidget.setTabOrder(current, content)

    def eventFilter(self, watched: QObject, event: QEvent) -> bool:
        index = next((key for key, item in self._items.items() if item is watched), None)
        if index is None or not isinstance(watched, QWidget):
            return super().eventFilter(watched, event)
        if event.type() == QEvent.Type.FocusIn:
            self._focus_frame.follow(watched)
        elif event.type() == QEvent.Type.FocusOut:
            self._focus_frame.hide()
        elif event.type() == QEvent.Type.Resize and self._focus_frame.parentWidget() is watched:
            self._focus_frame.setGeometry(watched.rect())
        elif event.type() == QEvent.Type.KeyPress and isinstance(event, QKeyEvent):
            keys = sorted(self._items)
            position = keys.index(index)
            destinations: dict[int, int] = {
                Qt.Key.Key_Up: keys[max(0, position - 1)],
                Qt.Key.Key_Down: keys[min(len(keys) - 1, position + 1)],
                Qt.Key.Key_Home: keys[0],
                Qt.Key.Key_End: keys[-1],
                Qt.Key.Key_Return: index,
                Qt.Key.Key_Enter: index,
                Qt.Key.Key_Space: index,
            }
            destination = destinations.get(event.key())
            if destination is not None:
                self._tabs.setCurrentIndex(destination)
                self._items[destination].setFocus(Qt.FocusReason.TabFocusReason)
                event.accept()
                return True
        return super().eventFilter(watched, event)

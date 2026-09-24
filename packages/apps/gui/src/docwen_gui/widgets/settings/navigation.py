"""Keyboard and accessible names for the Qt settings sidebar."""

from __future__ import annotations

from PySide6.QtCore import QEvent, QObject, QRectF, QSize, Qt, Signal
from PySide6.QtGui import QIcon, QKeyEvent, QPainter, QPaintEvent, QPalette, QPen
from PySide6.QtWidgets import QFrame, QPushButton, QSizePolicy, QTabWidget, QVBoxLayout, QWidget

from docwen_gui.styles.design_tokens import Sizing, Spacing
from docwen_gui.styles.ui_scale import dp, set_metric


class SettingsSidebar(QFrame):
    """Native Qt buttons share application typography, metrics and vector icons."""

    page_requested = Signal(int)

    def __init__(self, parent: QWidget) -> None:
        super().__init__(parent)
        self.setObjectName("settingsNavigation")
        self._buttons: dict[str, QPushButton] = {}
        self._layout = QVBoxLayout(self)
        self._layout.setContentsMargins(0, 0, 0, 0)
        set_metric(self._layout, "setSpacing", Spacing.XS)
        self._layout.addStretch(1)

    def add_page(self, key: str, index: int, title: str, icon: QIcon | None) -> QPushButton:
        button = QPushButton(title, self)
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
        button.clicked.connect(lambda _checked=False: self.page_requested.emit(index))
        self._buttons[key] = button
        self._layout.insertWidget(self._layout.count() - 1, button)
        return button

    def page_button(self, key: str) -> QPushButton | None:
        return self._buttons.get(key)

    def select_page(self, key: str) -> None:
        button = self.page_button(key)
        if button is not None:
            button.setChecked(True)

    def refresh_width(self) -> None:
        for button in self._buttons.values():
            button.ensurePolished()
        self.setFixedWidth(
            max(dp(168), max((button.sizeHint().width() for button in self._buttons.values()), default=0))
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

"""Help stays readable at screen edges without breaking keyboard navigation."""

from __future__ import annotations

import pytest
from PySide6.QtCore import QEvent, QPoint, Qt
from PySide6.QtWidgets import QApplication, QFrame, QLabel, QMenu, QScrollArea

from docwen_gui.widgets.settings.base_tab import _create_info_button

pytestmark = pytest.mark.gui


def test_keyboard_help_is_plain_scrollable_bounded_and_restores_focus(qapp, qtbot) -> None:
    text = "<literal>\n" + "Long explanation with readable line breaks.\n" * 70
    button = _create_info_button(text, accessible_name="Example field")
    qtbot.addWidget(button)
    screen = button.screen().availableGeometry()
    button.move(screen.bottomRight() - QPoint(40, 40))
    button.show()
    button.activateWindow()
    button.setFocus()
    qapp.processEvents()
    qtbot.keyClick(button, Qt.Key.Key_Return)
    menu = button.findChild(QMenu, "settingsHelpPopup")
    assert menu is not None and menu.isVisible()
    label = menu.findChild(QLabel, "settingsHelpPopupText")
    scroll = menu.findChild(QScrollArea)
    assert label is not None and scroll is not None
    qapp.processEvents()
    assert label.text() == text and label.textFormat() == Qt.TextFormat.PlainText
    assert screen.contains(menu.geometry())
    assert scroll.verticalScrollBar().maximum() > 0
    assert qapp.focusWidget() is scroll
    qtbot.keyClick(scroll, Qt.Key.Key_PageDown)
    assert scroll.verticalScrollBar().value() > 0
    qtbot.keyClick(scroll, Qt.Key.Key_Escape)
    assert not menu.isVisible()
    assert qapp.focusWidget() is button
    qtbot.keyClick(button, Qt.Key.Key_Return)
    assert len(button.findChildren(QMenu, "settingsHelpPopup")) == 1
    menu.close()


def test_hover_help_is_bounded_and_keeps_focus(qapp, qtbot) -> None:
    button = _create_info_button("First line\nSecond line")
    qtbot.addWidget(button)
    button.show()
    button.activateWindow()
    button.setFocus()
    qapp.processEvents()
    previous = qapp.focusWidget()
    QApplication.sendEvent(button, QEvent(QEvent.Type.ToolTip))
    popup = button.findChild(QFrame, "settingsHelpPopup")
    assert popup is not None and popup.isVisible()
    assert button.screen().availableGeometry().contains(popup.geometry())
    assert qapp.focusWidget() is previous
    QApplication.sendEvent(button, QEvent(QEvent.Type.Leave))
    assert not popup.isVisible()

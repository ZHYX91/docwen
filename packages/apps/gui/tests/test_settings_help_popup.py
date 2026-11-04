"""Help stays readable at screen edges without breaking keyboard navigation."""

from __future__ import annotations

import pytest
from PySide6.QtCore import QEvent, QPoint, Qt
from PySide6.QtGui import QCursor, QWheelEvent
from PySide6.QtWidgets import QApplication, QFrame, QLabel, QLineEdit, QMenu, QScrollArea, QVBoxLayout, QWidget

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
    QCursor.setPos(button.screen().availableGeometry().bottomRight())
    QApplication.sendEvent(button, QEvent(QEvent.Type.Leave))
    qtbot.waitUntil(lambda: button._hover_help is None)


@pytest.mark.parametrize("corner", ["top_left", "bottom_right"])
def test_hover_can_be_entered_scrolled_and_left(qapp, qtbot, corner) -> None:
    button = _create_info_button("Readable line\n" * 100)
    qtbot.addWidget(button)
    screen = button.screen().availableGeometry()
    button.move(screen.topLeft() + QPoint(20, 20) if corner == "top_left" else screen.bottomRight() - QPoint(40, 40))
    button.show()
    button.activateWindow()
    button.setFocus()
    qapp.processEvents()
    QApplication.sendEvent(button, QEvent(QEvent.Type.ToolTip))
    popup = button._hover_help
    assert popup is not None and screen.contains(popup.geometry())
    scroll = popup.findChild(QScrollArea)
    assert scroll is not None
    # Traverse the icon-to-popup gap without dismissing the text.
    QCursor.setPos(popup.mapToGlobal(QPoint(15, 15)))
    QApplication.sendEvent(button, QEvent(QEvent.Type.Leave))
    qtbot.mouseMove(scroll.viewport(), QPoint(15, 15))
    qtbot.wait(230)
    assert popup.isVisible()
    assert qapp.focusWidget() is button
    viewport = scroll.viewport()
    local = viewport.rect().center()
    wheel = QWheelEvent(
        local,
        viewport.mapToGlobal(local),
        QPoint(),
        QPoint(0, -120),
        Qt.MouseButton.NoButton,
        Qt.KeyboardModifier.NoModifier,
        Qt.ScrollPhase.NoScrollPhase,
        False,
    )
    QApplication.sendEvent(viewport, wheel)
    assert scroll.verticalScrollBar().value() > 0
    # Repeated tooltip requests retain the existing popup and scroll position.
    QApplication.sendEvent(button, QEvent(QEvent.Type.ToolTip))
    assert button._hover_help is popup
    assert scroll.verticalScrollBar().value() > 0
    QCursor.setPos(screen.topRight())
    QApplication.sendEvent(popup, QEvent(QEvent.Type.Leave))
    qtbot.waitUntil(lambda: button._hover_help is None)


@pytest.mark.parametrize("dismissal", ["escape", "outside_click", "focus", "hide", "deactivate"])
def test_hover_dismissal_and_reopen(qapp, qtbot, dismissal) -> None:
    host = QWidget()
    layout = QVBoxLayout(host)
    button = _create_info_button("First line\nSecond line", host)
    other = QLineEdit(host)
    layout.addWidget(button)
    layout.addWidget(other)
    qtbot.addWidget(host)
    host.show()
    host.activateWindow()
    button.setFocus()
    qapp.processEvents()
    QApplication.sendEvent(button, QEvent(QEvent.Type.ToolTip))
    assert button._hover_help is not None
    if dismissal == "escape":
        qtbot.keyClick(button, Qt.Key.Key_Escape)
    elif dismissal == "outside_click":
        qtbot.mouseClick(other, Qt.MouseButton.LeftButton)
    elif dismissal == "focus":
        other.setFocus()
    elif dismissal == "hide":
        host.hide()
    else:
        QApplication.sendEvent(host, QEvent(QEvent.Type.WindowDeactivate))
    qtbot.waitUntil(lambda: button._hover_help is None)
    host.show()
    host.activateWindow()
    button.setFocus()
    qapp.processEvents()
    QApplication.sendEvent(button, QEvent(QEvent.Type.ToolTip))
    assert button._hover_help is not None
    qtbot.keyClick(button, Qt.Key.Key_Return)
    assert button._hover_help is None
    menu = button.findChild(QMenu, "settingsHelpPopup")
    assert menu is not None and menu.isVisible()
    menu.close()


def test_registered_label_tooltip_does_not_capture_unrelated_tooltips(qapp, qtbot) -> None:
    host = QWidget()
    layout = QVBoxLayout(host)
    button = _create_info_button("Source help", host)
    source = QLabel("Source", host)
    unrelated = QLabel("Other", host)
    unrelated.setToolTip("Other help")
    button.watch_tooltip_source(source)
    for widget in (button, source, unrelated):
        layout.addWidget(widget)
    qtbot.addWidget(host)
    host.show()
    QApplication.sendEvent(source, QEvent(QEvent.Type.ToolTip))
    popup = button._hover_help
    assert popup is not None
    QApplication.sendEvent(unrelated, QEvent(QEvent.Type.ToolTip))
    assert button._hover_help is popup
    text = popup.findChild(QLabel, "settingsHelpPopupText")
    assert text is not None and text.text() == "Source help"
    button._close_hover_help()


@pytest.mark.parametrize("delete_owner", [False, True])
def test_hover_sources_and_pending_timers_survive_destruction(qapp, qtbot, delete_owner) -> None:
    import shiboken6

    host = QWidget()
    layout = QVBoxLayout(host)
    button = _create_info_button("Source help", host)
    source = QLabel("Source", host)
    button.watch_tooltip_source(source)
    button.watch_tooltip_source(source)
    assert len(button._help_sources) == 1
    layout.addWidget(button)
    layout.addWidget(source)
    host.show()
    qapp.processEvents()
    QApplication.sendEvent(source, QEvent(QEvent.Type.ToolTip))
    assert button._hover_help is not None
    QCursor.setPos(button.screen().availableGeometry().bottomRight())
    QApplication.sendEvent(source, QEvent(QEvent.Type.Leave))
    button._hover_focus_timer.start()
    if delete_owner:
        shiboken6.delete(host)
        assert not shiboken6.isValid(button)
    else:
        shiboken6.delete(source)
        assert not button._help_sources
        button._close_hover_if_outside()
        assert button._hover_help is None
        host.close()
        shiboken6.delete(host)
    qtbot.wait(230)


def test_source_can_outlive_its_info_button(qapp, qtbot) -> None:
    import shiboken6

    source = QLabel("Source")
    button = _create_info_button("Source help")
    button.watch_tooltip_source(source)
    shiboken6.delete(button)
    shiboken6.delete(source)
    qapp.processEvents()

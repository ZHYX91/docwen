"""Wheel gestures scroll forms; deliberate keyboard and step actions still work."""

import pytest
from PySide6.QtCore import QPoint, QPointF, Qt
from PySide6.QtGui import QWheelEvent
from PySide6.QtWidgets import QApplication, QScrollArea, QVBoxLayout, QWidget

from docwen_gui.widgets.value_controls import ScrollSafeComboBox, ScrollSafeDoubleSpinBox, ScrollSafeSpinBox

pytestmark = pytest.mark.gui


@pytest.mark.parametrize("kind", [ScrollSafeComboBox, ScrollSafeSpinBox, ScrollSafeDoubleSpinBox])
def test_wheel_does_not_change_a_value_but_keyboard_does(qtbot, kind):
    scroll = QScrollArea()
    qtbot.addWidget(scroll)
    page = QWidget()
    page.setMinimumHeight(900)
    layout = QVBoxLayout(page)
    control = kind(page)
    if isinstance(control, ScrollSafeComboBox):
        control.addItems(["First", "Second", "Third"])
        control.setCurrentIndex(1)
        read = control.currentIndex
    else:
        control.setValue(30)
        read = control.value
    layout.addWidget(control)
    layout.addStretch()
    scroll.setWidget(page)
    scroll.resize(350, 200)
    scroll.show()
    control.setFocus()
    original = read()
    event = QWheelEvent(
        QPointF(10, 10),
        QPointF(control.mapToGlobal(QPoint(10, 10))),
        QPoint(),
        QPoint(0, -120),
        Qt.MouseButton.NoButton,
        Qt.KeyboardModifier.NoModifier,
        Qt.ScrollPhase.NoScrollPhase,
        False,
    )
    QApplication.sendEvent(control, event)
    assert read() == original
    assert not event.isAccepted()
    qtbot.keyClick(control, Qt.Key.Key_Down)
    assert read() != original

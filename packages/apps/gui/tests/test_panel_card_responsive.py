"""Responsive layout tests for shared panel-card primitives."""

import pytest
from PySide6.QtWidgets import QApplication, QBoxLayout, QLineEdit

from docwen_gui.widgets.panel_card import FormRow

pytestmark = pytest.mark.gui


def test_form_row_stacks_when_label_and_control_do_not_fit(qapp: QApplication) -> None:
    control = QLineEdit("200")
    control.setMinimumWidth(160)
    row = FormRow("A translated option label that needs room", control)
    row.resize(220, 80)
    row.show()
    qapp.processEvents()

    assert row.content_layout.direction() == QBoxLayout.Direction.TopToBottom
    row.close()


def test_form_row_uses_horizontal_layout_when_space_is_available(qapp: QApplication) -> None:
    control = QLineEdit("200")
    control.setMinimumWidth(160)
    row = FormRow("Target size", control)
    row.resize(640, 40)
    row.show()
    qapp.processEvents()

    assert row.content_layout.direction() == QBoxLayout.Direction.LeftToRight
    row.close()

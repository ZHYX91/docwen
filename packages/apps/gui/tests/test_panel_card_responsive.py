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


def test_compound_size_field_stacks_before_numeric_text_is_clipped(qapp: QApplication) -> None:
    from PySide6.QtWidgets import QComboBox, QHBoxLayout, QWidget

    control = QWidget()
    layout = QHBoxLayout(control)
    layout.setContentsMargins(0, 0, 0, 0)
    layout.setSpacing(8)
    number = QLineEdit("200")
    unit = QComboBox()
    unit.addItems(["KB", "MB"])
    unit.setMinimumWidth(84)
    layout.addWidget(number, 1)
    layout.addWidget(unit)
    row = FormRow("文件大小上限", control)
    row.resize(274, 80)
    row.show()
    try:
        for _ in range(8):
            qapp.processEvents()
        assert row.content_layout.direction() == QBoxLayout.Direction.TopToBottom
        assert number.width() >= number.fontMetrics().horizontalAdvance(number.text()) + 24
        assert unit.width() >= unit.minimumWidth()
        assert not number.geometry().intersects(unit.geometry())
    finally:
        row.close()


def test_form_row_keeps_long_label_and_help_visible_when_control_stacks(qapp: QApplication) -> None:
    from PySide6.QtWidgets import QToolButton

    control = QLineEdit("value")
    control.setMinimumWidth(160)
    help_button = QToolButton()
    help_button.setFixedSize(18, 18)
    row = FormRow("图文输出位置（图片 + OCR）:", control, label_suffix=help_button)
    row.resize(320, 110)
    row.show()
    try:
        for _ in range(8):
            qapp.processEvents()
        assert row.content_layout.direction() == QBoxLayout.Direction.TopToBottom
        assert row.label.width() >= row.label.fontMetrics().horizontalAdvance(row.label.text())
        assert row.label_container.width() == row.contentsRect().width()
        assert not row.label_container.geometry().intersects(control.geometry())
        assert row.label_container.rect().contains(help_button.geometry())
        row.resize(700, 110)
        for _ in range(8):
            qapp.processEvents()
        assert row.content_layout.direction() == QBoxLayout.Direction.LeftToRight
        assert not row.label_container.geometry().intersects(control.geometry())
    finally:
        row.close()

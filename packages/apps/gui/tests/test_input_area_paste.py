"""Input-area paste button and keyboard scope through real Qt events."""

import pytest
from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QKeySequence
from PySide6.QtTest import QSignalSpy
from PySide6.QtWidgets import QApplication, QLineEdit, QVBoxLayout, QWidget
from shiboken6 import isValid

from docwen_gui.view_models.input_area_vm import InputAreaViewModel
from docwen_gui.view_models.main_window_vm import MainWindowViewModel
from docwen_gui.widgets.input_area import InputArea

pytestmark = pytest.mark.gui


@pytest.fixture
def widget(qapp, qtbot):
    main_vm = MainWindowViewModel(controller=None)
    value = InputArea(view_model=InputAreaViewModel(main_vm=main_vm))
    yield value
    main_vm.cancel_inspection()
    qtbot.waitUntil(lambda: not main_vm.inspection_busy)
    if isValid(value):
        value.deleteLater()


class TestPasteAction:
    def test_button_requests_one_paste(self, widget: InputArea, qtbot) -> None:
        spy = QSignalSpy(widget.paste_requested)

        qtbot.mouseClick(widget.paste_button, Qt.MouseButton.LeftButton)

        assert spy.count() == 1

    def test_split_menu_requests_plain_text_without_normal_paste(self, widget: InputArea, qtbot) -> None:
        normal = QSignalSpy(widget.paste_requested)
        plain = QSignalSpy(widget.paste_plain_text_requested)
        menu = widget.paste_menu_button.menu()
        assert menu is not None
        widget.show()

        def choose_plain_text() -> None:
            qtbot.keyClick(menu, Qt.Key.Key_Down)
            qtbot.keyClick(menu, Qt.Key.Key_Return)

        QTimer.singleShot(0, choose_plain_text)
        qtbot.mouseClick(widget.paste_menu_button, Qt.MouseButton.LeftButton)

        assert normal.count() == 0
        assert plain.count() == 1

    def test_ctrl_v_is_scoped_to_input_area_and_does_not_steal_line_edit_paste(
        self, widget: InputArea, qapp: QApplication, qtbot
    ) -> None:
        container = QWidget()
        layout = QVBoxLayout(container)
        editor = QLineEdit(container)
        layout.addWidget(widget)
        layout.addWidget(editor)
        qtbot.addWidget(container)
        container.show()
        container.activateWindow()
        qapp.processEvents()
        qapp.clipboard().setText("editor text")
        spy = QSignalSpy(widget.paste_requested)

        widget.setFocus()
        qtbot.waitUntil(widget.hasFocus)
        qtbot.keySequence(widget, QKeySequence.StandardKey.Paste)
        qapp.processEvents()
        assert spy.count() == 1

        editor.setFocus()
        qtbot.keySequence(editor, QKeySequence.StandardKey.Paste)
        qapp.processEvents()
        assert editor.text() == "editor text"
        assert spy.count() == 1
        container.close()

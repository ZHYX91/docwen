"""InputArea ViewModel and widget lifetime fixtures."""

from collections.abc import Generator

import pytest
from PySide6.QtWidgets import (
    QApplication,
)
from shiboken6 import isValid

from docwen_gui.view_models.input_area_vm import InputAreaViewModel
from docwen_gui.view_models.main_window_vm import MainWindowViewModel
from docwen_gui.widgets.input_area import (
    InputArea,
)


@pytest.fixture
def main_vm(qtbot) -> Generator[MainWindowViewModel, None, None]:
    vm = MainWindowViewModel(controller=None)
    yield vm
    vm.cancel_inspection()
    qtbot.waitUntil(lambda: not vm.inspection_busy)


@pytest.fixture
def input_vm(main_vm: MainWindowViewModel) -> InputAreaViewModel:
    return InputAreaViewModel(main_vm=main_vm)


@pytest.fixture
def widget(qapp: QApplication, input_vm: InputAreaViewModel) -> "Generator[InputArea, None, None]":
    w = InputArea(view_model=input_vm)
    yield w
    if isValid(w):
        w.deleteLater()

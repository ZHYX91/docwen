"""Smoke tests for ActionArea widget.

These tests validate widget construction, 7 setup mode layouts,
cancel flow, and Enter key scanning. Require a QApplication instance.
"""

from collections.abc import Generator

import pytest
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QFrame,
    QGridLayout,
    QLabel,
    QPushButton,
    QWidget,
)
from tests.support.gui_vm_fakes import FakeMainWindowViewModel

from docwen_gui import numbering_schemes
from docwen_gui.styles.action_area import build_action_area_stylesheet
from docwen_gui.styles.design_tokens import Sizing
from docwen_gui.styles.theme_manager import ThemeManager
from docwen_gui.styles.theme_semantics import get_theme_class_color
from docwen_gui.view_models._optimization_filter import OptimizationChoice, OptimizationChoicesResult
from docwen_gui.view_models.action_area_vm import ActionAreaViewModel
from docwen_gui.widgets.action_area import ActionArea

pytestmark = pytest.mark.gui

_NUMBERING_CONFIG = {
    "settings": {"order": ["gongwen_standard", "legal_standard"]},
    "schemes": {
        "gongwen_standard": {"name": "公文标准"},
        "legal_standard": {"name": "层级数字标准"},
    },
}


@pytest.fixture
def vm() -> ActionAreaViewModel:
    return ActionAreaViewModel(
        FakeMainWindowViewModel({"numbering": {"add": _NUMBERING_CONFIG}})  # type: ignore[arg-type]
    )


@pytest.fixture
def widget(qapp: QApplication, vm: ActionAreaViewModel) -> "Generator[ActionArea, None, None]":
    w = ActionArea(view_model=vm)
    yield w
    w.deleteLater()


def _install_optimization_lookup(monkeypatch: pytest.MonkeyPatch) -> None:
    items_by_category = {
        "document": OptimizationChoice("gongwen", "Gongwen", "gongwen", (), ("locale",)),
        "image": OptimizationChoice("invoice_cn", "Invoice CN", "invoice_cn", (), ("locale",)),
        "layout": OptimizationChoice("invoice_cn", "Invoice CN", "invoice_cn", (), ("locale",)),
    }

    def discover(_controller, *, locale, sources=(), target="md") -> OptimizationChoicesResult:
        del locale, target
        category = sources[0].source_category if sources else ""
        choice = items_by_category.get(category)
        return OptimizationChoicesResult(status="ready", choices=(choice,) if choice is not None else ())

    monkeypatch.setattr("docwen_gui.view_models.action_area_vm.discover_optimization_choices", discover)


def _combo_data(combo) -> list[object]:
    return [combo.itemData(i) for i in range(combo.count())]


def _combo_icon_center_color(combo, index: int) -> str:
    image = combo.itemIcon(index).pixmap(12, 12).toImage()
    return image.pixelColor(image.width() // 2, image.height() // 2).name().upper()


def _grid_position(widget: QWidget) -> tuple[int, int]:
    parent = widget.parentWidget()
    assert parent is not None
    layout = parent.layout()
    assert isinstance(layout, QGridLayout)
    index = layout.indexOf(widget)
    assert index >= 0
    position = layout.getItemPosition(index)
    assert isinstance(position, tuple) and len(position) == 4
    return int(position[0]), int(position[1])


__all__ = (
    "ActionArea",
    "ActionAreaViewModel",
    "QApplication",
    "QCheckBox",
    "QFrame",
    "QGridLayout",
    "QLabel",
    "QPushButton",
    "Qt",
    "Sizing",
    "ThemeManager",
    "_combo_data",
    "_combo_icon_center_color",
    "_grid_position",
    "_install_optimization_lookup",
    "build_action_area_stylesheet",
    "get_theme_class_color",
    "numbering_schemes",
    "pytest",
    "pytestmark",
    "vm",
    "widget",
)

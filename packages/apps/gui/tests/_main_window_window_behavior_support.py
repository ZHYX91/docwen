from __future__ import annotations

from typing import Any

import pytest
from PySide6.QtCore import QPoint, QRect
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QGridLayout, QWidget

from docwen_gui.view_models.interaction import (
    ConversionContext,
    MainWindowUiProjection,
    RightPanelSlot,
)
from docwen_gui.window_geometry import WindowRect

pytestmark = pytest.mark.gui


class _ConfigPort:
    def __init__(self, values: dict[str, object]) -> None:
        self.values = dict(values)
        self.get_calls: list[str] = []
        self.set_calls: list[tuple[str, object]] = []
        self.set_many_calls: list[dict[str, object]] = []

    def get(self, key: str, default: object = None) -> object:
        self.get_calls.append(key)
        return self.values.get(key, default)

    def set(self, key: str, value: object) -> bool:
        self.set_calls.append((key, value))
        self.values[key] = value
        return True

    def set_many(self, values: dict[str, Any]) -> bool:
        copied = dict(values)
        self.set_many_calls.append(copied)
        self.values.update(copied)
        return True

    def snapshot(self) -> dict[str, Any]:
        return {}


class _Controller:
    def __init__(self, values: dict[str, object]) -> None:
        self.config_port = _ConfigPort(
            {
                "gui.window.geometry_schema_version": 2,
                **values,
            }
        )

    def stop(self) -> None:
        return None


def _make_window(qapp, values: dict[str, object]):
    from docwen_gui.main_window import MainWindow
    from docwen_gui.view_models.main_window_vm import MainWindowViewModel

    controller = _Controller(values)
    window = MainWindow(view_model=MainWindowViewModel(controller=controller))  # type: ignore[arg-type]
    qapp.processEvents()
    return window, controller


def _root_grid(window: QWidget) -> QGridLayout:
    central = window.findChild(QWidget, "centralContainer")
    assert central is not None
    layout = central.layout()
    assert isinstance(layout, QGridLayout)
    return layout


_VISIBLE_RIGHT = MainWindowUiProjection(
    left_panel_visible=False,
    right_panel_visible=True,
    right_panel_slot=RightPanelSlot.CONVERSION,
    center_action_visible=True,
    info_area_visible=True,
    conversion_context=ConversionContext(
        category="document",
        current_format="docx",
        file_path="C:/tmp/test.docx",
    ),
    template_context=None,
)

_HIDDEN = MainWindowUiProjection(
    left_panel_visible=False,
    right_panel_visible=False,
    right_panel_slot=RightPanelSlot.NONE,
    center_action_visible=False,
    info_area_visible=True,
    conversion_context=None,
    template_context=None,
)

_VISIBLE_LEFT = MainWindowUiProjection(
    left_panel_visible=True,
    right_panel_visible=False,
    right_panel_slot=RightPanelSlot.NONE,
    center_action_visible=False,
    info_area_visible=True,
    conversion_context=None,
    template_context=None,
)

_VISIBLE_BOTH = MainWindowUiProjection(
    left_panel_visible=True,
    right_panel_visible=True,
    right_panel_slot=RightPanelSlot.CONVERSION,
    center_action_visible=True,
    info_area_visible=True,
    conversion_context=_VISIBLE_RIGHT.conversion_context,
    template_context=None,
)

__all__ = (
    "_HIDDEN",
    "_VISIBLE_BOTH",
    "_VISIBLE_LEFT",
    "_VISIBLE_RIGHT",
    "Any",
    "MainWindowUiProjection",
    "QPoint",
    "QRect",
    "QTest",
    "QWidget",
    "RightPanelSlot",
    "WindowRect",
    "_ConfigPort",
    "_make_window",
    "_root_grid",
    "pytest",
    "pytestmark",
)

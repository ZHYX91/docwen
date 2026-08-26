"""DocWen GUI package with lazy public exports.

Keeping package import lightweight lets ``python -m docwen_gui`` reach the
bundle-owned dependency egress guard before importing Qt or application code.
"""

from __future__ import annotations

from importlib import import_module
from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from docwen_core.version import __version__ as __version__
    from docwen_gui.app import create_main_window as create_main_window
    from docwen_gui.app import create_qapplication as create_qapplication
    from docwen_gui.app import run_gui as run_gui
    from docwen_gui.main_window import DEFAULT_CENTER_WIDTH as DEFAULT_CENTER_WIDTH
    from docwen_gui.main_window import DEFAULT_HEIGHT as DEFAULT_HEIGHT
    from docwen_gui.main_window import MIN_HEIGHT as MIN_HEIGHT
    from docwen_gui.main_window import MIN_WIDTH as MIN_WIDTH
    from docwen_gui.main_window import MainWindow as MainWindow
    from docwen_gui.qt_bridge.event_adapter import EventAdapter as EventAdapter
    from docwen_gui.qt_bridge.task_event_bridge import TaskEventBridge as TaskEventBridge
    from docwen_gui.view_models.main_window_vm import (
        MainWindowViewModel as MainWindowViewModel,
    )

_EXPORTS = {
    "DEFAULT_CENTER_WIDTH": ("docwen_gui.main_window", "DEFAULT_CENTER_WIDTH"),
    "DEFAULT_HEIGHT": ("docwen_gui.main_window", "DEFAULT_HEIGHT"),
    "MIN_HEIGHT": ("docwen_gui.main_window", "MIN_HEIGHT"),
    "MIN_WIDTH": ("docwen_gui.main_window", "MIN_WIDTH"),
    "EventAdapter": ("docwen_gui.qt_bridge.event_adapter", "EventAdapter"),
    "MainWindow": ("docwen_gui.main_window", "MainWindow"),
    "MainWindowViewModel": ("docwen_gui.view_models.main_window_vm", "MainWindowViewModel"),
    "TaskEventBridge": ("docwen_gui.qt_bridge.task_event_bridge", "TaskEventBridge"),
    "create_main_window": ("docwen_gui.app", "create_main_window"),
    "create_qapplication": ("docwen_gui.app", "create_qapplication"),
    "run_gui": ("docwen_gui.app", "run_gui"),
}

__all__ = [
    "DEFAULT_CENTER_WIDTH",
    "DEFAULT_HEIGHT",
    "MIN_HEIGHT",
    "MIN_WIDTH",
    "EventAdapter",
    "MainWindow",
    "MainWindowViewModel",
    "TaskEventBridge",
    "__version__",
    "create_main_window",
    "create_qapplication",
    "run_gui",
]


def __getattr__(name: str) -> Any:
    if name == "__version__":
        from docwen_core.version import __version__

        return __version__
    target = _EXPORTS.get(name)
    if target is None:
        raise AttributeError(name)
    module_name, attribute = target
    value = getattr(import_module(module_name), attribute)
    globals()[name] = value
    return value

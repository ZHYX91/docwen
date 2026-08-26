"""Test that all docwen_gui modules can be imported."""

import pytest

pytestmark = pytest.mark.unit


def test_gui_importable() -> None:
    """docwen_gui top-level package should be importable."""
    import docwen_gui

    assert docwen_gui.__version__ == "0.9.0"


def test_app_module_import() -> None:
    """app.py should be importable."""
    from docwen_gui.app import (
        create_main_window,
        create_qapplication,
        run_gui,
    )

    assert callable(create_qapplication)
    assert callable(create_main_window)
    assert callable(run_gui)


def test_main_window_module_import() -> None:
    """main_window.py should be importable."""
    from docwen_gui.main_window import MainWindow

    assert MainWindow is not None


def test_view_model_module_import() -> None:
    """ViewModels should be importable."""
    from docwen_gui.view_models.main_window_vm import MainWindowViewModel

    assert MainWindowViewModel is not None


def test_event_adapter_import() -> None:
    """EventAdapter should be importable."""
    from docwen_gui.qt_bridge.event_adapter import EventAdapter

    assert EventAdapter is not None


def test_task_event_bridge_import() -> None:
    """TaskEventBridge should be importable."""
    from docwen_gui.qt_bridge.task_event_bridge import TaskEventBridge

    assert TaskEventBridge is not None


def test_no_plugin_implementation_imports() -> None:
    """GUI modules must not bypass the application/runtime boundary into plugins."""
    import ast
    from pathlib import Path

    gui_src = Path(__file__).parent.parent / "src" / "docwen_gui"
    forbidden_prefixes = ("docwen_plugin_",)

    failures: list[str] = []
    for py_file in gui_src.rglob("*.py"):
        try:
            source = py_file.read_text(encoding="utf-8")
        except Exception:
            continue
        tree = ast.parse(source, filename=str(py_file))
        rel = py_file.relative_to(gui_src)
        for node in ast.walk(tree):
            if isinstance(node, ast.Import):
                for alias in node.names:
                    if alias.name.startswith(forbidden_prefixes):
                        failures.append(f"{rel}: imports {alias.name}")
            elif isinstance(node, ast.ImportFrom) and node.module and node.module.startswith(forbidden_prefixes):
                failures.append(f"{rel}: from {node.module} import ...")

    assert failures == [], "GUI code imports plugin implementations directly:\n" + "\n".join(failures)


def test_settings_models_do_not_export_dead_convert_config() -> None:
    """SettingsConfig is the aggregate root; no empty conversion placeholder model."""
    from docwen_gui import models
    from docwen_gui.models import settings_config

    assert not hasattr(settings_config, "ConvertConfig")
    assert not hasattr(models, "ConvertConfig")


def test_app_module_has_no_stale_i18n_phase_placeholder() -> None:
    """GUI app metadata must not claim i18n is still a future placeholder."""
    import inspect

    from docwen_gui import app

    source = inspect.getsource(app)
    assert "will be replaced by proper i18n" not in source
    assert "phase 5.2" not in source

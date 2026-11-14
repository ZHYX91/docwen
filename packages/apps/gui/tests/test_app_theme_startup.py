"""Application theme startup wiring."""

from __future__ import annotations

from types import SimpleNamespace
from typing import cast

import pytest

pytestmark = pytest.mark.unit


@pytest.mark.parametrize("stored_preset", ["default", "large", "small"])
def test_window_keeps_already_initialized_font_and_theme(qapp, monkeypatch, stored_preset) -> None:
    from docwen_gui.main_window import MainWindow
    from docwen_gui.styles.theme_manager import ThemeManager

    manager = ThemeManager()
    manager.initialize(qapp, "light", font_size_preset=stored_preset)
    monkeypatch.setattr(ThemeManager, "get_instance", classmethod(lambda cls: manager))
    calls = []
    window = SimpleNamespace(
        _view_model=SimpleNamespace(
            controller=SimpleNamespace(config_port=SimpleNamespace(get=lambda key, default: stored_preset))
        ),
        _apply_font_size_preset=lambda *args, **kwargs: calls.append(args),
        _apply_window_opacity=lambda: None,
        _initialize_app_icon=lambda: None,
    )
    try:
        MainWindow._load_initial_preferences(cast(MainWindow, window))
        assert calls == []
        assert manager.get_font_size_preset() == stored_preset
    finally:
        manager.initialize(qapp, "light")
        manager._disconnect_system_theme()


@pytest.mark.parametrize("stored_theme", ["light", "dark", "system"])
def test_initialize_application_theme_reads_injected_config(
    monkeypatch: pytest.MonkeyPatch,
    stored_theme: str,
) -> None:
    from docwen_gui import app as app_module
    from docwen_gui.styles.theme_manager import ThemeManager

    calls: list[tuple[object, object, object]] = []
    manager = SimpleNamespace(
        initialize=lambda app, theme, **appearance: calls.append((app, theme, appearance)),
    )
    monkeypatch.setattr(ThemeManager, "get_instance", classmethod(lambda cls: manager))

    qt_app = object()
    config_port = SimpleNamespace(
        get=lambda key, default=None: stored_theme if key == "gui.theme.default_theme" else default
    )
    controller = SimpleNamespace(config_port=config_port)

    app_module._initialize_application_theme(qt_app, controller)

    assert calls == [(qt_app, stored_theme, {"scale": 100, "font_size_preset": "default"})]


@pytest.mark.parametrize("stored_theme", [None, "", "sepia", 42])
def test_initialize_application_theme_falls_back_for_invalid_config(
    monkeypatch: pytest.MonkeyPatch,
    stored_theme: object,
) -> None:
    from docwen_gui import app as app_module
    from docwen_gui.styles.theme_manager import ThemeManager

    calls: list[str] = []
    manager = SimpleNamespace(
        initialize=lambda app, theme, **appearance: calls.append(theme),
        apply_ui_scale=lambda value: None,
        apply_font_size_preset=lambda value: None,
    )
    monkeypatch.setattr(ThemeManager, "get_instance", classmethod(lambda cls: manager))

    config_port = SimpleNamespace(
        get=lambda key, default=None: stored_theme if key == "gui.theme.default_theme" else default
    )
    app_module._initialize_application_theme(object(), SimpleNamespace(config_port=config_port))

    assert calls == ["light"]

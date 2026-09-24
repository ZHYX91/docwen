"""Application theme startup wiring."""

from __future__ import annotations

from types import SimpleNamespace

import pytest

pytestmark = pytest.mark.unit


@pytest.mark.parametrize("stored_theme", ["light", "dark", "system"])
def test_initialize_application_theme_reads_injected_config(
    monkeypatch: pytest.MonkeyPatch,
    stored_theme: str,
) -> None:
    from docwen_gui import app as app_module
    from docwen_gui.styles.theme_manager import ThemeManager

    calls: list[tuple[object, object]] = []
    manager = SimpleNamespace(
        initialize=lambda app, theme: calls.append((app, theme)),
        apply_ui_scale=lambda value: calls.append(("scale", value)),
        apply_font_size_preset=lambda value: calls.append(("font", value)),
    )
    monkeypatch.setattr(ThemeManager, "get_instance", classmethod(lambda cls: manager))

    qt_app = object()
    config_port = SimpleNamespace(
        get=lambda key, default=None: stored_theme if key == "gui.theme.default_theme" else default
    )
    controller = SimpleNamespace(config_port=config_port)

    app_module._initialize_application_theme(qt_app, controller)

    assert calls == [(qt_app, stored_theme), ("scale", 100), ("font", "default")]


@pytest.mark.parametrize("stored_theme", [None, "", "sepia", 42])
def test_initialize_application_theme_falls_back_for_invalid_config(
    monkeypatch: pytest.MonkeyPatch,
    stored_theme: object,
) -> None:
    from docwen_gui import app as app_module
    from docwen_gui.styles.theme_manager import ThemeManager

    calls: list[str] = []
    manager = SimpleNamespace(
        initialize=lambda app, theme: calls.append(theme),
        apply_ui_scale=lambda value: None,
        apply_font_size_preset=lambda value: None,
    )
    monkeypatch.setattr(ThemeManager, "get_instance", classmethod(lambda cls: manager))

    config_port = SimpleNamespace(
        get=lambda key, default=None: stored_theme if key == "gui.theme.default_theme" else default
    )
    app_module._initialize_application_theme(object(), SimpleNamespace(config_port=config_port))

    assert calls == ["light"]

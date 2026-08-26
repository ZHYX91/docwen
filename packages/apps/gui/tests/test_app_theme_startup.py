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

    calls: list[tuple[object, str]] = []
    manager = SimpleNamespace(initialize=lambda app, theme: calls.append((app, theme)))
    monkeypatch.setattr(ThemeManager, "get_instance", classmethod(lambda cls: manager))

    qt_app = object()
    config_port = SimpleNamespace(get=lambda key, default=None: stored_theme)
    controller = SimpleNamespace(config_port=config_port)

    app_module._initialize_application_theme(qt_app, controller)

    assert calls == [(qt_app, stored_theme)]


@pytest.mark.parametrize("stored_theme", [None, "", "sepia", 42])
def test_initialize_application_theme_falls_back_for_invalid_config(
    monkeypatch: pytest.MonkeyPatch,
    stored_theme: object,
) -> None:
    from docwen_gui import app as app_module
    from docwen_gui.styles.theme_manager import ThemeManager

    calls: list[str] = []
    manager = SimpleNamespace(initialize=lambda app, theme: calls.append(theme))
    monkeypatch.setattr(ThemeManager, "get_instance", classmethod(lambda cls: manager))

    config_port = SimpleNamespace(get=lambda key, default=None: stored_theme)
    app_module._initialize_application_theme(object(), SimpleNamespace(config_port=config_port))

    assert calls == ["light"]

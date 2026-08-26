"""High-DPI and semantic icon fidelity regressions."""

from __future__ import annotations

import pytest
from PySide6.QtCore import QSize
from PySide6.QtGui import QIcon

pytestmark = pytest.mark.gui


def test_svg_icon_keeps_two_x_backing_pixels(qapp) -> None:
    from docwen_gui.resources import load_svg_icon

    icon = load_svg_icon("about.svg")
    assert isinstance(icon, QIcon)
    assert not icon.isNull()

    pixmap = icon.pixmap(QSize(18, 18), 2.0)
    assert pixmap.width() == 36
    assert pixmap.height() == 36
    assert pixmap.devicePixelRatio() == pytest.approx(2.0)
    assert pixmap.deviceIndependentSize() == QSize(18, 18)


def test_settings_tabs_own_distinct_semantic_icons(qapp) -> None:
    from docwen_gui.widgets.settings.dialog import TAB_KEYS, SettingsDialog

    icons = {key: SettingsDialog._load_tab_icon(key) for key in TAB_KEYS}  # pyright: ignore[reportPrivateUsage]
    assert all(isinstance(icon, QIcon) and not icon.isNull() for icon in icons.values())
    assert len({icon.cacheKey() for icon in icons.values() if icon is not None}) == len(TAB_KEYS)


def test_settings_info_affordance_uses_dedicated_crisp_asset(qapp, monkeypatch: pytest.MonkeyPatch) -> None:
    from docwen_gui.widgets.settings import base_tab

    calls: list[str] = []

    def _capture(name: str, *, color=None):
        calls.append(name)
        return QIcon()

    monkeypatch.setattr(base_tab, "load_svg_icon", _capture)
    button = base_tab._create_info_button("More information")  # pyright: ignore[reportPrivateUsage]

    assert calls == ["info.svg"]
    assert button.iconSize() == QSize(14, 14)


def test_main_window_bottom_actions_use_twenty_pixel_vector_icons(main_window) -> None:
    assert main_window._font_size_btn.iconSize() == QSize(20, 20)  # pyright: ignore[reportPrivateUsage]
    assert main_window._about_btn.iconSize() == QSize(20, 20)  # pyright: ignore[reportPrivateUsage]
    assert main_window._settings_btn.iconSize() == QSize(20, 20)  # pyright: ignore[reportPrivateUsage]

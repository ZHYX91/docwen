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


def test_existing_svg_icon_follows_palette_changes_and_preserves_explicit_color(qapp) -> None:
    from PySide6.QtGui import QColor, QPalette

    from docwen_gui.resources import load_svg_icon

    original = qapp.palette()
    adaptive = load_svg_icon("info.svg")
    fixed = load_svg_icon("info.svg", color="#ff0000")
    assert isinstance(adaptive, QIcon)
    assert isinstance(fixed, QIcon)
    try:
        for color in ("#0f172a", "#f1f5f9", "#0f172a"):
            palette = QPalette(original)
            palette.setColor(QPalette.ColorRole.WindowText, QColor(color))
            qapp.setPalette(palette)
            for icon, expected in ((adaptive, color), (fixed, "#ff0000")):
                raster = icon.pixmap(QSize(24, 24), 2.0).toImage()
                opaque = [
                    raster.pixelColor(x, y).name()
                    for y in range(raster.height())
                    for x in range(raster.width())
                    if raster.pixelColor(x, y).alpha() == 255
                ]
                assert opaque
                assert set(opaque) == {expected}
    finally:
        qapp.setPalette(original)


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


def test_location_button_prefers_application_open_folder_svg(qapp, monkeypatch: pytest.MonkeyPatch) -> None:
    from docwen_gui.widgets import location_button as location_module

    calls: list[str] = []

    def load(name: str, *, color=None):
        calls.append(name)
        return QIcon()

    monkeypatch.setattr(location_module, "load_svg_icon", load)
    button = location_module.LocationButton(label="Locate")

    assert calls == ["open_folder.svg"]
    assert button.iconSize() == QSize(20, 20)


@pytest.mark.parametrize("asset", ["copy.svg", "delete.svg", "move_up.svg", "move_down.svg", "refresh.svg"])
def test_action_icons_use_disabled_palette_for_disabled_controls(qapp, asset: str) -> None:
    from PySide6.QtGui import QColor, QPalette
    from PySide6.QtWidgets import QPushButton

    from docwen_gui.resources import set_action_icon

    original = qapp.palette()
    button = QPushButton("Action")
    set_action_icon(button, asset)
    try:
        for foreground, disabled in (("#172331", "#8899aa"), ("#f1f5f9", "#64748b")):
            palette = QPalette(original)
            palette.setColor(QPalette.ColorGroup.Active, QPalette.ColorRole.WindowText, QColor(foreground))
            palette.setColor(QPalette.ColorGroup.Disabled, QPalette.ColorRole.WindowText, QColor(disabled))
            qapp.setPalette(palette)
            for enabled, expected in ((True, foreground), (False, disabled)):
                button.setEnabled(enabled)
                mode = QIcon.Mode.Normal if button.isEnabled() else QIcon.Mode.Disabled
                raster = button.icon().pixmap(QSize(18, 18), 2.0, mode).toImage()
                opaque = {
                    raster.pixelColor(x, y).name()
                    for y in range(raster.height())
                    for x in range(raster.width())
                    if raster.pixelColor(x, y).alpha() == 255
                }
                assert opaque == {expected}
        assert button.text() == "Action"
    finally:
        qapp.setPalette(original)
        button.deleteLater()


def test_narrow_action_button_keeps_icon_separate_from_wrapped_caption(qapp, qtbot) -> None:
    from docwen_gui.resources import load_svg_icon
    from docwen_gui.widgets.action_button import ActionButton

    button = ActionButton("Copy and edit this template")
    qtbot.addWidget(button)
    icon = load_svg_icon("copy.svg", color="#ff0000")
    assert isinstance(icon, QIcon)
    button.setIcon(icon)
    button.setIconSize(QSize(18, 18))
    button.resize(130, button.heightForWidth(130))
    button.show()
    qapp.processEvents()
    raster = button.grab().toImage()
    red_pixels = [
        (x, y)
        for y in range(raster.height())
        for x in range(raster.width())
        if (color := raster.pixelColor(x, y)).red() - max(color.green(), color.blue()) > 80
    ]
    assert red_pixels
    assert max(x for x, _y in red_pixels) < raster.width() // 3
    assert button.height() >= button.heightForWidth(button.width())

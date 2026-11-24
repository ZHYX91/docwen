"""High-DPI and semantic icon fidelity regressions."""

from __future__ import annotations

import pytest
from PySide6.QtCore import QSize, Qt
from PySide6.QtGui import QFont, QIcon

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
    assert button.iconSize() == QSize(17, 17)


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


@pytest.mark.parametrize(
    "asset", ["copy.svg", "clear.svg", "chevron_down.svg", "delete.svg", "move_up.svg", "move_down.svg", "refresh.svg"]
)
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


@pytest.mark.parametrize(
    "direction", [Qt.LayoutDirection.LeftToRight, Qt.LayoutDirection.RightToLeft], ids=["ltr", "rtl"]
)
def test_narrow_action_button_keeps_icon_separate_from_wrapped_caption(
    qapp, qtbot, direction: Qt.LayoutDirection
) -> None:
    from docwen_gui.resources import load_svg_icon
    from docwen_gui.widgets.action_button import ActionButton

    button = ActionButton("Copy and edit this template")
    qtbot.addWidget(button)
    button.setLayoutDirection(direction)
    # LCD text fringes can satisfy the red-icon mask on some font backends.
    # Keep caption antialiasing grayscale for this color-separation assertion.
    font = button.font()
    font.setStyleStrategy(font.styleStrategy() | QFont.StyleStrategy.NoSubpixelAntialias)
    button.setFont(font)
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
    if direction == Qt.LayoutDirection.LeftToRight:
        assert max(x for x, _y in red_pixels) < raster.width() // 3
    else:
        assert min(x for x, _y in red_pixels) > raster.width() * 2 // 3
    assert button.height() >= button.heightForWidth(button.width())


@pytest.mark.parametrize("theme", ["light", "dark"])
def test_arrow_uses_theme_alpha_in_painted_and_high_dpi_icons(qapp, theme) -> None:
    from PySide6.QtCore import QRect
    from PySide6.QtGui import QImage, QPainter, QPalette

    from docwen_gui.resources import load_svg_icon
    from docwen_gui.styles.theme_manager import ThemeManager

    original = qapp.palette()
    palette = ThemeManager._build_application_palette(theme)
    icon = load_svg_icon("move_up.svg")
    assert isinstance(icon, QIcon)
    try:
        qapp.setPalette(palette)
        for mode, group in (
            (QIcon.Mode.Normal, QPalette.ColorGroup.Active),
            (QIcon.Mode.Disabled, QPalette.ColorGroup.Disabled),
        ):
            expected = palette.color(group, QPalette.ColorRole.WindowText)
            painted = QImage(24, 24, QImage.Format.Format_ARGB32_Premultiplied)
            painted.fill(Qt.GlobalColor.transparent)
            painter = QPainter(painted)
            icon.paint(painter, QRect(0, 0, 24, 24), Qt.AlignmentFlag.AlignCenter, mode)
            painter.end()
            for raster in (painted, icon.pixmap(QSize(24, 24), 2.0, mode).toImage()):
                peak = max(
                    raster.pixelColor(x, y).alpha() for x in range(raster.width()) for y in range(raster.height())
                )
                assert peak == pytest.approx(expected.alpha(), abs=1)
    finally:
        qapp.setPalette(original)


def test_disabled_theme_paint_preserves_authored_svg_colors_and_opacity(qapp, tmp_path, monkeypatch) -> None:
    from PySide6.QtGui import QColor, QPalette

    from docwen_gui import resources

    asset = tmp_path / "mixed.svg"
    asset.write_text(
        '<svg xmlns="http://www.w3.org/2000/svg" width="30" height="10">'
        '<rect width="10" height="10" fill="#000" fill-opacity="0.5"/>'
        '<path d="M 10 5 H 20" stroke="black" stroke-width="4"/>'
        '<rect x="20" width="10" height="10" fill="#3b82f6"/></svg>',
        encoding="utf-8",
    )
    monkeypatch.setattr(resources, "asset_path", lambda _path: asset)
    icon = resources.load_svg_asset_icon(str(asset))
    assert isinstance(icon, QIcon)
    original = qapp.palette()
    try:
        palette = QPalette(original)
        palette.setColor(QPalette.ColorGroup.Disabled, QPalette.ColorRole.WindowText, QColor(30, 60, 90, 128))
        qapp.setPalette(palette)
        raster = icon.pixmap(QSize(30, 10), 1.0, QIcon.Mode.Disabled).toImage()
        assert raster.pixelColor(5, 5).alpha() == pytest.approx(64, abs=1)
        assert raster.pixelColor(15, 5).getRgb() == pytest.approx((30, 60, 90, 128), abs=2)
        assert raster.pixelColor(25, 5) == QColor("#3b82f6")
    finally:
        qapp.setPalette(original)

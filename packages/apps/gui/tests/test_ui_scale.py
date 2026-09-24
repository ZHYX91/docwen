"""Live content scale, geometry ownership and settings transaction regressions."""

from pathlib import Path

import pytest
from PySide6.QtCore import QSize
from PySide6.QtWidgets import QPushButton, QVBoxLayout, QWidget

from docwen_gui.styles.theme_manager import ThemeManager
from docwen_gui.styles.ui_scale import set_metric

pytestmark = pytest.mark.gui


def test_native_color_scheme_tracks_theme_and_releases_system_override(qapp, monkeypatch) -> None:
    from PySide6.QtCore import Qt
    from PySide6.QtGui import QStyleHints

    requests = []
    original = QStyleHints.setColorScheme

    def record(hints, scheme):
        requests.append(scheme)
        original(hints, scheme)

    monkeypatch.setattr(QStyleHints, "setColorScheme", record)
    manager = ThemeManager.get_instance()
    manager.initialize(qapp, "light")
    manager.apply_theme("dark")
    manager.apply_theme("system")
    manager.apply_font_size_preset("large")
    manager.apply_ui_scale(125)
    assert requests == [
        Qt.ColorScheme.Light,
        Qt.ColorScheme.Dark,
        Qt.ColorScheme.Unknown,
        Qt.ColorScheme.Unknown,
        Qt.ColorScheme.Unknown,
    ]
    manager.apply_ui_scale(100)
    manager.apply_font_size_preset("default")
    manager.apply_theme("light")


@pytest.mark.parametrize("percent", [90, 100, 110, 125, 150])
@pytest.mark.parametrize("theme", ["light", "dark"])
def test_checkbox_indicator_tracks_scale_and_keeps_native_keyboard(qapp, qtbot, percent, theme) -> None:
    from PySide6.QtCore import Qt
    from PySide6.QtWidgets import QStyle, QStyleOptionButton

    from docwen_gui.widgets.check_box import CheckBox

    manager = ThemeManager.get_instance()
    manager.initialize(qapp, theme)
    manager.apply_ui_scale(percent)
    checkbox = CheckBox("A scalable setting")
    qtbot.addWidget(checkbox)
    checkbox.show()
    checkbox.setFocus()
    qapp.processEvents()
    try:
        option = QStyleOptionButton()
        checkbox.initStyleOption(option)
        indicator = checkbox.style().subElementRect(QStyle.SubElement.SE_CheckBoxIndicator, option, checkbox)
        assert abs(indicator.width() - 18 * percent / 100) <= 1
        assert indicator.height() == indicator.width()
        label = checkbox.style().subElementRect(QStyle.SubElement.SE_CheckBoxContents, option, checkbox)
        assert label.left() > indicator.right()
        with qtbot.waitSignal(checkbox.toggled):
            qtbot.keyClick(checkbox, Qt.Key.Key_Space)
        assert checkbox.isChecked()
        image = checkbox.grab().toImage()
        assert not image.isNull()
    finally:
        manager.apply_ui_scale(100)
        manager.apply_theme("light")


def test_selected_input_keeps_compact_height_after_scale_round_trip(qapp, qtbot) -> None:
    from docwen_gui.view_models.input_area_vm import InputAreaViewModel
    from docwen_gui.view_models.main_window_vm import MainWindowViewModel
    from docwen_gui.widgets.input_area import InputArea

    manager = ThemeManager.get_instance()
    manager.initialize(qapp, "light")
    manager.apply_ui_scale(100)
    vm = InputAreaViewModel(main_vm=MainWindowViewModel(controller=None))
    widget = InputArea(view_model=vm)
    qtbot.addWidget(widget)
    widget.resize(460, 240)
    widget.show()
    vm._emit_message("Selected: example.md", "success")
    qapp.processEvents()
    try:
        initial = widget.minimumHeight()
        assert initial < 200
        for percent in (150, 90, 125, 100):
            manager.apply_ui_scale(percent)
            qapp.processEvents()
            assert widget.minimumHeight() == round(initial * percent / 100)
    finally:
        manager.apply_ui_scale(100)


def test_live_scale_has_no_rounding_drift_and_new_controls_match(qapp) -> None:
    manager = ThemeManager.get_instance()
    manager.initialize(qapp, "light")
    manager.apply_ui_scale(100)
    manager.apply_font_size_preset("default")
    root = QWidget()
    layout = QVBoxLayout(root)
    button = QPushButton("Scale", root)
    layout.addWidget(button)
    set_metric(layout, "setContentsMargins", 12, 8, 12, 8)
    set_metric(button, "setIconSize", QSize(20, 20))
    root.show()
    try:
        for percent in (150, 90, 125, 110, 100) * 3:
            manager.apply_ui_scale(percent)
            qapp.processEvents()
            fresh = QPushButton("New", root)
            set_metric(fresh, "setIconSize", QSize(20, 20))
            fresh.ensurePolished()
            assert button.iconSize() == fresh.iconSize() == QSize(round(20 * percent / 100), round(20 * percent / 100))
            assert layout.contentsMargins().left() == round(12 * percent / 100)
            assert qapp.font().pointSizeF() == pytest.approx(10.5 * percent / 100)
            fresh.deleteLater()
        assert button.iconSize() == QSize(20, 20)
        assert layout.contentsMargins().left() == 12
    finally:
        manager.apply_ui_scale(100)
        root.close()


def test_display_preferences_migrate_once_without_reinterpreting_old_scale(tmp_path: Path) -> None:
    from docwen_bundle.config_port import ConfigPortAdapter

    directory = tmp_path / "configs"
    directory.mkdir()
    path = directory / "gui.toml"
    path.write_text(
        '[dpi]\nui_scale=200\nenable_dpi_scaling=false\n[font]\nsize_preset="xlarge"\n[history]\nmax_recent=23\n',
        encoding="utf-8",
    )
    base = Path(__file__).resolve().parents[4] / "configs"
    port = ConfigPortAdapter(base_dir=base, user_dir=directory)
    assert port.get("gui.appearance.scale_percent") == 100
    assert port.get("gui.font.size_preset") == "large"
    assert port.get("gui.history.max_recent") == 23
    assert "[dpi]" not in path.read_text(encoding="utf-8")
    original = (path.read_bytes(), path.stat().st_mtime_ns)
    ConfigPortAdapter(base_dir=base, user_dir=directory)
    assert (path.read_bytes(), path.stat().st_mtime_ns) == original


def test_settings_apply_then_cancel_restores_last_saved_appearance(qapp, tmp_path: Path, monkeypatch) -> None:
    from docwen_application.controller import ApplicationController
    from docwen_bundle.config_port import ConfigPortAdapter
    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings import dialog as dialog_module
    from docwen_gui.widgets.settings.general_tab import GeneralTab

    manager = ThemeManager.get_instance()
    manager.initialize(qapp, "light")
    manager.apply_ui_scale(100)
    manager.apply_font_size_preset("default")
    base = Path(__file__).resolve().parents[4] / "configs"
    port = ConfigPortAdapter(base_dir=base, user_dir=tmp_path / "configs")
    vm = SettingsViewModel(controller=ApplicationController(config_port=port))
    dialog = dialog_module.SettingsDialog(view_model=vm)
    tab = dialog.findChild(GeneralTab)
    assert tab is not None
    monkeypatch.setattr(dialog_module, "_show_confirm", lambda *args, **kwargs: True)
    try:
        tab._scale_combo.setCurrentIndex(tab._scale_combo.findData(125))
        tab._font_combo.setCurrentIndex(tab._font_combo.findData("large"))
        assert (manager.get_ui_scale(), manager.get_font_size_preset()) == (125, "large")
        assert port.get("gui.appearance.scale_percent") == 100
        dialog._on_apply()
        assert port.get("gui.appearance.scale_percent") == 125
        assert port.get("gui.font.size_preset") == "large"
        tab._scale_combo.setCurrentIndex(tab._scale_combo.findData(90))
        tab._font_combo.setCurrentIndex(tab._font_combo.findData("small"))
        dialog.reject()
        assert (manager.get_ui_scale(), manager.get_font_size_preset()) == (125, "large")
        assert port.get("gui.appearance.scale_percent") == 125
    finally:
        dialog.close()
        manager.apply_ui_scale(100)
        manager.apply_font_size_preset("default")

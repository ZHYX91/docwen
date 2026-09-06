from __future__ import annotations

import pytest

pytestmark = pytest.mark.gui


@pytest.mark.parametrize("locale,font_size", [("zh_CN", 12), ("en_US", 15)])
def test_settings_cards_keep_field_columns_aligned_across_all_tabs(qapp, qtbot, locale, font_size) -> None:
    from PySide6.QtCore import QPoint
    from PySide6.QtGui import QFont
    from PySide6.QtWidgets import QWidget

    from docwen_gui.i18n import get_locale, set_locale
    from docwen_gui.models.settings_config import SettingsConfig
    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings.base_tab import _SettingsFormLayout
    from docwen_gui.widgets.settings.dialog import _TAB_SPECS

    previous = get_locale()
    set_locale(locale)
    try:
        for key, spec in _TAB_SPECS.items():
            tab = spec.factory(SettingsViewModel(config=SettingsConfig()))
            assert isinstance(tab, QWidget)
            tab.setFont(QFont("Microsoft YaHei", font_size))
            tab.show()
            try:
                wide_heights = None
                for width in (700, 420, 700):
                    tab.resize(width, 850)
                    for _ in range(12):
                        qapp.processEvents()
                    qtbot.wait(50)
                    for form in tab.findChildren(_SettingsFormLayout):
                        rows = form.field_rows
                        if not rows:
                            continue
                        assert len({row.control.x() for row in rows}) == 1, (key, width)
                        assert len({row.control.width() for row in rows}) == 1, (key, width)
                        for row in rows:
                            assert not row.label_container.geometry().intersects(row.control.geometry()), key
                            assert row.control.geometry().right() < row.width(), key
                            assert row.label.height() >= row.label.heightForWidth(row.label.width()), key
                    all_rows = [row for form in tab.findChildren(_SettingsFormLayout) for row in form.field_rows]
                    if all_rows:
                        assert len({row.control.mapTo(tab, QPoint()).x() for row in all_rows}) == 1, key
                        assert len({row.control.width() for row in all_rows}) == 1, key
                    if width == 700:
                        current_heights = [row.height() for row in all_rows]
                        if wide_heights is not None:
                            assert current_heights == wide_heights, key
                        wide_heights = current_heights
            finally:
                tab.close()
                tab.deleteLater()
                qapp.processEvents()
    finally:
        set_locale(previous)


def test_settings_checkbox_wraps_without_losing_native_mouse_keyboard_or_accessibility(qapp, qtbot) -> None:
    from PySide6.QtCore import Qt
    from PySide6.QtGui import QFont
    from PySide6.QtWidgets import QCheckBox

    from docwen_gui.widgets.settings.check_box import SettingsCheckBox

    text = "Number Suite captions and cross-references 题注与交叉引用"
    checkbox = SettingsCheckBox(text)
    qtbot.addWidget(checkbox)
    checkbox.setFont(QFont("Microsoft YaHei", 15))
    checkbox.resize(240, checkbox.heightForWidth(240))
    checkbox.show()
    qapp.processEvents()
    assert "\n" in QCheckBox.text(checkbox)
    assert checkbox.text() == checkbox.accessibleName() == text
    qtbot.mouseClick(checkbox, Qt.MouseButton.LeftButton)
    assert checkbox.isChecked()
    qtbot.keyClick(checkbox, Qt.Key.Key_Space)
    assert not checkbox.isChecked()
    wide = checkbox.fontMetrics().horizontalAdvance(text) + 100
    checkbox.resize(wide, checkbox.heightForWidth(wide))
    qapp.processEvents()
    assert QCheckBox.text(checkbox) == text


@pytest.mark.parametrize("locale", ["zh_CN", "en_US"])
@pytest.mark.parametrize("tab_name", ["export", "logging"])
def test_help_checkbox_uses_available_row_width(qapp, qtbot, locale, tab_name) -> None:
    from PySide6.QtWidgets import QCheckBox, QToolButton, QWidget

    from docwen_gui.i18n import get_locale, set_locale
    from docwen_gui.models.settings_config import SettingsConfig
    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings.check_box import SettingsCheckBox
    from docwen_gui.widgets.settings.dialog import _TAB_SPECS

    previous = get_locale()
    set_locale(locale)
    tab = _TAB_SPECS[tab_name].factory(SettingsViewModel(config=SettingsConfig()))
    assert isinstance(tab, QWidget)
    qtbot.addWidget(tab)
    try:
        tab.show()
        for width in (700, 420, 700):
            tab.resize(width, 850)
            qtbot.wait(80)
            for checkbox in tab.findChildren(SettingsCheckBox):
                wrapper = checkbox.parentWidget()
                assert wrapper is not None
                icons = wrapper.findChildren(QToolButton, "settingsInfoButton")
                if not icons:
                    continue
                icon = icons[0]
                assert checkbox.width() >= wrapper.width() - icon.width() - 8
                assert not checkbox.geometry().intersects(icon.geometry())
                assert checkbox.height() >= checkbox.heightForWidth(checkbox.width())
                if width == 700:
                    assert QCheckBox.text(checkbox) == checkbox.text()
    finally:
        tab.close()
        set_locale(previous)

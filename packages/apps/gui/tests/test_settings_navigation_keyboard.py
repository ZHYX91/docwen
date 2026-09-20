"""Exercise the real settings sidebar and compact selector through Qt key events."""

from __future__ import annotations

import pytest
from PySide6.QtCore import Qt
from PySide6.QtGui import QAccessible, QPalette
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QComboBox, QWidget

from docwen_gui.i18n import t
from docwen_gui.view_models.settings_vm import SettingsViewModel
from docwen_gui.widgets.settings.dialog import TAB_KEYS, TAB_NAMES, SettingsDialog

pytestmark = pytest.mark.gui


@pytest.fixture
def dialog(qapp):
    window = SettingsDialog(view_model=SettingsViewModel())
    window.resize(960, 800)
    window.show()
    window.activateWindow()
    qapp.processEvents()
    yield window
    window.close()
    qapp.processEvents()


def _item(dialog: SettingsDialog, key: str) -> QWidget:
    item = dialog.findChild(QWidget, f"settingsNavItem_{key}")
    assert item is not None
    return item


def test_all_sidebar_pages_and_language_expose_accessible_names(dialog) -> None:
    for key in TAB_KEYS:
        interface = QAccessible.queryAccessibleInterface(_item(dialog, key))
        assert interface is not None
        assert interface.text(QAccessible.Text.Name) == TAB_NAMES[key]
    language = dialog.findChild(QComboBox, "generalLanguageCombo")
    assert language is not None
    assert language.accessibleName() == t("settings.general.language_label")


def test_arrows_keep_sidebar_focus_enter_does_not_submit_and_tab_enters_content(dialog, qapp) -> None:
    first = _item(dialog, "general")
    frame = dialog.findChild(QWidget, "settingsNavigationFocusFrame")
    assert frame is not None and not frame.isVisible()
    first.setFocus(Qt.FocusReason.TabFocusReason)
    qapp.processEvents()
    assert frame.isVisible() and frame.parentWidget() is first
    focus_image = frame.grab().toImage()
    highlight = frame.palette().color(QPalette.ColorRole.Highlight)
    assert any(focus_image.pixelColor(focus_image.width() // 2, y) == highlight for y in range(4))
    QTest.keyClick(first, Qt.Key.Key_Down)
    assert dialog.current_section() == "text"
    assert qapp.focusWidget() is _item(dialog, "text")
    assert frame.isVisible() and frame.parentWidget() is _item(dialog, "text")
    assert frame.geometry() == _item(dialog, "text").rect()
    assert not first.focusPolicy() & Qt.FocusPolicy.TabFocus
    QTest.keyClick(_item(dialog, "text"), Qt.Key.Key_End)
    assert dialog.current_section() == "logging"
    assert qapp.focusWidget() is _item(dialog, "logging")
    QTest.keyClick(_item(dialog, "logging"), Qt.Key.Key_Return)
    assert dialog.isVisible()
    assert qapp.focusWidget() is _item(dialog, "logging")
    QTest.keyClick(_item(dialog, "logging"), Qt.Key.Key_Home)
    assert dialog.current_section() == "general"
    QTest.keyClick(first, Qt.Key.Key_Tab)
    assert qapp.focusWidget() is dialog.findChild(QComboBox, "generalLanguageCombo")
    assert not frame.isVisible()
    QTest.keyClick(qapp.focusWidget(), Qt.Key.Key_Tab, Qt.KeyboardModifier.ShiftModifier)
    assert qapp.focusWidget() is first


def test_compact_selector_can_traverse_multiple_pages_without_focus_escaping(dialog, qapp) -> None:
    dialog.resize(440, 600)
    qapp.processEvents()
    selector = dialog.findChild(QComboBox, "settingsPageSelector")
    assert selector is not None and selector.isVisible()
    selector.setFocus(Qt.FocusReason.TabFocusReason)
    QTest.keyClick(selector, Qt.Key.Key_Down)
    assert dialog.current_section() == "text"
    assert qapp.focusWidget() is selector
    QTest.keyClick(selector, Qt.Key.Key_Down)
    assert dialog.current_section() == "document"
    assert qapp.focusWidget() is selector
    assert selector.accessibleName() == t("settings.title")


def test_short_wide_window_keeps_all_pages_reachable_and_restores_sidebar(dialog, qapp) -> None:
    dialog.resize(960, 450)
    qapp.processEvents()
    selector = dialog.findChild(QComboBox, "settingsPageSelector")
    assert selector is not None and selector.isVisible()
    assert dialog._navigation.isHidden()
    selector.setFocus(Qt.FocusReason.TabFocusReason)
    QTest.keyClick(selector, Qt.Key.Key_End)
    assert dialog.current_section() == "logging"
    assert qapp.focusWidget() is selector
    assert selector.count() == len(TAB_KEYS)
    dialog.resize(960, 800)
    qapp.processEvents()
    assert selector.isHidden()
    assert dialog._navigation.isVisible()
    assert dialog.current_section() == "logging"

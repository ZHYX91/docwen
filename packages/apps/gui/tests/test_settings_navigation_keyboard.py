"""Exercise the real settings sidebar and compact selector through Qt key events."""

from __future__ import annotations

import pytest
from PySide6.QtCore import QSignalBlocker, Qt
from PySide6.QtGui import QAccessible, QFont, QPalette
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QComboBox, QFrame, QLabel, QPushButton, QWidget

from docwen_gui.i18n import get_locale, set_locale, t
from docwen_gui.styles.settings import build_settings_stylesheet
from docwen_gui.view_models.settings_vm import SettingsViewModel
from docwen_gui.widgets.settings import dialog as settings_dialog
from docwen_gui.widgets.settings.dialog import TAB_KEYS, TAB_NAMES, SettingsDialog, SettingsNavigationGroup

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


def _item(dialog: SettingsDialog, key: str) -> QPushButton:
    item = dialog.findChild(QPushButton, f"settingsNavItem_{key}")
    assert item is not None
    return item


def _item_is_in_view(dialog: SettingsDialog, key: str) -> bool:
    sidebar = dialog._navigation
    assert sidebar is not None
    item = _item(dialog, key)
    viewport = sidebar.viewport()
    return all(
        viewport.rect().contains(item.mapTo(viewport, point))
        for point in (item.rect().topLeft(), item.rect().bottomRight())
    )


def test_all_sidebar_pages_and_language_expose_accessible_names(dialog) -> None:
    for key in TAB_KEYS:
        interface = QAccessible.queryAccessibleInterface(_item(dialog, key))
        assert interface is not None
        assert interface.text(QAccessible.Text.Name) == TAB_NAMES[key]
    language = dialog.findChild(QComboBox, "generalLanguageCombo")
    assert language is not None
    assert language.accessibleName() == t("settings.general.language_label")
    group = dialog.findChild(QLabel, "settingsNavGroup_general")
    assert group is not None and group.text() == t("settings.navigation.general")
    assert dialog.findChild(QLabel, "settingsNavGroup_optimizations") is None
    selector = dialog._compact_navigation
    assert [selector.itemData(index) for index in range(selector.count())] == TAB_KEYS


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


def test_short_wide_window_scrolls_selected_page_into_view(dialog, qapp, qtbot) -> None:
    dialog.resize(960, 450)
    qapp.processEvents()
    selector = dialog.findChild(QComboBox, "settingsPageSelector")
    sidebar = dialog._navigation
    assert sidebar is not None and sidebar.isVisible()
    assert selector is not None and selector.isHidden()
    assert sidebar.verticalScrollBar().maximum() > 0
    first = _item(dialog, "general")
    first.setFocus(Qt.FocusReason.TabFocusReason)
    QTest.keyClick(first, Qt.Key.Key_End)
    assert dialog.current_section() == "logging"
    assert qapp.focusWidget() is _item(dialog, "logging")
    qtbot.waitUntil(lambda: _item_is_in_view(dialog, "logging"))
    assert sidebar.verticalScrollBar().value() > 0
    QTest.keyClick(_item(dialog, "logging"), Qt.Key.Key_Home)
    assert dialog.current_section() == "general"
    qtbot.waitUntil(lambda: _item_is_in_view(dialog, "general"))

    dialog.resize(440, 450)
    qapp.processEvents()
    assert selector.isVisible() and sidebar.isHidden()
    selector.setCurrentIndex(selector.findData("logging"))
    assert dialog.current_section() == "logging"
    assert selector.count() == len(TAB_KEYS)
    dialog.resize(960, 450)
    qapp.processEvents()
    assert selector.isHidden()
    assert sidebar.isVisible()
    assert dialog.current_section() == "logging"
    qtbot.waitUntil(lambda: _item_is_in_view(dialog, "logging"))


def test_page_keys_survive_different_selector_and_stack_positions(dialog, qapp, monkeypatch) -> None:
    dialog.resize(440, 600)
    qapp.processEvents()
    selector = dialog._compact_navigation
    with QSignalBlocker(selector):
        selector.insertSeparator(1)
    dialog._tab_widget.tabBar().moveTab(dialog._tab_widget.indexOf(dialog._tabs["logging"]), 0)

    selector.setCurrentIndex(selector.findData("document"))
    assert dialog.current_section() == "document"
    assert dialog._tab_widget.currentWidget() is dialog._tabs["document"]
    assert _item(dialog, "document").isChecked()

    assert dialog.activate_section("logging")
    assert dialog._tab_widget.currentIndex() == 0
    assert selector.currentData() == "logging"
    assert dialog.current_section() == "logging"
    reset_groups = []
    monkeypatch.setattr(settings_dialog, "_show_confirm", lambda *_args, **_kwargs: True)
    monkeypatch.setattr(dialog.view_model, "reset_group", lambda group: reset_groups.append(group) or True)
    monkeypatch.setattr(dialog, "_reload_tab_after_reset", lambda *_args: None)
    dialog._on_reset_tab()
    assert reset_groups == ["logging"]


@pytest.mark.parametrize("locale,font_size", [("zh_CN", 10), ("de_DE", 15), ("ar_SA", 15)])
def test_group_headings_do_not_add_pages_or_keyboard_stops(qapp, qtbot, monkeypatch, locale, font_size) -> None:
    previous_locale = get_locale()
    set_locale(locale)
    monkeypatch.setattr(
        settings_dialog,
        "NAVIGATION_GROUPS",
        (
            SettingsNavigationGroup("general", tuple(TAB_KEYS[:2])),
            SettingsNavigationGroup("optimizations", tuple(TAB_KEYS[2:])),
        ),
    )
    monkeypatch.setattr(
        settings_dialog,
        "TAB_NAMES",
        {key: t("settings.templates.title" if key == "templates" else f"settings.tabs.{key}") for key in TAB_KEYS},
    )
    window = None
    try:
        window = SettingsDialog(view_model=SettingsViewModel())
        window.setFont(QFont("sans-serif", font_size))
        window.setStyleSheet(build_settings_stylesheet("light", "large" if font_size == 15 else "default"))
        if locale == "ar_SA":
            window.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        window.resize(1100, 450)
        window.show()
        window.activateWindow()
        qapp.processEvents()
        sidebar = window._navigation
        selector = window._compact_navigation
        assert sidebar is not None and sidebar.isVisible()
        assert selector.count() == len(TAB_KEYS) == window._tab_widget.count()
        for key in ("general", "optimizations"):
            heading = window.findChild(QLabel, f"settingsNavGroup_{key}")
            assert heading is not None and heading.text() == t(f"settings.navigation.{key}")
            assert heading.focusPolicy() == Qt.FocusPolicy.NoFocus
        separator = window.findChild(QFrame, "settingsNavSeparator_optimizations")
        assert separator is not None and separator.focusPolicy() == Qt.FocusPolicy.NoFocus
        assert separator.height() == 1

        window.activate_section("text")
        text_item = _item(window, "text")
        text_item.setFocus(Qt.FocusReason.TabFocusReason)
        QTest.keyClick(text_item, Qt.Key.Key_Down)
        assert window.current_section() == "document"
        assert qapp.focusWidget() is _item(window, "document")
        assert selector.currentData() == "document"
        qtbot.waitUntil(lambda: _item_is_in_view(window, "document"))

        window.resize(440, 450)
        qapp.processEvents()
        assert selector.isVisible() and sidebar.isHidden()
        selector.setFocus(Qt.FocusReason.TabFocusReason)
        QTest.keyClick(selector, Qt.Key.Key_Up)
        assert window.current_section() == "text"
        assert selector.currentText().startswith(t("settings.navigation.general"))
        QTest.keyClick(selector, Qt.Key.Key_Down)
        assert window.current_section() == "document"
        assert selector.currentText().startswith(t("settings.navigation.optimizations"))
    finally:
        if window is not None:
            window.close()
        set_locale(previous_locale)

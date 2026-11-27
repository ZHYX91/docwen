"""Numbering editors keep their inputs, validation details and insertion menus usable."""

from __future__ import annotations

import pytest
from PySide6.QtGui import QAccessible
from PySide6.QtWidgets import QToolButton

from docwen_gui.i18n import get_locale, set_locale, t
from docwen_gui.styles.theme_manager import ThemeManager
from docwen_gui.widgets.settings.numbering_add_editor import NumberingAddDialog
from docwen_gui.widgets.settings.numbering_clean_editor import NumberingCleanDialog

pytestmark = pytest.mark.gui


@pytest.mark.parametrize("locale", ["de_DE", "zh_CN", "ar_SA"])
def test_placeholder_buttons_identify_and_target_their_heading_level(qapp, qtbot, monkeypatch, locale):
    original_locale = get_locale()
    try:
        set_locale(locale)
        dialog = NumberingAddDialog()
        qtbot.addWidget(dialog)
        selected = []
        monkeypatch.setattr(dialog, "_show_placeholder_menu", selected.append)
        buttons = [b for b in dialog.findChildren(QToolButton) if b.text() == "+"]
        assert len(buttons) == 9
        for level in range(1, 10):
            name = t("editors.numbering_add.insert_placeholder", level=level)
            (button,) = [b for b in buttons if b.accessibleName() == name]
            interface = QAccessible.queryAccessibleInterface(button)
            assert interface is not None and interface.text(QAccessible.Text.Name) == name
            assert button.toolTip() == name
            button.click()
        assert selected == list(range(1, 10))
    finally:
        set_locale(original_locale)


def test_long_regex_diagnostic_wraps_without_expanding_the_dialog(qapp, qtbot, monkeypatch):
    import re

    dialog = NumberingCleanDialog(
        config_data={"rules": [{"id": "sample", "pattern": "^a"}, {"id": "other", "pattern": "^c"}]}
    )
    monkeypatch.setattr(dialog, "_confirm", lambda *_: True)
    qtbot.addWidget(dialog)
    dialog.resize(640, 550)
    dialog.show()
    qapp.processEvents()
    width = dialog.width()
    group = "heading" * 20
    pattern = f"(?P<{group}>a)(?P<{group}>b)"
    with pytest.raises(re.error) as error:
        re.compile(pattern)
    dialog.pattern_edit.setText(pattern)
    qapp.processEvents()
    assert dialog.width() == width
    assert str(error.value) in dialog.test_result.toPlainText()
    assert pattern in dialog.test_result.toPlainText()
    assert dialog.test_result.horizontalScrollBar().maximum() == 0
    assert not dialog.save_btn.isEnabled()
    dialog.rule_list.setCurrentRow(1)
    assert not dialog.save_btn.isEnabled()
    dialog.rule_list.setCurrentRow(0)
    dialog.pattern_edit.setText("^b")
    assert dialog.save_btn.isEnabled()
    assert str(error.value) not in dialog.test_result.toPlainText()
    dialog.pattern_edit.setText("^a")


def test_out_of_range_repeat_count_can_be_corrected(qapp, qtbot):
    dialog = NumberingCleanDialog(config_data={"rules": [{"id": "sample", "pattern": "a"}]})
    qtbot.addWidget(dialog)
    dialog.pattern_edit.setText("a{999999999999}")
    assert not dialog.save_btn.isEnabled()
    assert dialog.status_label.text() == t("editors.numbering_clean.regex_error")
    dialog.pattern_edit.setText("b{2}")
    assert dialog.save_btn.isEnabled()
    dialog.pattern_edit.setText("a")


def test_large_german_scheme_form_keeps_a_whole_placeholder_visible(qapp, qtbot):
    original_locale = get_locale()
    manager = ThemeManager.get_instance()
    try:
        set_locale("de_DE")
        manager.initialize(qapp, "light", scale=150, font_size_preset="large")
        dialog = NumberingAddDialog()
        qtbot.addWidget(dialog)
        dialog.show()
        edit = dialog.level_edits[1]
        for width in (640, 1100, 640):
            dialog.resize(width, 720)
            qapp.processEvents()
            assert edit.width() >= edit.fontMetrics().horizontalAdvance("{1.arabic_half}")
            list_above_form = (
                dialog.scheme_list.mapTo(dialog, dialog.scheme_list.rect().bottomLeft()).y()
                < edit.mapTo(dialog, edit.rect().topLeft()).y()
            )
            assert list_above_form == (width == 640)
    finally:
        manager.initialize(qapp, "light")
        set_locale(original_locale)

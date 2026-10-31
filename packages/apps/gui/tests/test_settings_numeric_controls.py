"""Source Qt numeric geometry and native interaction; not packaged desktop acceptance."""

import pytest
from PySide6.QtCore import QFile, Qt
from PySide6.QtWidgets import (
    QDoubleSpinBox,
    QLineEdit,
    QSpinBox,
    QStyle,
    QStyleOptionFrame,
    QStyleOptionSpinBox,
    QWidget,
)

from docwen_gui.styles.theme_manager import ThemeManager
from docwen_gui.view_models.settings_vm import SettingsViewModel
from docwen_gui.widgets.settings.export_tab import ExportTab
from docwen_gui.widgets.settings.general_tab import GeneralTab
from docwen_gui.widgets.settings.link_tab import LinkTab
from docwen_gui.widgets.settings.logging_tab import LoggingTab
from docwen_gui.widgets.settings.numbering_clean_editor import NumberingCleanDialog

pytestmark = pytest.mark.gui


@pytest.mark.parametrize("theme,rtl,scale,preset", [("light", False, 100, "default"), ("dark", True, 150, "large")])
def test_numeric_editors_have_one_frame_readable_text_and_native_steps(qapp, qtbot, theme, rtl, scale, preset):
    manager = ThemeManager.get_instance()
    manager.initialize(qapp, theme)
    manager.apply_font_size_preset(preset)
    manager.apply_ui_scale(scale)
    pages: list[QWidget] = [factory(SettingsViewModel()) for factory in (ExportTab, LoggingTab, LinkTab, GeneralTab)]
    pages.append(NumberingCleanDialog())
    try:
        for direction in ("Up", "ChevronDown"):
            color = "white" if theme == "dark" else "black"
            assert QFile.exists(f":/qfluentwidgets/images/icons/{direction}_{color}.svg")
        for page in pages:
            qtbot.addWidget(page)
            page.setLayoutDirection(Qt.LayoutDirection.RightToLeft if rtl else Qt.LayoutDirection.LeftToRight)
            page.resize(850, 900)
            page.show()
            qapp.processEvents()
            for spin in page.findChildren(QWidget):
                if not isinstance(spin, (QSpinBox, QDoubleSpinBox)):
                    continue
                editor = spin.findChild(QLineEdit)
                assert editor is not None
                assert editor.minimumHeight() == 0
                frame = QStyleOptionFrame()
                frame.initFrom(editor)
                content = editor.style().subElementRect(QStyle.SubElement.SE_LineEditContents, frame, editor)
                assert content.height() >= editor.fontMetrics().height(), (
                    page,
                    spin,
                    content,
                    editor.fontMetrics().height(),
                )
                option = QStyleOptionSpinBox()
                spin.initStyleOption(option)
                up = spin.style().subControlRect(
                    QStyle.ComplexControl.CC_SpinBox, option, QStyle.SubControl.SC_SpinBoxUp, spin
                )
                down = spin.style().subControlRect(
                    QStyle.ComplexControl.CC_SpinBox, option, QStyle.SubControl.SC_SpinBoxDown, spin
                )
                assert up.height() > 0 and down.height() > 0 and not up.intersects(down)
                assert up.right() < editor.geometry().left() if rtl else up.left() > editor.geometry().right(), (
                    type(page).__name__,
                    spin.objectName(),
                    spin.rect(),
                    editor.geometry(),
                    up,
                    down,
                )
                spin.setEnabled(True)
                if isinstance(spin, QSpinBox):
                    spin.setValue(spin.minimum())
                else:
                    spin.setValue(spin.minimum())
                qtbot.keyClick(spin, Qt.Key.Key_Up)
                assert spin.value() > spin.minimum()
                before = spin.value()
                spin.setEnabled(False)
                qtbot.keyClick(spin, Qt.Key.Key_Up)
                assert spin.value() == before
    finally:
        for page in pages:
            page.close()
        manager.apply_ui_scale(100)
        manager.apply_font_size_preset("default")
        manager.apply_theme("light")

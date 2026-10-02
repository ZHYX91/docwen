"""Keep image size controls readable under the real application appearance."""

import pytest
from PySide6.QtCore import QPoint, QRect
from PySide6.QtWidgets import QApplication, QStyle, QStyleOptionFrame
from tests.support.gui_vm_fakes import FakeMainWindowViewModel

from docwen_gui.styles.theme_manager import ThemeManager
from docwen_gui.view_models.conversion_panel_vm import ConversionPanelViewModel
from docwen_gui.widgets.conversion_panel import ConversionPanel
from docwen_gui.widgets.panel_card import FormRow

pytestmark = pytest.mark.gui


@pytest.mark.parametrize("theme", ["light", "dark"])
@pytest.mark.parametrize("scale,font_preset", [(100, "default"), (150, "large")])
def test_size_limit_remains_readable_during_resize(qapp: QApplication, qtbot, theme, scale, font_preset):
    vm = ConversionPanelViewModel(FakeMainWindowViewModel())  # type: ignore[arg-type]
    panel = ConversionPanel(view_model=vm)
    qtbot.addWidget(panel)
    manager = ThemeManager.get_instance()
    manager.initialize(qapp, theme)
    manager.apply_font_size_preset(font_preset)
    manager.apply_ui_scale(scale)
    vm.set_file_info("image", "png", file_path="/test.png")
    number = panel._size_limit_edit
    unit = panel._size_unit_combo
    row = next(row for row in panel.findChildren(FormRow) if row.isAncestorOf(number))

    def readable_and_contained() -> None:
        option = QStyleOptionFrame()
        option.initFrom(number)
        content = number.style().subElementRect(QStyle.SubElement.SE_LineEditContents, option, number)
        margins = number.textMargins()
        content.adjust(margins.left(), margins.top(), -margins.right(), -margins.bottom())
        detail = (panel.width(), number.geometry(), content, number.fontMetrics().height())
        assert content.width() >= number.fontMetrics().horizontalAdvance(number.text()) + 4, detail
        assert content.height() >= number.fontMetrics().height(), detail
        rects = [QRect(control.mapTo(row, QPoint(0, 0)), control.size()) for control in (number, unit)]
        assert all(row.rect().contains(rect) for rect in rects), (row.rect(), rects)
        assert not rects[0].intersects(rects[1]), rects

    try:
        panel.show()
        for enabled in (False, True):
            vm.compress_mode = "limit_size" if enabled else "lossless"
            for width in (460, 310, 280, 460):
                panel.resize(width, 800)
                for _ in range(12):
                    qapp.processEvents()
                assert panel.width() == width
                readable_and_contained()
        assert vm.size_limit == 200
        assert number.text() == "200"
    finally:
        panel.close()
        manager.apply_ui_scale(100)
        manager.apply_font_size_preset("default")
        manager.apply_theme("light")

"""Input-action geometry during live resize, font changes, and UI scaling."""

import pytest
from PySide6.QtCore import QPoint, QRect
from PySide6.QtWidgets import QApplication

from docwen_gui.styles.theme_manager import ThemeManager
from docwen_gui.view_models.input_area_vm import InputAreaViewModel
from docwen_gui.view_models.main_window_vm import MainWindowViewModel
from docwen_gui.widgets.input_area import InputArea

pytestmark = pytest.mark.gui


@pytest.mark.parametrize("scale", [100, 150])
@pytest.mark.parametrize("font_preset", ["default", "large"])
def test_actions_remain_visible_and_separate_during_live_resize(
    qapp: QApplication, qtbot, scale: int, font_preset: str
) -> None:
    main_vm = MainWindowViewModel(controller=None)
    widget = InputArea(view_model=InputAreaViewModel(main_vm=main_vm))
    qtbot.addWidget(widget)
    manager = ThemeManager.get_instance()
    manager.initialize(qapp, "light")
    manager.apply_font_size_preset(font_preset)
    manager.apply_ui_scale(scale)
    controls = (widget.add_button, widget.paste_button, widget._paste_menu_button, widget.clear_button)

    def visible_and_separate() -> None:
        frame = widget._action_frame
        rects = [QRect(control.mapTo(frame, QPoint(0, 0)), control.size()) for control in controls]
        detail = (widget.width(), frame.rect(), rects, [control.sizeHint() for control in controls])
        assert all(frame.rect().contains(rect) for rect in rects), detail
        assert all(not left.intersects(right) for index, left in enumerate(rects) for right in rects[index + 1 :]), (
            detail
        )
        assert all(control.width() >= control.sizeHint().width() for control in controls), detail
        assert all(control.height() >= control.sizeHint().height() for control in controls), detail

    try:
        widget.show()
        for width in (720, 460, 360, 720):
            widget.resize(width, 760)
            qtbot.waitUntil(visible_and_separate, timeout=1000)
        for control, label in zip(
            (widget.add_button, widget.paste_button, widget.clear_button),
            ("Add a document", "Paste clipboard content", "Clear current input"),
            strict=True,
        ):
            control.setText(label)
        for width in (720, 460, 360, 720):
            widget.resize(width, 760)
            qtbot.waitUntil(visible_and_separate, timeout=1000)
    finally:
        widget.close()
        main_vm.cancel_inspection()
        qtbot.waitUntil(lambda: not main_vm.inspection_busy)
        manager.apply_ui_scale(100)
        manager.apply_font_size_preset("default")

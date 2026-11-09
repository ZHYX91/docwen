"""Input-action geometry during live resize, font changes, and UI scaling."""

from pathlib import Path

import pytest
from PySide6.QtCore import QPoint, QPointF, QRect, Qt
from PySide6.QtGui import QTextLayout, QTextOption
from PySide6.QtTest import QSignalSpy
from PySide6.QtWidgets import QApplication, QScrollArea, QVBoxLayout, QWidget

from docwen_core.models import FILE_INSPECTION_METADATA_KEY, FileRef
from docwen_gui.i18n import get_locale, set_locale
from docwen_gui.styles.theme_manager import ThemeManager
from docwen_gui.view_models.input_area_vm import InputAreaViewModel
from docwen_gui.view_models.main_window_vm import MainWindowViewModel
from docwen_gui.widgets.input_area import InputArea

pytestmark = pytest.mark.gui


@pytest.mark.parametrize(
    ("locale", "preset", "scale"),
    [("zh_CN", "small", 100), ("de_DE", "large", 150), ("ar_SA", "large", 150)],
)
def test_format_notice_fits_its_feedback_row_after_resizing(qapp, qtbot, locale, preset, scale) -> None:
    original_locale = get_locale()
    set_locale(locale)
    manager = ThemeManager.get_instance()
    manager.initialize(qapp, "light")
    manager.apply_font_size_preset(preset)
    manager.apply_ui_scale(scale)
    main_vm = MainWindowViewModel(controller=None)
    widget = InputArea(view_model=InputAreaViewModel(main_vm=main_vm))
    widget.setLayoutDirection(Qt.LayoutDirection.RightToLeft if locale == "ar_SA" else Qt.LayoutDirection.LeftToRight)
    selected = FileRef(
        path="/example/sample.txt",
        format="markdown",
        category="markdown",
        metadata={
            FILE_INSPECTION_METADATA_KEY: {
                "declared_format": "txt",
                "detected_format": "markdown",
                "warning_code": "FILE_FORMAT_COMPATIBLE_TEXT",
            }
        },
    )
    main_vm.set_selected_file(selected)
    widget.view_model.sync_selection([selected], current=True)
    # Match the main window's scrollable workflow column instead of forcing a
    # fixed height on a child whose translated content may need more room.
    scroll = QScrollArea()
    qtbot.addWidget(scroll)
    scroll.setWidgetResizable(True)
    content = QWidget()
    layout = QVBoxLayout(content)
    layout.setContentsMargins(0, 0, 0, 0)
    layout.addWidget(widget)
    layout.addStretch(1)
    scroll.setWidget(content)
    badge = widget._format_notice_label

    def fits() -> None:
        assert "Markdown" in badge.text() and "MARKDOWN" not in badge.text()
        assert widget._feedback_frame.rect().contains(badge.geometry())
        assert widget.rect().contains(QRect(badge.mapTo(widget, QPoint()), badge.size()))
        assert badge.height() == badge.heightForWidth(badge.width())
        detail_top = widget._selection_detail_label.mapTo(widget._feedback_frame, QPoint()).y()
        assert badge.geometry().bottom() < detail_top
        text_rect = badge.contentsRect()
        shaped = QTextLayout(badge.text(), badge.font(), badge)
        options = QTextOption()
        options.setWrapMode(QTextOption.WrapMode.WordWrap)
        shaped.setTextOption(options)
        shaped.beginLayout()
        height = 0.0
        while (line := shaped.createLine()).isValid():
            line.setLineWidth(text_rect.width() - badge.indent())
            line.setPosition(QPointF(0, height))
            height += line.height()
        shaped.endLayout()
        assert shaped.boundingRect().height() <= text_rect.height()

    try:
        scroll.resize(800, 400)
        scroll.show()
        qtbot.waitUntil(fits)
        wide_height = badge.height()
        for width in (360, 515, 800):
            scroll.resize(width, 400)
            qtbot.waitUntil(fits)
        assert badge.height() == wide_height
    finally:
        scroll.close()
        manager.apply_ui_scale(100)
        manager.apply_font_size_preset("default")
        set_locale(original_locale)


def test_clear_icon_resets_selection_without_deleting_the_file(qapp: QApplication, qtbot, tmp_path: Path) -> None:
    main_vm = MainWindowViewModel(controller=None)
    widget = InputArea(view_model=InputAreaViewModel(main_vm=main_vm))
    qtbot.addWidget(widget)
    source = tmp_path / "keep.md"
    source.write_text("Keep this document", encoding="utf-8")
    selected = FileRef(path=str(source), format="markdown", category="markdown")
    main_vm.set_selected_file(selected)
    widget.view_model.sync_selection([selected], current=True)
    cleared = QSignalSpy(widget.view_model.files_cleared)
    try:
        widget.show()
        widget.clear_button.setFocus()
        qtbot.keyClick(widget.clear_button, Qt.Key.Key_Space)
        assert cleared.count() == 1
        assert not widget.view_model.selection_message
        assert not widget._selection_label.text()
        assert source.read_text(encoding="utf-8") == "Keep this document"
    finally:
        widget.close()
        main_vm.cancel_inspection()
        qtbot.waitUntil(lambda: not main_vm.inspection_busy)


@pytest.mark.parametrize("scale", [90, 100, 110, 125, 150])
@pytest.mark.parametrize("font_preset", ["default", "large"])
@pytest.mark.parametrize("theme", ["light", "dark"])
def test_actions_remain_visible_and_separate_during_live_resize(
    qapp: QApplication, qtbot, scale: int, font_preset: str, theme: str
) -> None:
    main_vm = MainWindowViewModel(controller=None)
    widget = InputArea(view_model=InputAreaViewModel(main_vm=main_vm))
    qtbot.addWidget(widget)
    manager = ThemeManager.get_instance()
    manager.initialize(qapp, theme)
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
        assert all(control.width() >= control.sizeHint().width() for control in controls[:2]), detail
        assert all(control.height() >= control.sizeHint().height() for control in controls), detail
        assert rects[1].right() + 1 == rects[2].left(), detail
        assert rects[1].top() == rects[2].top(), detail
        assert rects[-1].right() == frame.width() - 1, detail
        assert all(control.width() == control.height() for control in controls[2:]), detail

    try:
        widget.show()
        for width in (720, 460, 360, 720):
            widget.resize(width, 760)
            qtbot.waitUntil(visible_and_separate, timeout=1000)
        compact_widths = [control.width() for control in controls]
        widget.resize(1000, 760)
        qtbot.waitUntil(visible_and_separate)
        assert [control.width() for control in controls] == compact_widths
        assert widget.paste_menu_button.arrowType() == Qt.ArrowType.NoArrow
        assert not widget.paste_menu_button.icon().isNull()
        for control, label in zip(
            (widget.add_button, widget.paste_button),
            ("Add a document", "Paste clipboard content"),
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
        manager.apply_theme("light")

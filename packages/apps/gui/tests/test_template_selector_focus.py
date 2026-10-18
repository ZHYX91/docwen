"""Template rows keep selection and keyboard focus without native double frames."""

from __future__ import annotations

import pytest
from PySide6.QtCore import QRect, Qt
from PySide6.QtGui import QImage, QPainter
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QListWidget, QListWidgetItem, QStyle, QStyleFactory, QStyleOptionViewItem

from docwen_gui.styles.template_selector import build_template_selector_stylesheet
from docwen_gui.widgets.template_item_delegate import SOURCE_ROLE, TemplateItemDelegate
from docwen_gui.widgets.template_selector import TemplateSelector

pytestmark = pytest.mark.gui


@pytest.mark.parametrize("direction", [Qt.LayoutDirection.LeftToRight, Qt.LayoutDirection.RightToLeft])
@pytest.mark.parametrize("draggable", [False, True])
def test_selected_row_does_not_gain_native_black_white_frame(qapp, direction, draggable):
    widget = QListWidget()
    widget.setObjectName("templateSelectorList")
    widget.setProperty("docwenTemplateList", True)
    widget.setLayoutDirection(direction)
    widget.setStyle(QStyleFactory.create("Windows"))
    widget.setStyleSheet(build_template_selector_stylesheet())
    item = QListWidgetItem("简体中文通用模板 / العربية", widget)
    item.setData(SOURCE_ROLE, "内置")
    delegate = TemplateItemDelegate(widget, draggable=draggable)
    option = QStyleOptionViewItem()
    option.initFrom(widget)
    option.rect = QRect(0, 0, 440, 44)
    option.state |= QStyle.StateFlag.State_Selected | QStyle.StateFlag.State_Active
    index = widget.model().index(0, 0)

    def render(focused):
        row = QStyleOptionViewItem(option)
        if focused:
            row.state |= QStyle.StateFlag.State_HasFocus
        else:
            row.state &= ~QStyle.StateFlag.State_HasFocus
        image = QImage(440, 44, QImage.Format.Format_ARGB32)
        image.fill(Qt.GlobalColor.transparent)
        painter = QPainter(image)
        delegate.paint(painter, row, index)
        painter.end()
        return image

    assert render(True) == render(False)
    widget.deleteLater()


def test_keyboard_navigation_keeps_current_template_and_list_focus(qapp):
    selector = TemplateSelector()
    selector.setStyleSheet(build_template_selector_stylesheet())
    selector.add_templates(["first", "second"])
    selector.resize(500, 240)
    selector.show()
    selector.activateWindow()
    selector._list.setCurrentRow(0)
    selector._list.setFocus(Qt.FocusReason.TabFocusReason)
    qapp.processEvents()
    QTest.keyClick(selector._list, Qt.Key.Key_Down)
    assert selector.get_selected() == "second"
    assert selector._list.hasFocus()
    QTest.keyClick(selector._list, Qt.Key.Key_Tab)
    assert selector.get_selected() == "second"
    selector.close()
    selector.deleteLater()


def test_management_list_keyboard_checkbox_and_order_remain_available(qapp, tmp_path):
    from docx import Document

    from docwen_gui.view_models.template_vm import TemplateViewModel
    from docwen_gui.widgets.settings.templates_tab import TemplatesTab
    from docwen_runtime.templates import TemplateManager, TemplateRegistry, TemplateStateStore

    builtin = tmp_path / "builtin"
    user = tmp_path / "user"
    builtin.mkdir()
    user.mkdir()
    for name in ("First", "Second"):
        Document().save(str(builtin / f"{name}.docx"))
    state = TemplateStateStore(tmp_path / "state.json", user_dir=user)
    registry = TemplateRegistry(builtin, extra_paths=[user], state_store=state, managed_user_dir=user)
    vm = TemplateViewModel(manager=TemplateManager(registry, state_store=state, user_dir=user))
    page = TemplatesTab(view_model=vm)
    page.setStyleSheet(build_template_selector_stylesheet())
    page.resize(700, 600)
    page.show()
    page.activateWindow()
    widget = page._lists["docx"]
    widget.setCurrentRow(0)
    widget.setFocus(Qt.FocusReason.TabFocusReason)
    qapp.processEvents()
    QTest.keyClick(widget, Qt.Key.Key_Down)
    assert widget.currentRow() == 1
    assert widget.hasFocus()
    resource_id = str(widget.currentItem().data(Qt.ItemDataRole.UserRole))
    QTest.keyClick(widget, Qt.Key.Key_Space)
    assert not vm.enabled[resource_id]
    first = widget.takeItem(0)
    widget.insertItem(1, first)
    page._persist_order("docx")
    assert [entry.id for entry in vm.templates if entry.target == "docx"] == [
        str(widget.item(i).data(Qt.ItemDataRole.UserRole)) for i in range(widget.count())
    ]
    page.close()
    page.deleteLater()

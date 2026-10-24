"""Filtering and session recents preserve catalog order and default ownership."""

from __future__ import annotations

import pytest
from PySide6.QtCore import Qt

from docwen_gui.widgets.template_selector import TemplateItemDetails, TemplateSelector

pytestmark = pytest.mark.gui


def test_search_preserves_selection_and_catalog_order(qtbot):
    selector = TemplateSelector()
    qtbot.addWidget(selector)
    selector.add_templates(
        ["first", "second", "third"],
        template_details={
            "first": TemplateItemDetails(display_name="Annual report"),
            "second": TemplateItemDetails(display_name="Note", usage_hint="Team minutes"),
            "third": TemplateItemDetails(display_name="报告"),
        },
    )
    selector.select_template("first", selection_source="auto_default")
    changes = []
    selector.template_selected.connect(changes.append)
    selector._search.setText("MINUTES")
    assert [selector._list.item(i).isHidden() for i in range(3)] == [True, False, True]
    assert selector.get_selected() == "first"
    assert changes == []
    selector._search.setText("报告")
    assert not selector._list.item(2).isHidden()
    selector._search.setText("no match")
    assert all(selector._list.item(i).isHidden() for i in range(3))
    assert selector.get_selected() == "first"
    selector._search.clear()
    assert [selector._list.item(i).data(Qt.ItemDataRole.UserRole) for i in range(3)] == ["first", "second", "third"]


def test_recents_come_from_started_tasks_are_bounded_and_pruned(qtbot):
    selector = TemplateSelector()
    qtbot.addWidget(selector)
    names = [f"template-{i}" for i in range(7)]
    selector.add_templates(names)
    selector.select_template(names[0], selection_source="auto_default")
    selector.select_template(names[1], selection_source="user")
    assert selector._recent == []
    for name in names:
        selector.mark_used(name)
    assert selector._recent == names[-1:-6:-1]
    selector.mark_used(names[-1])
    assert len(selector._recent) == 5
    selector._search.setText("missing")
    selector._select_recent(names[-1])
    assert selector.get_selected() == names[-1]
    assert selector._search.text() == ""
    feedback = selector.consume_selection_feedback()
    assert feedback is not None
    assert feedback.selection_source == "user"
    selector.add_templates(names[:2])
    assert selector._recent == []
    assert selector._recent_button.isHidden()

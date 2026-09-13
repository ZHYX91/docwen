"""Focused GUI coverage for template-management integration."""

from __future__ import annotations

import pytest

from docwen_gui.widgets.template_selector import TemplateSelector
from docwen_gui.widgets.template_selector_tabbed import TabbedTemplateSelector

pytestmark = pytest.mark.gui


def test_tabbed_selector_preserves_registry_order(qapp) -> None:
    selector = TabbedTemplateSelector()
    selector.load_templates(
        "docx",
        ["Third", "First", "Second"],
        preserve_order=True,
    )
    page = selector.get_selector("docx")
    assert page is not None
    assert [page._list.item(index).text() for index in range(page._list.count())] == [
        "Third",
        "First",
        "Second",
    ]
    selector.deleteLater()


def test_template_selector_exposes_management_actions(qapp) -> None:
    selector = TemplateSelector(template_type="docx")
    assert selector._empty_manage_button.text()
    selector.add_templates(["Default"])
    assert selector._manage_button.text()
    assert selector._manage_button.isVisible() or not selector.isVisible()
    selector.deleteLater()

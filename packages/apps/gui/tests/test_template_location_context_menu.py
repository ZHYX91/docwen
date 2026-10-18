"""Catalog-backed template locations use a non-modal menu harness."""

from __future__ import annotations

import pytest
from PySide6.QtCore import QPoint
from PySide6.QtGui import QAction
from PySide6.QtWidgets import QWidget

from docwen_gui.widgets import template_selector as selector_module
from docwen_gui.widgets.template_selector import TemplateItemDetails, TemplateSelector

pytestmark = pytest.mark.unit


def _spy_menu(monkeypatch: pytest.MonkeyPatch) -> list[QAction]:
    """Replace QMenu at the widget boundary; never enter native modal exec."""
    created: list[QAction] = []

    class FakeContextMenu:
        def __init__(self, parent: QWidget) -> None:
            self._parent = parent

        def addAction(self, label: str) -> QAction:
            action = QAction(label, self._parent)
            created.append(action)
            return action

        def exec(self, *_args: object) -> None:
            pass

    monkeypatch.setattr(selector_module, "QMenu", FakeContextMenu)
    return created


def test_context_menu_location_uses_catalog_path(qapp, monkeypatch: pytest.MonkeyPatch) -> None:
    actions = _spy_menu(monkeypatch)
    selector = TemplateSelector(template_type="docx")
    selector.add_templates(
        ["template.docx"],
        template_details={
            "template.docx": TemplateItemDetails(
                resource_id="template.docx",
                source_path="/catalog/source.docx",
            )
        },
    )
    selector.select_template("template.docx")
    seen: list[str] = []
    selector.location_requested.connect(seen.append)
    selector._show_list_context_menu(QPoint(-1, -1))
    assert len(actions) == 2
    actions[0].trigger()
    assert seen == ["/catalog/source.docx"]


def test_context_menu_does_not_offer_location_without_path(qapp, monkeypatch: pytest.MonkeyPatch) -> None:
    actions = _spy_menu(monkeypatch)
    selector = TemplateSelector(template_type="docx")
    selector.add_templates(["resource"], template_details={"resource": TemplateItemDetails()})
    selector.select_template("resource")
    selector._show_list_context_menu(QPoint(-1, -1))
    assert len(actions) == 1

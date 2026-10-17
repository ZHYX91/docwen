"""Catalog-backed template file location action avoids guessing filesystem paths."""

from PySide6.QtCore import QPoint
from PySide6.QtWidgets import QMenu

from docwen_gui.widgets.template_selector import TemplateItemDetails, TemplateSelector


def test_context_menu_location_uses_catalog_path(qapp, monkeypatch) -> None:
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
    actions = []
    seen: list[str] = []
    selector.location_requested.connect(seen.append)
    monkeypatch.setattr(QMenu, "exec", lambda menu, *_args: actions.extend(menu.actions()))
    selector._show_list_context_menu(QPoint(-1, -1))
    assert len(actions) == 2
    actions[0].trigger()
    assert seen == ["/catalog/source.docx"]


def test_context_menu_does_not_offer_location_without_path(qapp, monkeypatch) -> None:
    selector = TemplateSelector(template_type="docx")
    selector.add_templates(["resource"], template_details={"resource": TemplateItemDetails()})
    selector.select_template("resource")
    actions = []
    monkeypatch.setattr(QMenu, "exec", lambda menu, *_args: actions.extend(menu.actions()))
    selector._show_list_context_menu(QPoint(-1, -1))
    assert len(actions) == 1

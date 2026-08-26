"""Widget behaviour tests for TemplateSelector and TabbedTemplateSelector.

These tests validate:
- Construction and object names
- Empty state display
- Template list population and item rendering
- Selection tracking (user / auto_default / restore)
- Selection feedback and manual selection memory
- Location button state synchronisation
- Right-click context menu
- Item activation (double-click / Enter)
- Tab switching and cross-tab selection persistence
- Signal emission on user-facing paths
"""

from __future__ import annotations

from collections.abc import Iterator
from typing import Any

import pytest
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication, QListWidgetItem

from docwen_gui.i18n import t
from docwen_gui.widgets.template_selector import (
    TemplateItemDetails,
    TemplateSelectionFeedback,
    TemplateSelector,
)
from docwen_gui.widgets.template_selector_tabbed import TabbedTemplateSelector

pytestmark = pytest.mark.gui


def _sample_names() -> list[str]:
    """Three templates in unsorted order.

    TemplateSelector preserves caller order; TabbedTemplateSelector sorts.
    """
    return ["Standard Report", "Academic Paper", "Business Letter"]


_UNSORTED_FIRST = "Standard Report"

_UNSORTED_SECOND = "Academic Paper"

_UNSORTED_THIRD = "Business Letter"

_SORTED_FIRST = "Academic Paper"

_SORTED_SECOND = "Business Letter"

_SORTED_THIRD = "Standard Report"


def _sample_details() -> dict[str, TemplateItemDetails]:
    return {
        "Standard Report": TemplateItemDetails(
            usage_hint="Default document template",
            source_label="bundled",
            source_path="S:/DocWen/templates/Standard Report.docx",
            updated_label="2025-06-01 14:30",
        ),
        "Academic Paper": TemplateItemDetails(
            usage_hint="Academic formatting",
            source_label="bundled",
            updated_label="2025-05-15 09:00",
        ),
    }


def _assert_visible(widget: object) -> None:
    """Assert a widget is not hidden (works in offscreen mode)."""
    assert hasattr(widget, "isHidden")
    assert not widget.isHidden()  # type: ignore[union-attr]


def _assert_hidden(widget: object) -> None:
    """Assert a widget is hidden (works in offscreen mode)."""
    assert hasattr(widget, "isHidden")
    assert widget.isHidden()  # type: ignore[union-attr]


@pytest.fixture
def selector(qapp: QApplication) -> Iterator[TemplateSelector]:
    w = TemplateSelector(template_type="docx")
    w.show()
    yield w
    w.hide()
    w.deleteLater()


@pytest.fixture
def tabbed(qapp: QApplication) -> Iterator[TabbedTemplateSelector]:
    w = TabbedTemplateSelector()
    w.show()
    yield w
    w.hide()
    w.deleteLater()


__all__ = (
    "_SORTED_FIRST",
    "_SORTED_SECOND",
    "_SORTED_THIRD",
    "_UNSORTED_FIRST",
    "_UNSORTED_THIRD",
    "Any",
    "QApplication",
    "QListWidgetItem",
    "Qt",
    "TabbedTemplateSelector",
    "TemplateItemDetails",
    "TemplateSelectionFeedback",
    "TemplateSelector",
    "_assert_hidden",
    "_assert_visible",
    "_sample_details",
    "_sample_names",
    "pytest",
    "pytestmark",
    "selector",
    "t",
    "tabbed",
)

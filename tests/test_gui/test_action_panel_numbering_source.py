from __future__ import annotations

from types import SimpleNamespace

import pytest

pytest.importorskip("ttkbootstrap")

from docwen.gui.components.action_panel.base import ActionPanelBase

pytestmark = pytest.mark.unit


def test_action_panel_numbering_source_for_md_to_docx() -> None:
    panel = SimpleNamespace(file_type="docx")
    assert ActionPanelBase._get_numbering_option_source(panel) == "md"


def test_action_panel_numbering_source_for_doc_to_md() -> None:
    panel = SimpleNamespace(file_type="document")
    assert ActionPanelBase._get_numbering_option_source(panel) == "doc"


def test_action_panel_numbering_source_for_other_panels() -> None:
    panel = SimpleNamespace(file_type="xlsx")
    assert ActionPanelBase._get_numbering_option_source(panel) is None

"""Table structure controls preserve intent while OCR or models are unavailable."""

from __future__ import annotations

import pytest

from ._action_area_widget_support import vm as vm
from ._action_area_widget_support import widget as widget

pytestmark = pytest.mark.gui


@pytest.mark.parametrize("available", [True, False])
def test_table_option_tracks_ocr_and_retains_selection(widget, vm, monkeypatch, available):
    monkeypatch.setattr("docwen_core.text.table_recognition.table_recognition_available", lambda: available)
    vm.setup_for_document_file("/tables.docx")
    control = widget._table_recognition_cb
    assert control.isChecked()
    vm.set_file_to_md_option("extract_ocr", True)
    assert control.isEnabled() is available
    if available:
        control.click()
        assert vm.recognize_tables is False
    vm.set_file_to_md_option("extract_ocr", False)
    assert not control.isEnabled()
    assert control.isChecked() is (not available)
    vm.set_file_to_md_option("extract_ocr", True)
    assert control.isChecked() is (not available)
    assert vm.collect_options()["recognize_tables"] is (not available)
    if not available:
        assert control.toolTip()

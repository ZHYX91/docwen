"""Offscreen widget/model validation; this is not native-host acceptance."""

from __future__ import annotations

from typing import Any

import pytest
from tests.support.gui_vm_fakes import FakeMainWindowViewModel

from docwen_gui.i18n import t
from docwen_gui.view_models.conversion_panel_vm import ConversionPanelViewModel
from docwen_gui.widgets.conversion_panel import ConversionPanel

pytestmark = pytest.mark.gui


@pytest.mark.parametrize(
    ("value", "warning"),
    [
        ("1,999", "split_invalid_range_warning"),
        ("1-100001", "split_invalid_range_warning"),
        ("1,3x", "split_invalid_range_warning"),
        ("1,,3", "split_invalid_range_warning"),
        ("0,1", "split_invalid_range_warning"),
        ("1-3-4", "split_invalid_range_warning"),
        ("1-4", "split_all_pages_warning"),
        ("1-4,999", "split_invalid_range_warning"),
    ],
)
def test_invalid_complete_selection_disables_submission_and_recovers(qapp: Any, value: str, warning: str) -> None:
    vm = ConversionPanelViewModel(FakeMainWindowViewModel())  # type: ignore[arg-type]
    widget = ConversionPanel(view_model=vm)
    submitted = []
    vm.named_action_requested.connect(lambda *args: submitted.append(args))
    try:
        vm.set_file_info("layout", "pdf", file_path="/test.pdf")
        vm.set_pdf_info(4, "test.pdf")
        vm.split_mode = "custom"
        assert widget._page_input_edit is not None
        assert widget._split_pdf_button is not None
        assert widget._page_warning_label is not None
        widget._page_input_edit.setText(value)
        assert not widget._split_pdf_button.isEnabled()
        assert widget._page_warning_label.text() == t(f"conversion_panel.layout.{warning}")
        widget._on_split_pdf_clicked()
        assert submitted == []
        widget._page_input_edit.setText("3；1、3")
        assert widget._split_pdf_button.isEnabled()
        assert widget._page_warning_label.isHidden()
        widget._on_split_pdf_clicked()
        assert submitted == [("split_pdf", "/test.pdf", {"split_mode": "custom", "pages": [1, 3]})]
    finally:
        widget.deleteLater()


@pytest.mark.parametrize(("value", "mode"), [("*", "every_page"), ("#", "odd_even")])
def test_wildcards_and_metadata_refresh_keep_validation_local(qapp: Any, value: str, mode: str) -> None:
    vm = ConversionPanelViewModel(FakeMainWindowViewModel())  # type: ignore[arg-type]
    widget = ConversionPanel(view_model=vm)
    try:
        vm.set_file_info("layout", "pdf", file_path="/test.pdf")
        vm.set_pdf_info(4, "test.pdf")
        vm.split_mode = "custom"
        assert widget._page_input_edit is not None
        assert widget._split_pdf_button is not None
        assert widget._page_warning_label is not None
        widget._page_input_edit.setText(value)
        assert widget._split_pdf_button.isEnabled()
        assert vm.parse_split_input(value) == (mode, None)
        vm.set_pdf_info(1, "test.pdf")
        widget._update_page_validation(value)
        assert not widget._split_pdf_button.isEnabled()
        assert widget._page_warning_label.text() == t("conversion_panel.layout.split_mode_single_page_warning")
    finally:
        widget.deleteLater()

"""PDF split rejection explains the disabled action and recovers after editing."""

import pytest
from tests.support.gui_vm_fakes import FakeMainWindowViewModel

from docwen_gui.i18n import t
from docwen_gui.view_models.conversion_panel_vm import ConversionPanelViewModel
from docwen_gui.widgets.conversion_panel import ConversionPanel

pytestmark = pytest.mark.gui


@pytest.mark.parametrize(
    ("value", "key"),
    [
        ("1-2,3", "split_all_pages_warning"),
        ("abc", "split_invalid_range_warning"),
        ("8", "split_invalid_range_warning"),
    ],
)
def test_split_rejection_explains_disabled_button_and_recovers(qapp, value: str, key: str) -> None:
    vm = ConversionPanelViewModel(FakeMainWindowViewModel())  # type: ignore[arg-type]
    widget = ConversionPanel(view_model=vm)
    try:
        vm.set_file_info("layout", "pdf", file_path="/test.pdf")
        vm.set_pdf_info(3, "test.pdf")
        vm.split_mode = "custom"
        widget._page_input_edit.setText(value)
        assert not widget._split_pdf_button.isEnabled()
        assert not widget._page_warning_label.isHidden()
        assert widget._page_warning_label.text() == t(f"conversion_panel.layout.{key}")
        widget._page_input_edit.setText("1-2")
        assert widget._split_pdf_button.isEnabled()
        assert widget._page_warning_label.isHidden()
        assert not widget._page_warning_label.text()
    finally:
        widget.deleteLater()


def test_explicit_modes_hide_page_input_and_keep_custom_draft_local(qapp):
    vm = ConversionPanelViewModel(FakeMainWindowViewModel())  # type: ignore[arg-type]
    widget = ConversionPanel(view_model=vm)
    submitted = []
    vm.named_action_requested.connect(lambda *args: submitted.append(args))
    try:
        vm.set_file_info("layout", "pdf", file_path="/test.pdf")
        vm.set_pdf_info(4, "test.pdf")
        assert widget._page_range_row is not None
        assert widget._split_mode_combo is not None
        assert widget._page_range_row.isHidden()
        assert widget._page_range_help is not None and widget._page_range_help.isHidden()
        widget._split_pdf_button.click()
        assert submitted[-1][2] == {"split_mode": "every_page"}
        widget._split_mode_combo.setCurrentIndex(widget._split_mode_combo.findData("odd_even"))
        widget._split_pdf_button.click()
        assert submitted[-1][2] == {"split_mode": "odd_even"}
        widget._split_mode_combo.setCurrentIndex(widget._split_mode_combo.findData("custom"))
        assert not widget._page_range_row.isHidden()
        assert not widget._page_range_help.isHidden()
        assert not widget._split_pdf_button.isEnabled()
        widget._page_input_edit.setText("4,2")
        widget._split_pdf_button.click()
        assert submitted[-1][2] == {"split_mode": "custom", "pages": [2, 4]}
        vm.split_mode = "every_page"
        assert widget._page_range_row.isHidden()
        vm.split_mode = "custom"
        assert widget._page_input_edit.text() == "4,2"
        vm.set_file_info("layout", "pdf", file_path="/other.pdf")
        assert vm.page_input == "" and vm.split_mode == "every_page"
    finally:
        widget.deleteLater()


def test_unavailable_layout_defaults_never_invent_a_conversion_route(qapp):
    main = FakeMainWindowViewModel({"layout.document_format": "XLSX", "layout.render_format": "WEBP"})
    vm = ConversionPanelViewModel(main)  # type: ignore[arg-type]
    widget = ConversionPanel(view_model=vm)
    emitted = []
    vm.conversion_requested.connect(lambda *args: emitted.append(args))
    try:
        vm.set_file_info("layout", "pdf", file_path="/test.pdf")
        assert widget._layout_export_combo.findText("XLSX") == -1
        assert widget._layout_render_format_combo.findText("WEBP") == -1
        widget._layout_export_button.click()
        assert emitted[-1][0] == "docx"
        widget._layout_render_button.click()
        assert emitted[-1][0] == "png"
        assert main.controller.config_port.get("layout.document_format") == "XLSX"
        assert main.controller.config_port.get("layout.render_format") == "WEBP"
    finally:
        widget.deleteLater()

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

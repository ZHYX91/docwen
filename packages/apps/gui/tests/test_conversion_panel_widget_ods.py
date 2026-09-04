"""Focused GUI contracts for XLSX protection-aware ODS delivery."""

from __future__ import annotations

import zipfile
from collections.abc import Generator
from pathlib import Path

import pytest
from PySide6.QtWidgets import QApplication, QLabel, QLineEdit
from tests.support.gui_vm_fakes import FakeMainWindowViewModel

from docwen_gui.view_models.conversion_panel_vm import ConversionPanelViewModel
from docwen_gui.widgets.conversion_panel import ConversionPanel

pytestmark = pytest.mark.gui


@pytest.fixture
def vm() -> ConversionPanelViewModel:
    return ConversionPanelViewModel(FakeMainWindowViewModel())  # type: ignore[arg-type]


@pytest.fixture
def widget(qapp: QApplication, vm: ConversionPanelViewModel) -> Generator[ConversionPanel, None, None]:
    panel = ConversionPanel(view_model=vm)
    yield panel
    panel.deleteLater()


def _write_xlsx(path: Path, *, protected: bool) -> Path:
    protection = '<sheetProtection sheet="1" />' if protected else ""
    with zipfile.ZipFile(path, "w") as package:
        package.writestr(
            "xl/workbook.xml",
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" />',
        )
        package.writestr(
            "xl/worksheets/sheet1.xml",
            f'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">{protection}</worksheet>',
        )
    return path


def test_protected_xlsx_to_ods_requires_request_scoped_password_and_explicit_consent(
    widget: ConversionPanel,
    vm: ConversionPanelViewModel,
    tmp_path: Path,
) -> None:
    protected = _write_xlsx(tmp_path / "protected.xlsx", protected=True)
    vm.set_file_info("spreadsheet", "xlsx", file_path=str(protected))
    password_edit = widget._spreadsheet_password_edit
    consent = widget._spreadsheet_protection_loss_checkbox
    assert password_edit is not None
    assert password_edit.echoMode() == QLineEdit.EchoMode.Password
    assert consent is not None
    assert consent.isChecked() is False

    emitted: list[tuple[str, str, dict]] = []
    vm.conversion_requested.connect(lambda f, fp, o: emitted.append((f, fp, o)))
    assert widget.conversion_combo is not None
    widget.conversion_combo.setCurrentText("ODS")
    password_edit.setText("test")
    assert widget.conversion_button is not None
    assert widget.conversion_button.isEnabled() is False
    consent.setChecked(True)
    assert widget.conversion_button.isEnabled() is True
    widget.conversion_button.click()

    assert emitted == [
        (
            "ods",
            str(protected),
            {
                "spreadsheet_password": "test",
                "allow_spreadsheet_protection_loss": True,
            },
        )
    ]
    assert password_edit.text() == ""
    assert consent.isChecked() is False


def test_protected_batch_discloses_files_and_does_not_reuse_credentials(
    widget: ConversionPanel,
    vm: ConversionPanelViewModel,
    tmp_path: Path,
) -> None:
    first = _write_xlsx(tmp_path / "one.xlsx", protected=True)
    second = _write_xlsx(tmp_path / "two.xlsx", protected=True)
    vm.set_file_info(
        "spreadsheet",
        "xlsx",
        file_path=str(first),
        file_list=[str(first), str(second)],
        ui_mode="batch",
    )

    assert widget._spreadsheet_password_edit is None
    assert widget._spreadsheet_protection_loss_checkbox is None
    warning = next(label for label in widget.findChildren(QLabel) if "one.xlsx" in label.text())
    assert "two.xlsx" in warning.text()
    assert widget.conversion_combo is not None
    widget.conversion_combo.setCurrentText("ODS")
    assert widget.conversion_button is not None
    assert widget.conversion_button.isEnabled() is False


def test_unprotected_xlsx_does_not_request_credentials(
    widget: ConversionPanel,
    vm: ConversionPanelViewModel,
    tmp_path: Path,
) -> None:
    workbook = _write_xlsx(tmp_path / "plain.xlsx", protected=False)

    vm.set_file_info("spreadsheet", "xlsx", file_path=str(workbook))

    assert vm.spreadsheet_protected_files == ()
    assert widget._spreadsheet_password_edit is None
    assert widget._spreadsheet_protection_loss_checkbox is None


def test_ods_protection_options_do_not_leak_to_other_target(
    widget: ConversionPanel,
    vm: ConversionPanelViewModel,
    tmp_path: Path,
) -> None:
    protected = _write_xlsx(tmp_path / "protected.xlsx", protected=True)
    vm.set_file_info("spreadsheet", "xlsx", file_path=str(protected))
    assert widget._spreadsheet_password_edit is not None
    assert widget._spreadsheet_protection_loss_checkbox is not None
    widget._spreadsheet_password_edit.setText("test")
    widget._spreadsheet_protection_loss_checkbox.setChecked(True)
    assert widget.conversion_combo is not None
    widget.conversion_combo.setCurrentText("XLS")
    emitted: list[tuple[str, str, dict]] = []
    vm.conversion_requested.connect(lambda f, fp, o: emitted.append((f, fp, o)))

    assert widget.conversion_button is not None
    widget.conversion_button.click()

    assert emitted[-1][2] == {}

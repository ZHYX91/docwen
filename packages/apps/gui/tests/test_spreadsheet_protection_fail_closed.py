"""Fail-closed spreadsheet protection state contracts."""

import pytest

from docwen_gui.spreadsheet_protection import SpreadsheetProtectionInfo, inspect_xlsx_protection

pytestmark = pytest.mark.unit


def test_malformed_xlsx_has_an_unknown_protection_state(tmp_path) -> None:
    source = tmp_path / "malformed.xlsx"
    source.write_bytes(b"not-an-ooxml-package")

    info = inspect_xlsx_protection(str(source))

    assert info.is_unknown
    assert not info.is_protected


def test_encrypted_office_package_requires_a_password() -> None:
    info = SpreadsheetProtectionInfo(
        path="encrypted.xlsx",
        status="protected",
        protected_parts=("encrypted-package",),
    )

    assert info.is_protected
    assert info.requires_password


@pytest.mark.parametrize("kind", ["none", "workbook", "sheet"])
def test_empty_protection_elements_are_not_enabled_locks(tmp_path, kind):
    from openpyxl import Workbook

    source = tmp_path / "book.xlsx"
    workbook = Workbook()
    if kind == "workbook":
        workbook.security.lockStructure = True
    elif kind == "sheet":
        sheet = workbook.active
        assert sheet is not None
        sheet.protection.sheet = True
    workbook.save(source)
    info = inspect_xlsx_protection(str(source))
    assert info.status == ("none" if kind == "none" else "protected")

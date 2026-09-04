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

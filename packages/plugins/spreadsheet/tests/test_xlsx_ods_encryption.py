from __future__ import annotations

from pathlib import Path
from zipfile import ZipFile

import pytest
from msoffcrypto.format.ooxml import OOXMLFile
from openpyxl import Workbook

from docwen_plugin_spreadsheet.format_conversion.xlsx_ods_policy import (
    XlsxOdsPolicyError,
    inspect_xlsx_ods_policy,
    prepare_xlsx_for_ods,
)

pytestmark = pytest.mark.integration


def _write_encrypted_workbook(path: Path, password: str) -> None:
    plain = path.with_name("plain.xlsx")
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet["A1"] = "encrypted value"
    workbook.save(plain)
    with plain.open("rb") as source, path.open("wb") as target:
        OOXMLFile(source).encrypt(password, target)


def test_prepare_xlsx_for_ods_decrypts_only_the_private_copy(tmp_path: Path) -> None:
    source = tmp_path / "encrypted.xlsx"
    prepared = tmp_path / "prepared.xlsx"
    _write_encrypted_workbook(source, "secret")
    source_before = source.read_bytes()

    result = prepare_xlsx_for_ods(
        source,
        prepared,
        password="secret",
        allow_protection_loss=True,
    )

    assert source.read_bytes() == source_before
    assert result.protection_removed is True
    assert result.removed_protection_elements == ("encrypted-package",)
    assert inspect_xlsx_ods_policy(prepared).password_protected_elements == ()
    with ZipFile(prepared) as package:
        assert "xl/workbook.xml" in package.namelist()


@pytest.mark.parametrize(
    ("password", "allow_protection_loss", "expected_code"),
    [
        (None, False, "PROTECTION_PASSWORD_REQUIRED"),
        ("wrong", True, "PROTECTION_PASSWORD_INVALID"),
        ("secret", False, "PROTECTION_LOSS_CONSENT_REQUIRED"),
    ],
)
def test_prepare_xlsx_for_ods_rejects_encryption_admission_failures(
    tmp_path: Path,
    password: str | None,
    allow_protection_loss: bool,
    expected_code: str,
) -> None:
    source = tmp_path / "encrypted.xlsx"
    _write_encrypted_workbook(source, "secret")

    with pytest.raises(XlsxOdsPolicyError) as exc_info:
        prepare_xlsx_for_ods(
            source,
            tmp_path / "prepared.xlsx",
            password=password,
            allow_protection_loss=allow_protection_loss,
        )

    assert exc_info.value.diagnostic_code == expected_code

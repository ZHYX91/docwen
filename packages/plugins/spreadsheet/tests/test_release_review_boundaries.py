"""Typed, coordinate-owned ODS cached values rather than XML keyword scanning."""

from pathlib import Path
from xml.sax.saxutils import escape
from zipfile import ZipFile

import pytest

from docwen_plugin_spreadsheet.format_conversion.xlsx_ods_policy import (
    CachedExternalCell,
    XlsxOdsPolicyError,
    XlsxOdsPreparation,
    validate_prepared_ods,
)

pytestmark = pytest.mark.contract


def _write(path: Path, cell: str, extra: str = "") -> None:
    content = (
        '<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" '
        'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" '
        'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" '
        'xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0">'
        '<office:body><office:spreadsheet><table:table table:name="Sheet">'
        f"<table:table-row>{cell}</table:table-row>{extra}"
        "</table:table></office:spreadsheet></office:body></office:document-content>"
    )
    with ZipFile(path, "w") as package:
        package.writestr("content.xml", content)


@pytest.mark.parametrize("value", ["#REF!", "prefix #REF! suffix", "REF!"])
def test_release_review_ods_literal_ref_string(tmp_path: Path, value: str) -> None:
    path = tmp_path / "literal.ods"
    _write(
        path,
        f'<table:table-cell office:value-type="string"><text:p>{escape(value)}</text:p></table:table-cell>',
        "<text:p>unrelated #REF! documentation</text:p>",
    )
    preparation = XlsxOdsPreparation(
        output_path=str(path),
        external_links_flattened=True,
        protection_removed=False,
        flattened_cells=(CachedExternalCell("Sheet", "A1", "str", value),),
    )
    validate_prepared_ods(path, preparation)


@pytest.mark.parametrize(
    "mutation", ["wrong-value", "wrong-type", "error", "formula", "external-source", "external-formula"]
)
def test_release_review_ods_owned_negatives(tmp_path: Path, mutation: str) -> None:
    path = tmp_path / "invalid.ods"
    attributes = 'office:value-type="string"'
    value = "#REF!"
    extra = ""
    if mutation == "wrong-value":
        value = "wrong"
    elif mutation == "wrong-type":
        attributes = 'office:value-type="float" office:value="30"'
    elif mutation == "error":
        attributes += ' calcext:value-type="error"'
    elif mutation == "formula":
        attributes += ' table:formula="of:=#REF!"'
    elif mutation == "external-formula":
        attributes += " table:formula=\"of:='file:///tmp/external.xlsx'#$A$1\""
    else:
        extra = "<table:table-source/>"
    _write(path, f"<table:table-cell {attributes}><text:p>{value}</text:p></table:table-cell>", extra)
    preparation = XlsxOdsPreparation(
        output_path=str(path),
        external_links_flattened=True,
        protection_removed=False,
        flattened_cells=(CachedExternalCell("Sheet", "A1", "str", "#REF!"),),
    )
    with pytest.raises(XlsxOdsPolicyError) as caught:
        validate_prepared_ods(path, preparation)
    assert caught.value.diagnostic_code == "EXTERNAL_LINK_FLATTENING_NOT_DELIVERED"

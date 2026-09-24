"""Formula cache warnings distinguish serialized empty strings and missing values."""

import csv
from pathlib import Path

import pytest
from tests.support.formula_cache_fixture import write_formula_cache_fixture

from docwen_plugin_spreadsheet.csv_xlsx.converter import XlsxToCsvConverter, XlsxToTsvConverter

from ._csv_xlsx_support import _build_fake_context

pytestmark = pytest.mark.contract


@pytest.mark.parametrize(
    ("converter", "target", "separator"), [(XlsxToCsvConverter, "csv", ","), (XlsxToTsvConverter, "tsv", "\t")]
)
def test_formula_cache_states_and_bounded_multisheet_locations(tmp_path: Path, converter, target, separator) -> None:
    source = tmp_path / "cache.xlsx"
    write_formula_cache_fixture(source)
    original = source.read_bytes()
    staging = tmp_path / "stage"
    staging.mkdir()
    result = converter().convert(_build_fake_context(str(source), str(staging), target_format=target))
    assert result.success
    warnings = [diagnostic for diagnostic in result.diagnostics if diagnostic.level == "warning"]
    assert len(warnings) == 1
    message = warnings[0].message
    assert "27 formula cell(s)" in message
    assert "Calc!A1" not in message
    assert "Calc!B1" in message and "Calc!C1" in message
    assert "Calc!D1" not in message and "Calc!E1" not in message
    assert "Other!A18" in message and "Other!A19" not in message
    assert "7 more not listed" in message
    assert len(result.artifacts) == 2
    with Path(result.artifacts[0].staging_path).open(encoding="utf-8-sig", newline="") as stream:
        assert next(csv.reader(stream, delimiter=separator)) == ["", "", "", "0", "ordinary"]
    assert source.read_bytes() == original

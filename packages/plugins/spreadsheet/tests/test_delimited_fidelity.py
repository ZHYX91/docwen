"""Delimited inputs retain text semantics across spreadsheet routes."""

from pathlib import Path

import openpyxl
import pytest

from docwen_plugin_spreadsheet.csv_xlsx.converter import _build_delimited_workbook
from docwen_plugin_spreadsheet.delimited import decoded_samples
from docwen_plugin_spreadsheet.to_markdown.converter import _read_csv_flexible

pytestmark = [pytest.mark.golden, pytest.mark.contract]


@pytest.mark.parametrize("sep", [",", "\t"])
def test_delimited_fidelity_persisted_cells(tmp_path: Path, sep: str) -> None:
    values = ["1234567890123456789", "00123", "0.1234567890123456789", "=1+1", " 42 ", "#N/A"]
    source = tmp_path / "input.txt"
    source.write_text(sep.join(values), encoding="utf-8")
    workbook, rows = _build_delimited_workbook(str(source), sep=sep)
    output = tmp_path / "output.xlsx"
    workbook.save(output)
    workbook.close()
    loaded = openpyxl.load_workbook(output)
    try:
        assert rows == 1
        assert loaded.active is not None
        assert [cell.value for cell in loaded.active[1]] == values
        assert all(cell.data_type == "s" for cell in loaded.active[1])
    finally:
        loaded.close()
    frame = _read_csv_flexible(str(source), "tsv" if sep == "\t" else "csv")
    assert frame.iloc[0].tolist() == values


@pytest.mark.parametrize("encoding", ["utf-8", "utf-8-sig", "utf-16"])
def test_delimited_fidelity_sample_boundary(tmp_path: Path, encoding: str) -> None:
    source = tmp_path / "boundary.csv"
    # Short cells keep the workbook within Excel's cell length limit.
    content = "a,b\n" * 16383 + "a,中文\n"
    source.write_bytes(content.encode(encoding))
    assert decoded_samples(str(source))[0][0] == ("utf-16" if encoding == "utf-16" else "utf-8-sig")
    workbook, rows = _build_delimited_workbook(str(source), sep=",")
    try:
        assert rows == 16384
        assert workbook.active.cell(rows, 2).value == "中文"
    finally:
        workbook.close()
    assert _read_csv_flexible(str(source), "csv").iloc[-1, 1] == "中文"

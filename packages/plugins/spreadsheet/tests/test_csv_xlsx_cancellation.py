"""Cancellation and large-shape contracts for delimited/XLSX conversion."""

from __future__ import annotations

import csv
from pathlib import Path

import openpyxl
import pytest

from docwen_core.errors import CancellationRequested
from docwen_plugin_spreadsheet.csv_xlsx.converter import (
    CsvToXlsxConverter,
    TsvToXlsxConverter,
    XlsxToCsvConverter,
    XlsxToTsvConverter,
    _build_delimited_workbook,
)

from ._csv_xlsx_support import _build_fake_context

pytestmark = [pytest.mark.golden, pytest.mark.contract]


class _CancelAfterChecks:
    def __init__(self, fail_at: int) -> None:
        self.fail_at = fail_at
        self.checks = 0

    @property
    def is_cancelled(self) -> bool:
        return self.checks >= self.fail_at

    def check(self) -> None:
        self.checks += 1
        if self.checks >= self.fail_at:
            raise CancellationRequested("test cancellation")


@pytest.mark.parametrize(
    ("converter_type", "suffix", "sep"),
    [
        (CsvToXlsxConverter, ".csv", ","),
        (TsvToXlsxConverter, ".tsv", "\t"),
    ],
)
def test_delimited_to_xlsx_propagates_mid_parse_cancellation(
    tmp_path: Path,
    converter_type: type,
    suffix: str,
    sep: str,
) -> None:
    source = tmp_path / f"large{suffix}"
    with source.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.writer(handle, delimiter=sep)
        for index in range(2500):
            writer.writerow([index, f"value-{index}"])

    staging = tmp_path / "staging"
    staging.mkdir()
    context = _build_fake_context(str(source), str(staging), target_format="xlsx")
    cancellation = _CancelAfterChecks(fail_at=3)
    context._cancellation = cancellation

    with pytest.raises(CancellationRequested, match="test cancellation"):
        converter_type().convert(context)

    assert cancellation.checks == 3
    assert context.workspace.registered_artifacts == []
    assert not list(staging.glob("*.xlsx"))


@pytest.mark.parametrize(
    ("converter_type", "target_format", "suffix"),
    [
        (XlsxToCsvConverter, "csv", ".csv"),
        (XlsxToTsvConverter, "tsv", ".tsv"),
    ],
)
def test_xlsx_to_delimited_propagates_mid_write_cancellation_without_registering_partial_artifact(
    tmp_path: Path,
    converter_type: type,
    target_format: str,
    suffix: str,
) -> None:
    source = tmp_path / "large.xlsx"
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    assert sheet is not None
    for index in range(2500):
        sheet.append([index, f"value-{index}"])
    workbook.save(source)
    workbook.close()

    staging = tmp_path / "staging"
    staging.mkdir()
    context = _build_fake_context(str(source), str(staging), target_format=target_format)

    class _CancelAfterOutputStarts(_CancelAfterChecks):
        def check(self) -> None:
            if any(path.stat().st_size > 0 for path in staging.glob(f"*{suffix}")):
                super().check()

    cancellation = _CancelAfterOutputStarts(fail_at=1)
    context._cancellation = cancellation

    with pytest.raises(CancellationRequested, match="test cancellation"):
        converter_type().convert(context)

    assert cancellation.checks == 1
    assert context.workspace.registered_artifacts == []
    partials = list(staging.glob(f"*{suffix}"))
    assert len(partials) == 1
    assert partials[0].stat().st_size > 0


@pytest.mark.parametrize("sep", [",", "\t"])
def test_single_wide_row_can_cancel_before_completing(tmp_path: Path, sep: str) -> None:
    source = tmp_path / "wide.txt"
    source.write_text(sep.join(["value"] * 16000), encoding="utf-8")
    cancellation = _CancelAfterChecks(2)
    with pytest.raises(CancellationRequested):
        _build_delimited_workbook(str(source), sep=sep, cancel_check=cancellation.check)
    assert cancellation.checks == 2


@pytest.mark.parametrize("converter", [CsvToXlsxConverter, TsvToXlsxConverter])
def test_cancel_after_xlsx_save_closes_workbook_and_does_not_register(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, converter
) -> None:
    source = tmp_path / "input.csv"
    source.write_text("literal", encoding="utf-8")
    staging = tmp_path / "staging"
    staging.mkdir()
    context = _build_fake_context(str(source), str(staging), target_format="xlsx")
    cancellation = _CancelAfterChecks(1)
    context._cancellation = cancellation
    saved = False
    closed = False
    original_save = openpyxl.Workbook.save
    original_close = openpyxl.Workbook.close

    def save(workbook, path):
        nonlocal saved
        original_save(workbook, path)
        saved = True

    def close(workbook):
        nonlocal closed
        closed = True
        original_close(workbook)

    def check():
        if saved:
            raise CancellationRequested("after save")

    monkeypatch.setattr(cancellation, "check", check)
    monkeypatch.setattr(openpyxl.Workbook, "save", save)
    monkeypatch.setattr(openpyxl.Workbook, "close", close)
    with pytest.raises(CancellationRequested, match="after save"):
        converter().convert(context)
    assert saved and closed
    assert context.workspace.registered_artifacts == []
    partials = list(staging.glob("*.xlsx"))
    assert len(partials) == 1
    assert partials[0].stat().st_size > 0


@pytest.mark.parametrize("sep", [",", "\t"])
def test_delimited_wide_row_preserves_all_columns_without_timing_assumptions(
    tmp_path: Path,
    sep: str,
) -> None:
    source = tmp_path / "wide.txt"
    values = [f"c{index:04d}" for index in range(1024)]
    with source.open("w", encoding="utf-8", newline="") as handle:
        csv.writer(handle, delimiter=sep).writerow(values)

    workbook, row_count = _build_delimited_workbook(str(source), sep=sep)
    try:
        assert row_count == 1
        sheet = workbook.active
        assert sheet is not None
        assert sheet.max_column == len(values)
        assert sheet.cell(1, 1).value == values[0]
        assert sheet.cell(1, len(values)).value == values[-1]
    finally:
        workbook.close()


def test_formula_cache_scan_cancellation_closes_both_workbook_views(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_plugin_spreadsheet.csv_xlsx import converter as converter_module

    source = tmp_path / "input.xlsx"
    source.write_bytes(b"placeholder")
    staging = tmp_path / "staging"
    staging.mkdir()
    context = _build_fake_context(str(source), str(staging), target_format="csv")

    class _WorkbookView:
        def __init__(self) -> None:
            self.sheetnames = ["Sheet1"]
            self.closed = False

        def close(self) -> None:
            self.closed = True

    values = _WorkbookView()
    formulas = _WorkbookView()
    monkeypatch.setattr(converter_module, "_load_xlsx_views", lambda _path: (values, formulas))

    def cancel_scan(*_args: object, **_kwargs: object) -> tuple[int, list[str]]:
        raise CancellationRequested("scan cancellation")

    monkeypatch.setattr(converter_module, "_find_unavailable_formula_caches", cancel_scan)

    with pytest.raises(CancellationRequested, match="scan cancellation"):
        XlsxToCsvConverter().convert(context)

    assert values.closed is True
    assert formulas.closed is True
    assert context.workspace.registered_artifacts == []

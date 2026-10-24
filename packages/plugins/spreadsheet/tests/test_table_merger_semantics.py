"""Content invariants for real XLSX merges, independent of historical output."""

from __future__ import annotations

import re
from collections.abc import Sequence
from datetime import datetime
from pathlib import Path
from typing import Any
from zipfile import ZipFile

import openpyxl
import pytest
from openpyxl.utils.cell import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet
from tests.support.config import FakeConfigView
from tests.support.execution import FakeExecutionContext
from tests.support.logging import FakePluginLogger
from tests.support.progress import FakeProgressSink
from tests.support.workspace import FakeWorkspaceHandle

from docwen_core.cancellation import CancellationToken
from docwen_core.models.file_ref import FileRef
from docwen_core.models.request import ConversionRequest
from docwen_plugin_spreadsheet.table_merger.converter import TableMergerConverter
from docwen_plugin_spreadsheet.table_merger.semantics import value_identity

pytestmark = pytest.mark.contract


def _write(path: Path, rows: Sequence[Sequence[Any]], mode: str = "row") -> None:
    book = openpyxl.Workbook()
    sheet = book.active
    assert sheet is not None
    for row in zip(*rows, strict=True) if mode == "col" else rows:
        sheet.append(tuple(row))
    book.save(path)
    book.close()


def _run(tmp_path: Path, mode: str = "row", offset_range: int = 0):
    paths = [tmp_path / "base.xlsx", tmp_path / "collect.xlsx"]
    staging = tmp_path / "staging"
    staging.mkdir(exist_ok=True)
    context = FakeExecutionContext(
        request=ConversionRequest(
            request_id="merge-semantics",
            input_refs=[FileRef(path=str(p), format="xlsx", category="spreadsheet") for p in paths],
            target_format="xlsx",
            action_name="merge_tables",
            options={"merge_mode": mode, "offset_range": offset_range},
        ),
        workspace=FakeWorkspaceHandle(str(paths[0]), str(staging)),
        config=FakeConfigView(),
        progress=FakeProgressSink(),
        cancellation=CancellationToken(),
        logger=FakePluginLogger(),
    )
    return TableMergerConverter().convert(context)


def _records(path: str, mode: str) -> Sequence[tuple[object, ...]]:
    book = openpyxl.load_workbook(path, data_only=True)
    try:
        sheet = book.active
        assert sheet is not None
        rows = list(sheet.values)
        return list(zip(*rows, strict=True)) if mode == "col" else rows
    finally:
        book.close()


@pytest.mark.parametrize("mode", ["row", "col"])
@pytest.mark.parametrize(
    ("base", "collect", "expected", "decision"),
    [
        (["Alice", 7, None], ["Alice", 7, None], [("Alice", 7, None)], "skipped"),
        (["Alice", 7, None], ["Alice", None, None], [("Alice", 7, None)], "skipped"),
        (["Alice", None, None], ["Alice", 7, None], [("Alice", 7, None)], "replaced"),
        (["Alice", 7, None], ["Alice", None, 9], [("Alice", 7, None), ("Alice", None, 9)], "added"),
        (["Alice", 7, None], ["Alice", 8, None], [("Alice", 7, None), ("Alice", 8, None)], "added"),
    ],
)
def test_each_nonempty_field_survives_coverage_decisions(tmp_path, mode, base, collect, expected, decision):
    header = ["Name", "Old", "New"]
    _write(tmp_path / "base.xlsx", [header, base], mode)
    _write(tmp_path / "collect.xlsx", [header, collect], mode)
    originals = {p: p.read_bytes() for p in tmp_path.glob("*.xlsx")}
    result = _run(tmp_path, mode)
    assert result.success
    assert _records(result.artifacts[0].staging_path, mode) == [tuple(header), *expected]
    summary = result.metrics.extra["merge_summaries"][0]
    assert summary["counts"][decision] == (2 if decision == "skipped" else 1)
    assert summary["locations"][decision][-1] == ("row 2" if mode == "row" else "column 2")
    assert all(path.read_bytes() == before for path, before in originals.items())


@pytest.mark.parametrize("mode", ["row", "col"])
@pytest.mark.parametrize(
    ("first", "second"),
    [
        ("110101199001010011", "110101199001010012"),
        ("001", "1"),
        ("1e3", "1000"),
        ("1", 1),
        (True, 1),
        (" A", "A"),
        ("alpha", "beta"),
    ],
)
def test_distinct_identities_are_never_implicitly_deduplicated(tmp_path, mode, first, second):
    _write(tmp_path / "base.xlsx", [["ID", "Value"], [first, "A"]], mode)
    _write(tmp_path / "collect.xlsx", [["ID", "Value"], [second, "A"]], mode)
    result = _run(tmp_path, mode)
    assert result.success
    records = _records(result.artifacts[0].staging_path, mode)
    assert len(records) == 3
    assert {value_identity(row[0]) for row in records[1:]} == {value_identity(first), value_identity(second)}


def test_numeric_identity_does_not_round_integers_before_comparison():
    assert value_identity(2**53 + 1) != value_identity(float(2**53))
    assert value_identity(7) == value_identity(7.0)


@pytest.mark.parametrize("mode", ["row", "col", "cell"])
@pytest.mark.parametrize("formula_source", ["base", "collect"])
def test_missing_formula_cache_stops_before_creating_output(tmp_path, mode, formula_source):
    for name in ("base", "collect"):
        _write(tmp_path / f"{name}.xlsx", [["ID", "Value"], ["Alice", "=1+2" if name == formula_source else 4]])
    result = _run(tmp_path, mode)
    assert not result.success
    assert result.artifacts == []
    assert result.error.diagnostic_code == "MERGE-FORMULA-CACHE-MISSING"
    assert f"{formula_source}.xlsx" in result.error.message
    assert "Sheet!B2" in result.diagnostics[0].location
    assert "1 formula" in result.error.message
    assert not list((tmp_path / "staging").iterdir())


@pytest.mark.parametrize(
    ("cache_type", "cache", "expected"), [("n", "3", 7), ("str", "", 4), ("str", "=literal", "=literal,4")]
)
def test_cached_number_and_cached_empty_string_are_distinguished_from_unknown(tmp_path, cache_type, cache, expected):
    base = tmp_path / "base.xlsx"
    formula = f'"{cache}"' if cache_type == "str" else "1+2"
    _write(base, [["ID", "Value"], ["Alice", f"={formula}"]])
    with ZipFile(base) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    xml = parts["xl/worksheets/sheet1.xml"].decode()
    xml, count = re.subn(
        r'<c r="B2"[^>]*><f>.*?</f><v></v></c>', f'<c r="B2" t="{cache_type}"><f>{formula}</f><v>{cache}</v></c>', xml
    )
    assert count == 1
    parts["xl/worksheets/sheet1.xml"] = xml.encode()
    with ZipFile(base, "w") as archive:
        for name, content in parts.items():
            archive.writestr(name, content)
    _write(tmp_path / "collect.xlsx", [["ID", "Value"], ["Alice", 4]])
    result = _run(tmp_path, "cell")
    assert result.success
    assert _records(result.artifacts[0].staging_path, "row")[1][1] == expected


@pytest.mark.parametrize("mode", ["row", "col"])
@pytest.mark.parametrize("leading_source", ["base", "collect"])
def test_alignment_preserves_leading_fields_and_complementary_values(tmp_path, mode, leading_source):
    base = [["ID", "Old", "New"], ["Alice", 7, None]]
    collect = [["ID", "Old", "New"], ["Alice", None, 9]]
    leading = base if leading_source == "base" else collect
    leading[0].insert(0, "Extra")
    leading[1].insert(0, "keep")
    _write(tmp_path / "base.xlsx", base, mode)
    _write(tmp_path / "collect.xlsx", collect, mode)
    result = _run(tmp_path, mode, offset_range=2)
    assert result.success
    values = [value for row in _records(result.artifacts[0].staging_path, mode) for value in row]
    assert all(value in values for value in (7, 9, "keep", "Extra"))


@pytest.mark.parametrize(("source", "success"), [("base", False), ("collect", True)])
def test_cache_policy_follows_the_sheets_retained_in_the_output(tmp_path, source, success):
    for name in ("base", "collect"):
        path = tmp_path / f"{name}.xlsx"
        _write(path, [["ID", "Value"], ["Alice", 3]])
        if name == source:
            book = openpyxl.load_workbook(path)
            extra = book.create_sheet("Extra")
            extra["A1"] = "=1+2"
            book.save(path)
            book.close()
    result = _run(tmp_path, "cell")
    assert result.success is success
    if success:
        book = openpyxl.load_workbook(result.artifacts[0].staging_path)
        try:
            assert book.sheetnames == ["Sheet"]
        finally:
            book.close()
    else:
        assert "Extra!A1" in result.error.message


def _set_typed_value(path: Path, coordinate: str, value: Any, data_type: str) -> None:
    book = openpyxl.load_workbook(path)
    sheet = book.active
    assert isinstance(sheet, Worksheet)
    sheet[coordinate] = value
    sheet[coordinate].data_type = data_type
    book.save(path)
    book.close()


@pytest.mark.parametrize("mode", ["row", "col"])
@pytest.mark.parametrize(("base_type", "collect_type"), [("s", "e"), ("e", "s"), ("s", "s"), ("e", "e")])
def test_coverage_distinguishes_saved_text_and_error_records(tmp_path, mode, base_type, collect_type):
    paths = [tmp_path / "base.xlsx", tmp_path / "collect.xlsx"]
    for path, data_type in zip(paths, (base_type, collect_type), strict=True):
        _write(path, [["Name", "Value"], ["Alice", "#N/A"]], mode)
        _set_typed_value(path, "B2", "#N/A", data_type)
        saved = openpyxl.load_workbook(path, data_only=True)
        assert isinstance(saved.active, Worksheet)
        assert (saved.active["B2"].value, saved.active["B2"].data_type) == ("#N/A", data_type)
        saved.close()
    originals = {path: path.read_bytes() for path in paths}

    result = _run(tmp_path, mode)
    assert result.success
    expected_types = [base_type] if base_type == collect_type else [base_type, collect_type]
    for data_only in (False, True):
        saved = openpyxl.load_workbook(result.artifacts[0].staging_path, data_only=data_only)
        try:
            sheet = saved.active
            assert isinstance(sheet, Worksheet)
            records = list(sheet.iter_rows()) if mode == "row" else list(sheet.iter_cols())
            assert len(records) == 1 + len(expected_types)
            assert [[(cell.value, cell.data_type) for cell in record] for record in records[1:]] == [
                [("Alice", "s"), ("#N/A", data_type)] for data_type in expected_types
            ]
        finally:
            saved.close()
    summary = result.metrics.extra["merge_summaries"][0]
    assert summary["counts"] == {
        "added": 0 if base_type == collect_type else 1,
        "replaced": 0,
        "skipped": 2 if base_type == collect_type else 1,
        "combined": 0,
    }
    assert all(path.read_bytes() == before for path, before in originals.items())


@pytest.mark.parametrize("mode", ["row", "col"])
@pytest.mark.parametrize("collect_type", ["s", "e"])
@pytest.mark.parametrize("distance", [1, 20])
def test_alignment_and_similarity_use_saved_error_types(tmp_path, mode, collect_type, distance):
    base = tmp_path / "base.xlsx"
    collect = tmp_path / "collect.xlsx"
    _write(base, [["#N/A"]], mode)
    _write(collect, [["#N/A"]], mode)
    _set_typed_value(base, "A1", "#N/A", "e" if collect_type == "s" else "s")
    coordinate = f"A{distance + 1}" if mode == "row" else f"{get_column_letter(distance + 1)}1"
    _set_typed_value(base, coordinate, "#N/A", collect_type)
    _set_typed_value(collect, "A1", "#N/A", collect_type)
    originals = {path: path.read_bytes() for path in (base, collect)}
    base_book = openpyxl.load_workbook(base, data_only=True)
    collect_book = openpyxl.load_workbook(collect, data_only=True)
    try:
        base_ws, collect_ws = base_book.active, collect_book.active
        assert isinstance(base_ws, Worksheet) and isinstance(collect_ws, Worksheet)
        converter = TableMergerConverter()
        offset = (distance, 0) if mode == "row" else (0, distance)
        assert converter._find_best_offset(base_ws, collect_ws, 1) == offset
        assert converter._suggest_offset_by_value_matching(base_ws, collect_ws) == offset
        cells = converter._iter_cells_for_alignment(collect_ws)
        assert converter._calc_overlap(base_ws, cells, 0, 0) == 0
        assert converter._calc_overlap(base_ws, cells, *offset) == 1
        similarity = converter._row_similarity if mode == "row" else converter._col_similarity
        assert similarity(base_ws, 1, collect_ws, 1, 0) == 0.0
        assert similarity(base_ws, distance + 1, collect_ws, 1, 0) == 1.0
    finally:
        base_book.close()
        collect_book.close()
    assert all(path.read_bytes() == before for path, before in originals.items())


@pytest.mark.parametrize("mode", ["row", "col"])
@pytest.mark.parametrize("collect_type", ["s", "e"])
def test_inserted_record_is_placed_next_to_the_same_error_type(tmp_path, mode, collect_type):
    base = tmp_path / "base.xlsx"
    collect = tmp_path / "collect.xlsx"
    _write(base, [["ID", "Value"], ["old", "#N/A"], ["old", "#N/A"]], mode)
    _write(collect, [["ID", "Value"], ["incoming", "#N/A"]], mode)
    other_type = "e" if collect_type == "s" else "s"
    _set_typed_value(base, "B2", "#N/A", other_type)
    _set_typed_value(base, "B3" if mode == "row" else "C2", "#N/A", collect_type)
    _set_typed_value(collect, "B2", "#N/A", collect_type)
    originals = {path: path.read_bytes() for path in (base, collect)}
    result = _run(tmp_path, mode)
    assert result.success
    saved = openpyxl.load_workbook(result.artifacts[0].staging_path, data_only=True)
    try:
        sheet = saved.active
        assert isinstance(sheet, Worksheet)
        records = list(sheet.iter_rows()) if mode == "row" else list(sheet.iter_cols())
        assert [[(cell.value, cell.data_type) for cell in record] for record in records[1:]] == [
            [("old", "s"), ("#N/A", other_type)],
            [("old", "s"), ("#N/A", collect_type)],
            [("incoming", "s"), ("#N/A", collect_type)],
        ]
    finally:
        saved.close()
    assert all(path.read_bytes() == before for path, before in originals.items())


@pytest.mark.parametrize("base_type", ["s", "e"])
def test_cell_mode_keeps_equal_text_policy_across_error_types(tmp_path, base_type):
    for name, data_type in (("base", base_type), ("collect", "e" if base_type == "s" else "s")):
        path = tmp_path / f"{name}.xlsx"
        _write(path, [["ID", "Value"], ["Alice", "#N/A"]])
        _set_typed_value(path, "B2", "#N/A", data_type)
    originals = {path: path.read_bytes() for path in tmp_path.glob("*.xlsx")}
    result = _run(tmp_path, "cell")
    assert result.success
    saved = openpyxl.load_workbook(result.artifacts[0].staging_path, data_only=True)
    try:
        sheet = saved.active
        assert isinstance(sheet, Worksheet)
        assert (sheet["B2"].value, sheet["B2"].data_type) == ("#N/A", base_type)
    finally:
        saved.close()
    assert all(path.read_bytes() == before for path, before in originals.items())


@pytest.mark.parametrize("mode", ["row", "col"])
@pytest.mark.parametrize("insert", [False, True])
@pytest.mark.parametrize(
    ("value", "data_type"),
    [
        ("=1+2", "s"),
        ("#N/A", "s"),
        ("#N/A", "e"),
        (True, "b"),
        (datetime(2026, 10, 24), "d"),
    ],
)
def test_row_and_column_writes_preserve_saved_scalar_types(tmp_path, mode, insert, value, data_type):
    _write(tmp_path / "base.xlsx", [["ID", "Value"], ["same", "other" if insert else None]], mode)
    _write(tmp_path / "collect.xlsx", [["ID", "Value"], ["same", value]], mode)
    _set_typed_value(tmp_path / "collect.xlsx", "B2", value, data_type)
    originals = {path: path.read_bytes() for path in tmp_path.glob("*.xlsx")}
    result = _run(tmp_path, mode)
    assert result.success
    for data_only in (False, True):
        book = openpyxl.load_workbook(result.artifacts[0].staging_path, data_only=data_only)
        try:
            sheet = book.active
            assert isinstance(sheet, Worksheet)
            cells = [cell for row in sheet for cell in row if cell.value == value]
            assert len(cells) == 1
            assert cells[0].data_type == data_type
        finally:
            book.close()
    assert all(path.read_bytes() == before for path, before in originals.items())


@pytest.mark.parametrize(
    ("base", "collect", "data_type", "expected"),
    [
        (None, "=literal", "s", "=literal"),
        ("=literal", "=literal", "s", "=literal"),
        ("=literal", "suffix", "s", "=literal,suffix"),
        (None, "#N/A", "s", "#N/A"),
        (None, "#N/A", "e", "#N/A"),
        ("#N/A", "#N/A", "e", "#N/A"),
    ],
)
def test_cell_merge_preserves_the_result_origin_type(tmp_path, base, collect, data_type, expected):
    for name, value in [("base", base), ("collect", collect)]:
        path = tmp_path / f"{name}.xlsx"
        _write(path, [["ID", "Value"], ["same", value]])
        if value is not None:
            _set_typed_value(path, "B2", value, data_type)
    result = _run(tmp_path, "cell")
    assert result.success
    for data_only in (False, True):
        book = openpyxl.load_workbook(result.artifacts[0].staging_path, data_only=data_only)
        sheet = book.active
        assert isinstance(sheet, Worksheet)
        assert sheet["B2"].value == expected
        assert sheet["B2"].data_type == data_type
        book.close()


@pytest.mark.parametrize(("value", "data_type"), [("=literal", "s"), ("#N/A", "s"), ("#N/A", "e")])
def test_unmerge_preserves_literal_and_error_types_after_save(tmp_path, value, data_type):
    book = openpyxl.Workbook()
    sheet = book.active
    assert isinstance(sheet, Worksheet)
    sheet["A1"] = value
    sheet["A1"].data_type = data_type
    sheet.merge_cells("A1:B2")
    TableMergerConverter._unmerge_all_cells(sheet)
    output = tmp_path / "expanded.xlsx"
    book.save(output)
    book.close()
    for data_only in (False, True):
        saved = openpyxl.load_workbook(output, data_only=data_only)
        saved_sheet = saved.active
        assert isinstance(saved_sheet, Worksheet)
        assert [(cell.value, cell.data_type) for row in saved_sheet for cell in row] == [(value, data_type)] * 4
        saved.close()


def test_base_with_inactive_chartsheet_keeps_the_chart_part(tmp_path):
    from openpyxl.chart import BarChart, Reference

    for name in ("base", "collect"):
        _write(tmp_path / f"{name}.xlsx", [["ID", "Value"], ["same", 3]])
    base = tmp_path / "base.xlsx"
    book = openpyxl.load_workbook(base)
    chart = BarChart()
    sheet = book.active
    assert isinstance(sheet, Worksheet)
    chart.add_data(Reference(sheet, min_col=2, min_row=1, max_row=2), titles_from_data=True)
    book.create_chartsheet("Chart").add_chart(chart)
    book.save(base)
    book.close()
    original = base.read_bytes()
    result = _run(tmp_path, "cell")
    assert result.success
    saved = openpyxl.load_workbook(result.artifacts[0].staging_path)
    assert saved.sheetnames == ["Sheet", "Chart"]
    assert len(saved.chartsheets) == 1
    saved.close()
    with ZipFile(result.artifacts[0].staging_path) as package:
        assert "xl/charts/chart1.xml" in package.namelist()
    assert base.read_bytes() == original

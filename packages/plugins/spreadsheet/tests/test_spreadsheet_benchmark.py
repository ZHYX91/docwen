"""Opt-in, reproducible source performance observations without timing gates."""

import csv
import json
import os
import platform
import threading
import time
import tracemalloc
from pathlib import Path

import openpyxl
import pytest

from docwen_core.errors import CancellationRequested
from docwen_plugin_spreadsheet.csv_xlsx.converter import (
    XlsxToCsvConverter,
    XlsxToTsvConverter,
    _build_delimited_workbook,
)

from ._csv_xlsx_support import _build_fake_context

pytestmark = [
    pytest.mark.integration,
    pytest.mark.slow,
    pytest.mark.skipif(
        os.environ.get("DOCWEN_RUN_SPREADSHEET_BENCHMARK") != "1", reason="Opt-in spreadsheet performance baseline"
    ),
]


@pytest.mark.parametrize(("rows", "columns", "cell_size"), [(50000, 8, 16), (2, 16000, 16), (64, 4, 32767)])
@pytest.mark.parametrize("sep", [",", "\t"])
def test_spreadsheet_performance_observation(
    tmp_path: Path, capsys, rows: int, columns: int, cell_size: int, sep: str
) -> None:
    source = tmp_path / "input.txt"
    value = "x" * cell_size
    with source.open("w", encoding="utf-8", newline="") as stream:
        writer = csv.writer(stream, delimiter=sep)
        for _ in range(rows):
            writer.writerow([value] * columns)
    output = tmp_path / "output.xlsx"
    tracemalloc.start()
    started = time.perf_counter()
    workbook, count = _build_delimited_workbook(str(source), sep=sep)
    parsed = time.perf_counter()
    workbook.save(output)
    workbook.close()
    saved = time.perf_counter()
    staging = tmp_path / "export"
    staging.mkdir()
    converter = XlsxToCsvConverter() if sep == "," else XlsxToTsvConverter()
    result = converter.convert(
        _build_fake_context(str(output), str(staging), target_format="csv" if sep == "," else "tsv")
    )
    exported = time.perf_counter()
    assert result.success
    with Path(result.artifacts[0].staging_path).open(encoding="utf-8-sig", newline="") as stream:
        roundtrip_rows = 0
        for row in csv.reader(stream, delimiter=sep):
            assert len(row) == columns and row[-1] == value
            roundtrip_rows += 1
        assert roundtrip_rows == rows
    _, peak = tracemalloc.get_traced_memory()
    tracemalloc.stop()
    loaded = openpyxl.load_workbook(output, read_only=True)
    try:
        assert count == rows and loaded.active is not None
        assert loaded.active.cell(rows, columns).value == value
    finally:
        loaded.close()
    requested = threading.Event()
    cancelled_at = []

    def request() -> None:
        cancelled_at.append(time.perf_counter())
        requested.set()

    timer = threading.Timer(0.01, request)
    timer.start()

    def check() -> None:
        if requested.is_set():
            raise CancellationRequested("benchmark cancellation")

    outcome = "completed-before-cancel"
    try:
        workbook, _ = _build_delimited_workbook(str(source), sep=sep, cancel_check=check)
        workbook.close()
    except CancellationRequested:
        outcome = "cancelled"
    ended = time.perf_counter()
    timer.cancel()
    timer.join()
    receipt = {
        "rows": rows,
        "columns": columns,
        "cellCodePoints": cell_size,
        "separator": repr(sep),
        "parseSeconds": parsed - started,
        "saveSeconds": saved - parsed,
        "exportSeconds": exported - saved,
        "peakPythonBytes": peak,
        "cancellation": outcome,
        "cancelResponseSeconds": ended - cancelled_at[0] if cancelled_at and outcome == "cancelled" else None,
        "host": platform.platform(),
        "python": platform.python_version(),
        "openpyxl": openpyxl.__version__,
    }
    with capsys.disabled():
        print("SPREADSHEET_BASELINE=" + json.dumps(receipt, sort_keys=True), flush=True)  # noqa: T201 -- benchmark artifact

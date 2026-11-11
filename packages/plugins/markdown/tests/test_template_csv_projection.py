"""Template-to-CSV uses the shared bounded ordinary-worksheet projection."""

from __future__ import annotations

import csv
from pathlib import Path
from zipfile import ZipFile

import pytest
from lxml import etree
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference

from docwen_core.cancellation import CancellationToken
from docwen_core.errors import CancellationRequested
from docwen_plugin_markdown.to_spreadsheet import converter, template_xlsx
from docwen_plugin_markdown.to_spreadsheet.converter import MdToCsvConverter

from .conftest import make_context

pytestmark = pytest.mark.contract

_S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def _context(tmp_path: Path, template: Path):
    source = tmp_path / "source.md"
    source.write_text("---\nname: Filled\n---\n", encoding="utf-8")
    return make_context(str(source), target_format="csv", options={"template_name": str(template)})


def _capture_views(monkeypatch: pytest.MonkeyPatch):
    loaded, closed = [], []
    load, close = template_xlsx.load_template_workbook, Workbook.close

    def capture(path):
        workbook = load(path)
        loaded.append(workbook)
        return workbook

    def close_view(workbook):
        closed.append(workbook)
        close(workbook)

    monkeypatch.setattr(template_xlsx, "load_template_workbook", capture)
    monkeypatch.setattr(Workbook, "close", close_view)
    return loaded, closed


@pytest.mark.parametrize("chart_index", [0, 1, 2])
def test_chartsheet_position_does_not_change_ordinary_csv_identity(tmp_path: Path, chart_index: int) -> None:
    template = tmp_path / "template.xlsx"
    workbook = Workbook()
    first = workbook.active
    first.title = "First"
    first.append(["{{name}}", 1])
    workbook.create_sheet("Second")["A1"] = "=1+2"
    chart = BarChart()
    chart.add_data(Reference(first, min_col=2, min_row=1, max_row=1))
    workbook.create_chartsheet("Plot", chart_index).add_chart(chart)
    workbook.save(template)
    before = template.read_bytes()
    context, workspace = _context(tmp_path, template)

    result = MdToCsvConverter().convert(context)

    assert result.success, result.error
    assert [item.metadata["sheet_name"] for item in result.artifacts] == ["First", "Second"]
    assert [item.metadata["sheet_index"] for item in result.artifacts] == [0, 1]
    assert [item.metadata["sheet_count"] for item in result.artifacts] == [2, 2]
    assert [item.is_primary for item in result.artifacts] == [True, False]
    assert [item.suggested_name for item in result.artifacts] == ["source_First.csv", "source_Second.csv"]
    actual = []
    for artifact in result.artifacts:
        with Path(artifact.staging_path).open(encoding="utf-8-sig", newline="") as handle:
            actual.append(list(csv.reader(handle)))
    assert actual == [[["Filled", "1"]], [["=1+2"]]]
    assert workspace.registered_artifacts == result.artifacts
    assert template.read_bytes() == before


def test_chart_only_template_refuses_empty_success(tmp_path: Path) -> None:
    template = tmp_path / "template.xlsx"
    workbook = Workbook()
    source = workbook.active
    source["A1"] = 1
    chart = BarChart()
    chart.add_data(Reference(source, min_col=1, min_row=1, max_row=1))
    workbook.create_chartsheet("Plot").add_chart(chart)
    workbook.remove(source)
    workbook.save(template)
    context, workspace = _context(tmp_path, template)

    result = MdToCsvConverter().convert(context)

    assert not result.success and result.error.diagnostic_code == "SPREADSHEET-NO-WORKSHEETS"
    assert result.artifacts == workspace.registered_artifacts == []
    assert list(Path(workspace.staging_dir).iterdir()) == []


def test_saved_sparse_template_does_not_gain_blank_cells(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    template = tmp_path / "template.xlsx"
    workbook = Workbook()
    workbook.active.cell(101, 100, "last")
    workbook.save(template)
    before = template.read_bytes()
    loaded, closed = _capture_views(monkeypatch)
    context, _ = _context(tmp_path, template)

    result = MdToCsvConverter().convert(context)

    assert result.success, result.error
    with Path(result.artifacts[0].staging_path).open(encoding="utf-8-sig", newline="") as handle:
        rows = list(csv.reader(handle))
    assert len(rows) == 101 and all(len(row) == 100 for row in rows)
    assert rows[-1][-1] == "last"
    assert sum(value != "" for row in rows for value in row) == 1
    assert len(loaded) == 1 and loaded == closed
    assert set(loaded[0].worksheets[0]._cells) == {(101, 100)}
    assert template.read_bytes() == before


@pytest.mark.parametrize(("shapes", "success"), [([(2, 3)], True), ([(3, 3)], False), ([(2, 2), (2, 2)], False)])
def test_template_output_budget_is_checked_for_all_sheets_before_writing(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, shapes: list[tuple[int, int]], success: bool
) -> None:
    template = tmp_path / "template.xlsx"
    workbook = Workbook()
    for index, (row, column) in enumerate(shapes):
        sheet = workbook.active if index == 0 else workbook.create_sheet()
        sheet.cell(row, column, "last")
    workbook.save(template)
    before = template.read_bytes()
    loaded, closed = _capture_views(monkeypatch)
    monkeypatch.setattr(converter, "MAX_SPREADSHEET_PROJECTION_CELLS", 6)
    context, workspace = _context(tmp_path, template)

    result = MdToCsvConverter().convert(context)

    assert result.success is success
    if success:
        with Path(result.artifacts[0].staging_path).open(encoding="utf-8-sig", newline="") as handle:
            assert list(csv.reader(handle)) == [["", "", ""], ["", "", "last"]]
    else:
        assert result.error.diagnostic_code == "SPREADSHEET-PROJECTION-LIMIT"
        assert result.artifacts == workspace.registered_artifacts == []
        assert list(Path(workspace.staging_dir).iterdir()) == []
        assert workspace._next_id == 0
    assert loaded == closed
    assert template.read_bytes() == before


@pytest.mark.parametrize("kind", ["merge", "stored"])
def test_template_allocations_are_rejected_before_the_workbook_library(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, kind: str
) -> None:
    template = tmp_path / "template.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    for index in range(1, 8 if kind == "stored" else 2):
        sheet.cell(index, 1, "value")
    workbook.save(template)
    if kind == "merge":
        with ZipFile(template) as package:
            parts = {name: package.read(name) for name in package.namelist()}
        root = etree.fromstring(parts["xl/worksheets/sheet1.xml"])
        merges = etree.SubElement(root, f"{{{_S}}}mergeCells", count="1")
        etree.SubElement(merges, f"{{{_S}}}mergeCell", ref="A1:D2")
        parts["xl/worksheets/sheet1.xml"] = etree.tostring(root)
        with ZipFile(template, "w") as package:
            for name, payload in parts.items():
                package.writestr(name, payload)
    before = template.read_bytes()

    def forbidden_load(_path):
        pytest.fail("The expansive template must be rejected before native loading")

    monkeypatch.setattr(template_xlsx, "load_template_workbook", forbidden_load)
    monkeypatch.setattr(converter, "MAX_SPREADSHEET_PROJECTION_CELLS", 6)
    context, workspace = _context(tmp_path, template)

    result = MdToCsvConverter().convert(context)

    assert not result.success and result.error.diagnostic_code == "SPREADSHEET-PROJECTION-LIMIT"
    assert result.artifacts == workspace.registered_artifacts == []
    assert list(Path(workspace.staging_dir).iterdir()) == []
    assert template.read_bytes() == before


@pytest.mark.parametrize(("rows", "columns"), [(3001, 1), (1, 3001)])
def test_template_csv_cancellation_cleans_the_entire_private_batch_and_closes_view(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, rows: int, columns: int
) -> None:
    template = tmp_path / "template.xlsx"
    workbook = Workbook()
    workbook.active["A1"] = "completed first sheet"
    workbook.create_sheet("Cancel").cell(rows, columns, "last")
    workbook.save(template)
    before = template.read_bytes()
    context, workspace = _context(tmp_path, template)
    token = CancellationToken()
    context._cancellation = token
    loaded, closed = _capture_views(monkeypatch)
    load = template_xlsx.load_template_workbook

    class Values(dict):
        visits = 0

        def get(self, key, default=None):
            self.visits += 1
            if self.visits == 17:
                token.cancel("inside CSV rectangle")
            return super().get(key, default)

    def capture_values(path):
        view = load(path)
        view["Cancel"]._cells = Values(view["Cancel"]._cells)
        return view

    monkeypatch.setattr(template_xlsx, "load_template_workbook", capture_values)

    with pytest.raises(CancellationRequested):
        MdToCsvConverter().convert(context)

    assert len(loaded) == 1 and loaded == closed
    assert 17 <= loaded[0]["Cancel"]._cells.visits <= 1000
    assert len(loaded[0]["Cancel"]._cells) == 1
    assert workspace.registered_artifacts == []
    assert list(Path(workspace.staging_dir).iterdir()) == []
    assert template.read_bytes() == before


def test_existing_staging_path_is_neither_overwritten_nor_owned_for_cleanup(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    template = tmp_path / "template.xlsx"
    workbook = Workbook()
    workbook.active["A1"] = "replacement"
    workbook.save(template)
    context, workspace = _context(tmp_path, template)
    existing = Path(workspace.staging_dir) / "existing.csv"
    existing.write_bytes(b"previous owner")
    monkeypatch.setattr(workspace, "create_artifact_path", lambda _kind, _suffix: str(existing))

    result = MdToCsvConverter().convert(context)

    assert not result.success
    assert existing.read_bytes() == b"previous owner"
    assert workspace.registered_artifacts == []
    assert list(Path(workspace.staging_dir).iterdir()) == [existing]

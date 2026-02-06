from __future__ import annotations

from pathlib import Path

import openpyxl
import pytest

import docwen.table_merger.core as table_core
from docwen.table_merger.core import TableMerger


pytestmark = pytest.mark.unit


def _write_xlsx(path: Path, value: str = "x") -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = value
    wb.save(path)
    wb.close()


def test_preprocess_table_xlsx_copies_into_temp_dir(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    src = tmp_path / "a.xlsx"
    _write_xlsx(src)

    temp_dir = tmp_path / "tmp"
    temp_dir.mkdir()

    merger = TableMerger()
    merger.temp_dir = str(temp_dir)

    monkeypatch.setattr(table_core, "detect_actual_file_format", lambda *_args, **_kwargs: "xlsx", raising=True)
    monkeypatch.setattr("docwen.utils.workspace_manager.prepare_input_file", lambda p, *_args, **_kwargs: p, raising=True)

    out = merger._preprocess_table(str(src), is_base=True)
    out_path = Path(out)
    assert out_path.exists() is True
    assert out_path.name.endswith("_base.xlsx")
    assert out in merger.temp_files


def test_preprocess_table_csv_converts_and_removes_intermediate(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    src = tmp_path / "a.csv"
    src.write_text("a,b\n1,2\n", encoding="utf-8")

    temp_dir = tmp_path / "tmp"
    temp_dir.mkdir()

    intermediate = tmp_path / "intermediate.xlsx"

    def _stub_csv_to_xlsx(_csv_path: str) -> str:
        _write_xlsx(intermediate, value="mid")
        return str(intermediate)

    merger = TableMerger()
    merger.temp_dir = str(temp_dir)

    monkeypatch.setattr(table_core, "detect_actual_file_format", lambda *_args, **_kwargs: "csv", raising=True)
    monkeypatch.setattr("docwen.utils.workspace_manager.prepare_input_file", lambda p, *_args, **_kwargs: p, raising=True)
    monkeypatch.setattr(table_core, "csv_to_xlsx", _stub_csv_to_xlsx, raising=True)

    out = merger._preprocess_table(str(src), is_base=False)
    out_path = Path(out)
    assert out_path.exists() is True
    assert out_path.name.endswith("_collect.xlsx")
    assert intermediate.exists() is False


def test_preprocess_table_unsupported_format_raises(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    src = tmp_path / "a.bin"
    src.write_bytes(b"x")

    temp_dir = tmp_path / "tmp"
    temp_dir.mkdir()

    merger = TableMerger()
    merger.temp_dir = str(temp_dir)

    monkeypatch.setattr(table_core, "detect_actual_file_format", lambda *_args, **_kwargs: "pdf", raising=True)
    monkeypatch.setattr("docwen.utils.workspace_manager.prepare_input_file", lambda p, *_args, **_kwargs: p, raising=True)

    with pytest.raises(ValueError) as excinfo:
        merger._preprocess_table(str(src), is_base=False)

    assert "不支持的表格格式" in str(excinfo.value)


def test_merge_tables_smoke_flow(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    base = tmp_path / "base.xlsx"
    collect = tmp_path / "collect.xlsx"
    _write_xlsx(base, value="base")
    _write_xlsx(collect, value="collect")

    merger = TableMerger()

    monkeypatch.setattr(table_core, "t", lambda *_args, **_kwargs: "merged", raising=True)
    monkeypatch.setattr("docwen.utils.workspace_manager.get_output_directory", lambda *_args, **_kwargs: str(tmp_path), raising=True)

    monkeypatch.setattr(TableMerger, "_preprocess_table", lambda self, file_path, **_kwargs: file_path, raising=True)
    monkeypatch.setattr(TableMerger, "_unmerge_all_cells", lambda *_args, **_kwargs: None, raising=True)
    monkeypatch.setattr(TableMerger, "_find_best_offset", lambda *_args, **_kwargs: (0, 0), raising=True)
    monkeypatch.setattr(TableMerger, "_merge_by_row", lambda *_args, **_kwargs: None, raising=True)

    ok, msg, out_path = merger.merge_tables(str(base), [str(collect)], mode=TableMerger.MODE_BY_ROW)

    assert ok is True
    assert "成功汇总" in msg
    assert out_path is not None
    assert Path(out_path).exists() is True
    assert Path(out_path).suffix == ".xlsx"
    assert merger.temp_dir is not None
    assert Path(merger.temp_dir).exists() is False


def test_merge_cell_values_rules() -> None:
    merger = TableMerger()

    assert merger._merge_cell_values(None, None) is None
    assert merger._merge_cell_values("", "x") == "x"
    assert merger._merge_cell_values("x", "") == "x"
    assert merger._merge_cell_values("1", 2) == 3
    assert merger._merge_cell_values("A", "A") == "A"
    assert merger._merge_cell_values("A", "B") == "A,B"


def test_try_convert_to_number() -> None:
    merger = TableMerger()

    assert merger._try_convert_to_number(1) == 1
    assert merger._try_convert_to_number(1.5) == 1.5
    assert merger._try_convert_to_number(" 2 ") == 2
    assert merger._try_convert_to_number("2.0") == 2
    assert merger._try_convert_to_number("2.5") == 2.5
    assert merger._try_convert_to_number("x") is None
    assert merger._try_convert_to_number(None) is None


def test_calculate_overlap_and_find_best_offset() -> None:
    base_wb = openpyxl.Workbook()
    base_ws = base_wb.active
    base_ws["B2"] = "X"

    collect_wb = openpyxl.Workbook()
    collect_ws = collect_wb.active
    collect_ws["A1"] = "X"

    merger = TableMerger()

    assert merger._calculate_overlap(base_ws, collect_ws, row_offset=1, col_offset=1) == 1
    assert merger._calculate_overlap(base_ws, collect_ws, row_offset=0, col_offset=0) == 0
    assert merger._find_best_offset(base_ws, collect_ws) == (1, 1)

    base_wb.close()
    collect_wb.close()


def test_unmerge_all_cells_fills_values() -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "V"
    ws.merge_cells("A1:B2")

    merger = TableMerger()
    merger._unmerge_all_cells(ws)

    assert len(list(ws.merged_cells.ranges)) == 0
    assert ws["A1"].value == "V"
    assert ws["A2"].value == "V"
    assert ws["B1"].value == "V"
    assert ws["B2"].value == "V"

    wb.close()

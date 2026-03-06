"""services 单元测试。"""

from __future__ import annotations

from pathlib import Path

import openpyxl
import pytest

from docwen.table_merger.core import TableMerger

pytestmark = pytest.mark.unit


def _write_xlsx(path: Path, value: str) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = value
    wb.save(path)
    wb.close()


def test_table_merger_preprocess_table_avoids_same_stem_collision(tmp_path: Path) -> None:
    dir1 = tmp_path / "d1"
    dir2 = tmp_path / "d2"
    dir1.mkdir()
    dir2.mkdir()

    f1 = dir1 / "same.xlsx"
    f2 = dir2 / "same.xlsx"
    _write_xlsx(f1, "ONE")
    _write_xlsx(f2, "TWO")

    merger = TableMerger()
    merger.temp_dir = str(tmp_path / "temp")
    Path(merger.temp_dir).mkdir()

    p1 = merger._preprocess_table(str(f1), is_base=False)
    p2 = merger._preprocess_table(str(f2), is_base=False)

    assert p1 != p2

"""utils 单元测试。"""

from __future__ import annotations

from pathlib import Path

import pytest

from docwen.utils.workspace_manager import replace_file_atomic

pytestmark = pytest.mark.unit


def test_replace_file_atomic_overwrites_target_and_removes_temp(tmp_path: Path) -> None:
    target = tmp_path / "out.txt"
    target.write_text("old", encoding="utf-8")

    temp_file = tmp_path / "temp.txt"
    temp_file.write_text("new", encoding="utf-8")

    replaced = replace_file_atomic(str(temp_file), str(target))

    assert replaced == str(target)
    assert target.read_text(encoding="utf-8") == "new"
    assert temp_file.exists() is False


def test_replace_file_atomic_can_create_backup(tmp_path: Path) -> None:
    target = tmp_path / "out.txt"
    target.write_text("old", encoding="utf-8")

    temp_file = tmp_path / "temp.txt"
    temp_file.write_text("new", encoding="utf-8")

    replace_file_atomic(str(temp_file), str(target), create_backup=True)

    backups = sorted(tmp_path.glob("out.txt.bak*"))
    assert len(backups) == 1
    assert backups[0].read_text(encoding="utf-8") == "old"

"""utils 单元测试。"""

from __future__ import annotations

from pathlib import Path

import pytest

from docwen.utils.workspace_manager import (
    collect_intermediate_items_from_dir,
    save_intermediates_from_temp_dir,
)

pytestmark = pytest.mark.unit


@pytest.mark.unit
def test_collect_intermediate_items_excludes_input_and_respects_excludes(tmp_path: Path) -> None:
    work = tmp_path / "work"
    work.mkdir()

    (work / "input.md").write_text("x", encoding="utf-8")
    (work / "a.docx").write_bytes(b"x")
    (work / "skip.bin").write_bytes(b"x")
    (work / "img").mkdir()

    items = collect_intermediate_items_from_dir(
        str(work),
        exclude_filenames=["skip.bin"],
    )

    names = {Path(i.path).name for i in items}
    assert "input.md" not in names
    assert "skip.bin" not in names
    assert "a.docx" in names
    assert "img" in names

    kinds = {Path(i.path).name: i.kind for i in items}
    assert kinds["img"] == "intermediate_dir"
    assert kinds["a.docx"] == "intermediate_docx"


@pytest.mark.unit
def test_save_intermediates_from_temp_dir_moves_and_skips_input_prefix(tmp_path: Path) -> None:
    work = tmp_path / "work"
    out = tmp_path / "out"
    work.mkdir()
    out.mkdir()

    (work / "input.pdf").write_bytes(b"%PDF-1.4\n%fake\n")
    (work / "a.tmp").write_bytes(b"x")

    saved = save_intermediates_from_temp_dir(
        str(work),
        str(out),
        move=True,
        include_dirs=False,
    )

    assert (out / "a.tmp").exists()
    assert not (work / "a.tmp").exists()
    assert not any(Path(item.path).name == "input.pdf" for item, _ in saved)

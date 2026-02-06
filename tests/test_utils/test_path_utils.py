from __future__ import annotations

import os
import sys
from pathlib import Path

import pytest

from docwen.utils.path_utils import (
    ensure_dir_exists,
    generate_output_path,
    is_subpath,
    normalize_path,
    safe_join_path,
    split_path,
    validate_extension,
)


@pytest.mark.unit
def test_ensure_dir_exists(tmp_path: Path) -> None:
    target = tmp_path / "a" / "b"
    assert ensure_dir_exists(str(target)) is True
    assert target.exists()
    assert target.is_dir()


@pytest.mark.unit
def test_normalize_path_basic(tmp_path: Path) -> None:
    p = tmp_path / "XyZ" / "File.txt"
    normalized = normalize_path(str(p))
    assert os.path.isabs(normalized)
    if sys.platform == "win32":
        assert normalized == normalized.lower()


@pytest.mark.unit
def test_safe_join_path_allows_inside(tmp_path: Path) -> None:
    base = tmp_path / "base"
    base.mkdir()
    p = safe_join_path(str(base), "child", "file.txt")
    assert is_subpath(p, str(base)) is True


@pytest.mark.unit
def test_safe_join_path_blocks_traversal(tmp_path: Path) -> None:
    base = tmp_path / "base"
    base.mkdir()
    with pytest.raises(ValueError):
        safe_join_path(str(base), "..", "evil.txt")


@pytest.mark.unit
def test_validate_extension() -> None:
    assert validate_extension("a.docx", "docx") == "a.docx"
    assert validate_extension("a.txt", "docx") == "a.docx"
    assert validate_extension("a", "md") == "a.md"


@pytest.mark.unit
def test_generate_output_path_strips_existing_timestamp(tmp_path: Path) -> None:
    output_dir = tmp_path / "out"
    input_path = tmp_path / "report_20240101_010101_fromMd.docx"

    out = generate_output_path(
        str(input_path),
        output_dir=str(output_dir),
        section="s1",
        add_timestamp=False,
        description="fromMd",
        file_type="docx",
    )

    dir_path, base, ext = split_path(out)
    assert Path(dir_path) == output_dir
    assert base == "report_s1_fromMd"
    assert ext == "docx"


"""utils 单元测试。"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

from docwen.utils.path_utils import (
    collect_files_from_folder,
    ensure_dir_exists,
    ensure_unique_directory_path,
    generate_named_output_path,
    generate_output_path,
    is_subpath,
    make_output_stem,
    normalize_path,
    safe_join_path,
    split_path,
    strip_timestamp_suffix,
    validate_extension,
)


pytestmark = pytest.mark.unit


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
    assert Path(normalized).is_absolute()
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


@pytest.mark.unit
def test_generate_output_path_appends_counter_when_name_conflicts(tmp_path: Path) -> None:
    output_dir = tmp_path / "out"
    output_dir.mkdir()

    input_path = tmp_path / "report.docx"
    (output_dir / "report.docx").write_text("old", encoding="utf-8")

    out = generate_output_path(
        str(input_path),
        output_dir=str(output_dir),
        add_timestamp=False,
        file_type="docx",
    )

    assert Path(out).parent == output_dir
    assert Path(out).name == "report_001.docx"


@pytest.mark.unit
def test_strip_timestamp_suffix_removes_suffix_and_trailing_parts() -> None:
    assert strip_timestamp_suffix("report_20240101_010101_fromMd") == "report"
    assert strip_timestamp_suffix("report") == "report"


@pytest.mark.unit
def test_make_output_stem_respects_timestamp_override() -> None:
    stem = make_output_stem(
        "C:/in/report.docx",
        section="s1",
        add_timestamp=True,
        description="fromDocx",
        timestamp_override="20990101_000000",
    )
    assert stem == "report_s1_20990101_000000_fromDocx"


@pytest.mark.unit
def test_generate_named_output_path_uses_timestamp_and_uniquifies(tmp_path: Path) -> None:
    out_dir = tmp_path / "out"
    out_dir.mkdir()

    p1 = generate_named_output_path(
        output_dir=str(out_dir),
        base_name="合并PDF",
        file_type="pdf",
        add_timestamp=True,
        timestamp_override="20990101_000000",
    )
    Path(p1).write_text("x", encoding="utf-8")

    p2 = generate_named_output_path(
        output_dir=str(out_dir),
        base_name="合并PDF",
        file_type="pdf",
        add_timestamp=True,
        timestamp_override="20990101_000000",
    )
    assert Path(p1).name == "合并PDF_20990101_000000.pdf"
    assert Path(p2).name == "合并PDF_20990101_000000_001.pdf"


@pytest.mark.unit
def test_ensure_unique_directory_path_adds_counter(tmp_path: Path) -> None:
    base = tmp_path / "out"
    base.mkdir()

    target = base / "report_20240101_010101_fromPdf"
    target.mkdir()

    unique = ensure_unique_directory_path(str(target))
    assert unique.endswith("_001")
    assert not Path(unique).exists()


@pytest.mark.unit
def test_collect_files_from_folder_ignores_dirs_dedups_and_sorts(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    a = tmp_path / "b.bin"
    b = tmp_path / "a.bin"
    a.write_text("x", encoding="utf-8")
    b.write_text("x", encoding="utf-8")

    git_dir = tmp_path / ".git"
    git_dir.mkdir()
    (git_dir / "c.bin").write_text("x", encoding="utf-8")

    monkeypatch.setattr(
        "docwen.utils.path_utils.validate_file_format",
        lambda *_args, **_kwargs: {"actual_format": "pdf"},
        raising=True,
    )
    monkeypatch.setattr("docwen.utils.path_utils.ACTUAL_FORMAT_TO_CATEGORY", {"pdf": "layout"}, raising=False)

    files = collect_files_from_folder(str(tmp_path))
    assert [Path(p).name for p in files] == ["a.bin", "b.bin"]

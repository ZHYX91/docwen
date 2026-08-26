from __future__ import annotations

import tomllib
from pathlib import Path

import pytest

pytestmark = pytest.mark.contract

_PROJECT_ROOT = Path(__file__).resolve().parents[2]
_PYPROJECT = _PROJECT_ROOT / "pyproject.toml"
_MAX_TEST_FILE_LINES = 700
_BANNED_TEST_FILES = {
    "test_batch2_qt_widgets.py": "批处理 Qt Widgets 已按组件拆分，不应回流到单一上帝文件",
}


def _line_count(file_path: Path) -> int:
    with file_path.open(encoding="utf-8") as handle:
        return sum(1 for _ in handle)


def _configured_test_roots() -> list[Path]:
    with _PYPROJECT.open("rb") as handle:
        config = tomllib.load(handle)
    raw_paths = config["tool"]["pytest"]["ini_options"]["testpaths"]
    assert isinstance(raw_paths, list) and raw_paths
    roots = [_PROJECT_ROOT / str(path) for path in raw_paths]
    assert len(roots) == len(set(roots)), "pytest testpaths must not contain duplicates"
    assert all(root.is_dir() for root in roots), "pytest testpaths must all exist"
    return roots


def test_every_test_module_stays_within_the_size_limit() -> None:
    offenders: list[str] = []
    discovered: set[str] = set()

    for test_root in _configured_test_roots():
        for file_path in sorted(test_root.rglob("test_*.py")):
            relative_path = file_path.relative_to(_PROJECT_ROOT).as_posix()
            if relative_path in discovered:
                offenders.append(f"{relative_path}: collected from overlapping pytest testpaths")
                continue
            discovered.add(relative_path)
            if file_path.name in _BANNED_TEST_FILES:
                offenders.append(f"{relative_path}: {_BANNED_TEST_FILES[file_path.name]}")
                continue

            line_count = _line_count(file_path)
            if line_count > _MAX_TEST_FILE_LINES:
                offenders.append(f"{relative_path}: {line_count} lines > {_MAX_TEST_FILE_LINES}")

    assert not offenders, "Test-file size governance failed:\n" + "\n".join(offenders)

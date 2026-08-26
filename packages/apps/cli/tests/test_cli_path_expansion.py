from __future__ import annotations

from pathlib import Path

import pytest

from docwen_cli.utils import expand_paths

pytestmark = pytest.mark.integration


def test_expand_paths_accepts_powershell_drag_and_drop_quoting(tmp_path: Path) -> None:
    source = tmp_path / "dragged input.md"
    source.write_text("payload", encoding="utf-8")
    expected = [str(source.resolve())]

    for raw in (
        str(source),
        f"'{source}'",
        f'"{source}"',
        f"& '{source}'",
        f'   & "{source}"   ',
    ):
        assert expand_paths([raw]) == expected


def test_expand_paths_recurses_directories_without_extension_prefilter(tmp_path: Path) -> None:
    top = tmp_path / "top.md"
    nested = tmp_path / "nested" / "payload.unknown"
    deeper = tmp_path / "nested" / "deeper" / "report.docx"
    for path in (top, nested, deeper):
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(path.name, encoding="utf-8")

    assert expand_paths([str(tmp_path)]) == [
        str(deeper.resolve()),
        str(nested.resolve()),
        str(top.resolve()),
    ]


def test_expand_paths_preserves_first_seen_order_and_deduplicates(tmp_path: Path) -> None:
    first = tmp_path / "b.docx"
    second = tmp_path / "a.docx"
    nested = tmp_path / "nested" / "c.docx"
    for path in (first, second, nested):
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(path.name, encoding="utf-8")

    assert expand_paths([str(first), str(second), str(first), str(tmp_path)]) == [
        str(first.resolve()),
        str(second.resolve()),
        str(nested.resolve()),
    ]


def test_expand_paths_prunes_tool_directories_but_keeps_explicit_files(tmp_path: Path) -> None:
    kept = tmp_path / "keep.docx"
    hidden = tmp_path / ".GIT" / "hidden.docx"
    dependency = tmp_path / "node_modules" / "dependency.docx"
    for path in (kept, hidden, dependency):
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(path.name, encoding="utf-8")

    assert expand_paths([str(tmp_path)]) == [str(kept.resolve())]
    assert expand_paths([str(hidden)]) == [str(hidden.resolve())]

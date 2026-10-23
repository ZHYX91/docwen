"""Regression tests for literal paths and batch admission accounting."""

from __future__ import annotations

from argparse import Namespace
from pathlib import Path
from types import SimpleNamespace

import pytest

from docwen_cli.commands import convert
from docwen_cli.utils import expand_paths

pytestmark = pytest.mark.contract


def test_literal_bracketed_file_takes_precedence_over_glob(tmp_path: Path) -> None:
    literal = tmp_path / "Report [1].md"
    other = tmp_path / "Report 1.md"
    literal.write_text("intended", encoding="utf-8")
    other.write_text("unrelated", encoding="utf-8")

    assert expand_paths([str(literal)]) == [str(literal.resolve())]
    assert expand_paths([f'"{literal}"']) == [str(literal.resolve())]

    folder = tmp_path / "Archive [2026]"
    folder.mkdir()
    nested = folder / "note.md"
    nested.write_text("content", encoding="utf-8")
    assert expand_paths([str(folder)]) == [str(nested.resolve())]


@pytest.mark.parametrize(("json_mode", "quiet"), [(False, False), (False, True), (True, False)])
@pytest.mark.parametrize("explicit_batch", [False, True])
def test_invalid_batch_member_reaches_result_accounting_in_every_mode(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    json_mode: bool,
    quiet: bool,
    explicit_batch: bool,
) -> None:
    good = str(tmp_path / "good.md")
    missing = str(tmp_path / "missing.md")
    Path(good).write_text("# valid", encoding="utf-8")
    args = Namespace(
        command="batch" if explicit_batch else "convert",
        batch_command="convert" if explicit_batch else None,
        action="",
        to="md",
        files=[good, missing],
        batch=explicit_batch,
        json=json_mode,
        quiet=quiet,
        timeout=3,
        jobs=1,
    )
    monkeypatch.setattr(
        convert, "validate_files",
        lambda _files, **_kwargs: ([good], [(missing, "missing file")], []),
    )
    monkeypatch.setattr(convert, "validate_execution_options", lambda _args: None)
    monkeypatch.setattr(convert, "runtime_route_catalog", lambda _controller: object())
    monkeypatch.setattr(
        convert, "_resolve_runtime_routes",
        lambda *_args, **_kwargs: {good: SimpleNamespace(options=())},
    )
    monkeypatch.setattr(convert, "build_execution_options", lambda *_args, **_kwargs: {})

    calls: list[tuple[list[str], dict[str, object]]] = []

    def record_batch(_controller, _action, selected, *_args, **kwargs):
        calls.append((selected, kwargs))
        return 9

    monkeypatch.setattr(convert, "_execute_batch", record_batch)
    assert convert.execute_convert(args, SimpleNamespace(has_runtime=True)) == 9
    assert len(calls) == 1
    assert calls[0][0] == [good]
    assert calls[0][1]["invalid_files"] == [(missing, "missing file")]
    assert calls[0][1]["original_input_files"] == [good, missing]

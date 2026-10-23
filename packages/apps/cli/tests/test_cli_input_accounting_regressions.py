"""Regression tests for literal paths and batch admission accounting."""

from __future__ import annotations

import json
from argparse import Namespace
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock

import pytest

from docwen_cli.commands import convert
from docwen_cli.utils import expand_paths
from docwen_core.models.result import ConversionResult

from ._cli_contract_support import _make_execution_args
from .capability_fixtures import bundled_available_runtime_catalog

pytestmark = pytest.mark.integration


def test_literal_bracketed_file_takes_precedence_over_glob(tmp_path: Path) -> None:
    literal = tmp_path / "Report [1].md"
    other = tmp_path / "Report 1.md"
    literal.write_text("intended", encoding="utf-8")
    other.write_text("unrelated", encoding="utf-8")

    assert expand_paths([str(literal)]) == [str(literal.resolve())]
    assert expand_paths([f'"{literal}"']) == [str(literal.resolve())]
    other.unlink()
    assert expand_paths([str(literal)]) == [str(literal.resolve())]

    folder = tmp_path / "Archive [2026]"
    folder.mkdir()
    nested = folder / "note.md"
    nested.write_text("content", encoding="utf-8")
    assert expand_paths([str(folder)]) == [str(nested.resolve())]
    assert expand_paths([str(folder / "note.md")]) == [str(nested.resolve())]


def test_nonliteral_glob_still_expands_matching_files(tmp_path: Path) -> None:
    first = tmp_path / "Report 1.md"
    second = tmp_path / "Report 2.md"
    for source in (first, second):
        source.write_text("content", encoding="utf-8")
    assert expand_paths([str(tmp_path / "Report [12].md")]) == [str(first.resolve()), str(second.resolve())]


@pytest.mark.parametrize(("json_mode", "quiet"), [(False, False), (False, True), (True, False)])
@pytest.mark.parametrize("explicit_batch", [False, True])
@pytest.mark.parametrize("missing_first", [False, True])
def test_actual_batch_summary_counts_rejected_input_in_every_mode(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    capsys: pytest.CaptureFixture[str],
    json_mode: bool,
    quiet: bool,
    explicit_batch: bool,
    missing_first: bool,
) -> None:
    good = tmp_path / "good.md"
    good.write_text("# valid", encoding="utf-8")
    files = [str(good), str(tmp_path / "missing.md")]
    if missing_first:
        files.reverse()
    args = _make_execution_args(json=json_mode, quiet=quiet, batch=explicit_batch, files=files)
    args.command = "batch" if explicit_batch else "convert"
    args.batch_command = "convert" if explicit_batch else None
    args.to = "docx"
    controller = MagicMock()
    controller.has_runtime = True
    controller.execute_single.return_value = ConversionResult(task_id="good", success=True)
    catalog = bundled_available_runtime_catalog()
    monkeypatch.setattr(convert, "runtime_route_catalog", lambda _controller: catalog)

    assert convert.execute_convert(args, controller) == 9
    controller.execute_single.assert_called_once()
    assert controller.execute_single.call_args.args[0].input_refs[0].path == str(good)
    captured = capsys.readouterr()
    if json_mode:
        payload = json.loads(captured.out)
        assert payload["success"] is False
        assert payload["error"]["code"] == "batch_partial_failure"
        assert {key: payload["data"][key] for key in ("total", "succeeded", "failed")} == {
            "total": 2,
            "succeeded": 1,
            "failed": 1,
        }
        assert [item["input"] for item in payload["data"]["results"]] == files
        assert captured.err == ""
    elif quiet:
        assert captured.out == captured.err == ""
    else:
        assert "总计: 2 文件  |  成功: 1  |  失败: 1" in captured.out


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
        convert,
        "validate_files",
        lambda _files, **_kwargs: ([good], [(missing, "missing file")], []),
    )
    monkeypatch.setattr(convert, "validate_execution_options", lambda _args: None)
    monkeypatch.setattr(convert, "runtime_route_catalog", lambda _controller: object())
    monkeypatch.setattr(
        convert,
        "_resolve_runtime_routes",
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

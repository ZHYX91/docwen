"""CLI 单元测试。"""

from __future__ import annotations

from types import SimpleNamespace

import pytest

from docwen.cli.main import _run_action_for_files

pytestmark = pytest.mark.unit


class _Args(SimpleNamespace):
    def __getattr__(self, name: str):
        return None


def test_execute_headless_batch_flag_forces_batch(monkeypatch: pytest.MonkeyPatch) -> None:
    called = {"action": 0, "batch": 0}
    passed = {}

    monkeypatch.setattr("docwen.cli.utils.expand_paths", lambda paths: list(paths), raising=True)
    monkeypatch.setattr("docwen.cli.utils.validate_files", lambda files: (list(files), [], []), raising=True)
    monkeypatch.setattr(
        "docwen.cli.utils.create_progress_callback", lambda **_kwargs: lambda *_a, **_k: None, raising=True
    )

    def _stub_execute_action(*_args, **_kwargs):
        called["action"] += 1
        return 0

    def _stub_execute_batch(*_args, **_kwargs):
        called["batch"] += 1
        passed.update(_kwargs)
        return 0

    monkeypatch.setattr("docwen.cli.executor.execute_action", _stub_execute_action, raising=True)
    monkeypatch.setattr("docwen.cli.executor.execute_batch", _stub_execute_batch, raising=True)

    args = _Args(
        files=["a.docx"],
        quiet=True,
        verbose=False,
        json=False,
        continue_on_error=False,
        batch=True,
        jobs=3,
        timing=True,
    )

    assert _run_action_for_files("convert", args.files, options={}, args=args) == 0
    assert called == {"action": 0, "batch": 1}
    assert passed.get("max_workers") == 3
    assert passed.get("include_timing") is True


def test_execute_headless_without_batch_flag_uses_single_file_path(monkeypatch: pytest.MonkeyPatch) -> None:
    called = {"action": 0, "batch": 0}

    monkeypatch.setattr("docwen.cli.utils.expand_paths", lambda paths: list(paths), raising=True)
    monkeypatch.setattr("docwen.cli.utils.validate_files", lambda files: (list(files), [], []), raising=True)
    monkeypatch.setattr(
        "docwen.cli.utils.create_progress_callback", lambda **_kwargs: lambda *_a, **_k: None, raising=True
    )

    def _stub_execute_action(*_args, **_kwargs):
        called["action"] += 1
        return 0

    def _stub_execute_batch(*_args, **_kwargs):
        called["batch"] += 1
        return 0

    monkeypatch.setattr("docwen.cli.executor.execute_action", _stub_execute_action, raising=True)
    monkeypatch.setattr("docwen.cli.executor.execute_batch", _stub_execute_batch, raising=True)

    args = _Args(
        files=["a.docx"],
        quiet=True,
        verbose=False,
        json=False,
        continue_on_error=False,
        batch=False,
    )

    assert _run_action_for_files("convert", args.files, options={}, args=args) == 0
    assert called == {"action": 1, "batch": 0}

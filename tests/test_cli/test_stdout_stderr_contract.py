"""CLI 单元测试。"""

from __future__ import annotations

import pytest


pytestmark = pytest.mark.unit


def test_cli_run_does_not_print_to_stdout_in_main(
    monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]
) -> None:
    import importlib

    cli_main_module = importlib.import_module("docwen.cli.main")
    import docwen.cli_run as cli_run

    monkeypatch.setattr(cli_run, "initialize_app", lambda **_kwargs: (False, "warn"), raising=True)
    monkeypatch.setattr(cli_run, "init_logging_system", lambda: None, raising=True)
    monkeypatch.setattr(cli_main_module, "main", lambda: 0, raising=True)

    code = cli_run.main()
    captured = capsys.readouterr()

    assert code == 0
    assert captured.out == ""
    assert "警告" in captured.err

"""CLI 单元测试。"""

from __future__ import annotations

import io
import sys

import pytest

pytestmark = pytest.mark.unit


@pytest.mark.unit
def test_cli_run_early_reconfigure_stdio_handles_cp1252(monkeypatch: pytest.MonkeyPatch) -> None:
    import docwen.cli_run as cli_run

    out_buf = io.BytesIO()
    err_buf = io.BytesIO()
    fake_out = io.TextIOWrapper(out_buf, encoding="cp1252", errors="strict", line_buffering=True)
    fake_err = io.TextIOWrapper(err_buf, encoding="cp1252", errors="strict", line_buffering=True)

    monkeypatch.setattr(sys, "stdout", fake_out, raising=False)
    monkeypatch.setattr(sys, "stderr", fake_err, raising=False)

    cli_run._early_reconfigure_stdio()

    assert getattr(sys.stdout, "encoding", None) == "utf-8"
    assert getattr(sys.stderr, "encoding", None) == "utf-8"
    assert getattr(sys.stdout, "errors", None) == "replace"
    assert getattr(sys.stderr, "errors", None) == "replace"

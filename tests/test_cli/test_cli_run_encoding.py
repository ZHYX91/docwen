"""CLI 单元测试。"""

from __future__ import annotations

import builtins
import io
import sys
import types

import pytest

pytestmark = pytest.mark.unit


@pytest.mark.unit
def test_cli_runearly_reconfigure_stdio_handles_cp1252(monkeypatch: pytest.MonkeyPatch) -> None:
    from docwen_cli.main import ensure_console_utf8

    out_buf = io.BytesIO()
    err_buf = io.BytesIO()
    fake_out = io.TextIOWrapper(out_buf, encoding="cp1252", errors="strict", line_buffering=True)
    fake_err = io.TextIOWrapper(err_buf, encoding="cp1252", errors="strict", line_buffering=True)

    monkeypatch.setattr(sys, "stdout", fake_out, raising=False)
    monkeypatch.setattr(sys, "stderr", fake_err, raising=False)

    # Reset global flag so the function actually runs.
    import docwen_cli.main as cli_main

    monkeypatch.setattr(cli_main, "_CODEPAGE_SET", False, raising=True)

    ensure_console_utf8()

    assert getattr(sys.stdout, "encoding", None) == "utf-8"
    assert getattr(sys.stderr, "encoding", None) == "utf-8"
    assert getattr(sys.stdout, "errors", None) == "replace"
    assert getattr(sys.stderr, "errors", None) == "replace"


@pytest.mark.unit
def test_console_encoding_silently_handles_windows_codepage_failure(
    monkeypatch: pytest.MonkeyPatch, caplog: pytest.LogCaptureFixture
) -> None:
    """Windows 控制台代码页设置失败时应静默降级，仍正确配置 stdio 为 UTF-8。"""
    import docwen_cli.main as cli_main
    from docwen_cli.main import ensure_console_utf8

    out_buf = io.BytesIO()
    err_buf = io.BytesIO()
    fake_out = io.TextIOWrapper(out_buf, encoding="cp1252", errors="strict", line_buffering=True)
    fake_err = io.TextIOWrapper(err_buf, encoding="cp1252", errors="strict", line_buffering=True)

    monkeypatch.setattr(cli_main, "_CODEPAGE_SET", False, raising=True)
    monkeypatch.setattr(sys, "stdout", fake_out, raising=False)
    monkeypatch.setattr(sys, "stderr", fake_err, raising=False)
    monkeypatch.setattr(sys, "platform", "win32", raising=False)

    class _Kernel32:
        def SetConsoleOutputCP(self, _value):
            raise RuntimeError("cp fail")

        def SetConsoleCP(self, _value):
            raise AssertionError("should not reach SetConsoleCP after failure")

    monkeypatch.setitem(
        sys.modules, "ctypes", types.SimpleNamespace(windll=types.SimpleNamespace(kernel32=_Kernel32()))
    )

    with caplog.at_level("WARNING"):
        ensure_console_utf8()

    # 即使 codepage API 失败，stdio 仍应被重新配置为 UTF-8
    assert getattr(sys.stdout, "encoding", None) == "utf-8"
    assert getattr(sys.stderr, "encoding", None) == "utf-8"
    # 新入口不在此处记录日志警告 —— 由调用方或监控层负责


@pytest.mark.unit
def test_console_encoding_fail_fast_raises_when_stdio_init_fails(monkeypatch: pytest.MonkeyPatch) -> None:
    from docwen_cli.main import ensure_console_utf8

    original_import = builtins.__import__

    def _failing_import(name, globals=None, locals=None, fromlist=(), level=0):
        if name == "io":
            raise RuntimeError("io import failed")
        return original_import(name, globals, locals, fromlist, level)

    monkeypatch.setenv("DOCWEN_FAIL_FAST", "1")
    monkeypatch.setattr(builtins, "__import__", _failing_import)

    with pytest.raises(RuntimeError, match="io import failed"):
        ensure_console_utf8()

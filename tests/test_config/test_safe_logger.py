"""config 单元测试。"""

from __future__ import annotations

import io
import sys

import pytest

from docwen.config.safe_logger import SafeLogger, disable, enable, info, safe_log

pytestmark = pytest.mark.unit


@pytest.mark.unit
def test_safe_logger_disable_enable() -> None:
    safe_log.enable()
    assert safe_log.is_enabled() is True
    disable()
    assert safe_log.is_enabled() is False
    enable()
    assert safe_log.is_enabled() is True


@pytest.mark.unit
def test_safe_logger_log_does_not_crash(monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]) -> None:
    logger = SafeLogger()

    monkeypatch.setattr(logger, "_setup_logger", lambda: None, raising=True)
    monkeypatch.setattr(logger, "_logger", None, raising=False)
    logger.set_name("test_logger")

    logger.info("这是一条信息消息: %d", 42)
    out = capsys.readouterr()
    assert "test_logger" in out.out
    assert "INFO" in out.out


@pytest.mark.unit
def test_module_level_info_does_not_raise() -> None:
    info("x")


@pytest.mark.unit
def test_safe_logger_cp1252_stream_does_not_crash(monkeypatch: pytest.MonkeyPatch) -> None:
    buf = io.BytesIO()
    fake_out = io.TextIOWrapper(buf, encoding="cp1252", errors="strict", line_buffering=True)
    monkeypatch.setattr(sys, "stdout", fake_out, raising=False)
    monkeypatch.setattr(sys, "__stdout__", fake_out, raising=False)

    logger = SafeLogger()
    monkeypatch.setattr(logger, "_setup_logger", lambda: None, raising=True)
    monkeypatch.setattr(logger, "_logger", None, raising=False)
    logger.set_name("test_logger")

    logger.info("中文输出")
    fake_out.flush()
    captured = buf.getvalue().decode("cp1252", errors="strict")
    assert "中文" not in captured
    assert "?" in captured

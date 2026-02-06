from __future__ import annotations

import pytest

from docwen.config.safe_logger import SafeLogger, disable, enable, info, safe_log


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


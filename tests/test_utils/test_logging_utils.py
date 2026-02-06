from __future__ import annotations

import os
import time
from pathlib import Path

import pytest

import docwen.utils.logging_utils as logging_utils


@pytest.mark.unit
def test_pre_init_logging_resets_handlers() -> None:
    root = logging_utils.pre_init_logging()
    assert root is not None
    assert len(root.handlers) == 1


@pytest.mark.unit
def test_generate_log_path_uses_safe_prefix(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    src_dir = tmp_path / "src"
    src_dir.mkdir()
    script_path = src_dir / "run.py"
    script_path.write_text("x=1\n", encoding="utf-8")

    monkeypatch.setattr(logging_utils.sys, "argv", [str(script_path)])
    monkeypatch.setattr(
        logging_utils.config_manager,
        "get_logging_config",
        lambda: {"file_prefix": 'do:c*wen?<>|"', "retention_days": 7},
        raising=True,
    )

    log_path = logging_utils.generate_log_path()
    assert str(tmp_path).lower() in log_path.lower()
    assert "logs" in log_path.lower()
    base = os.path.basename(log_path)
    assert base.endswith(".log")
    assert base.startswith("docwen_")
    for ch in '\\/*?:"<>|':
        assert ch not in base


@pytest.mark.integration
def test_init_logging_system_and_clean_old_logs(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    logs_dir = tmp_path / "logs"
    logs_dir.mkdir()

    log_file = logs_dir / "docwen_20990101.log"
    log_file.write_text("x\n", encoding="utf-8")

    old_file = logs_dir / "docwen_20000101.log"
    old_file.write_text("old\n", encoding="utf-8")
    old_mtime = time.time() - 10 * 86400
    os.utime(old_file, (old_mtime, old_mtime))

    monkeypatch.setattr(logging_utils, "_logging_initialized", False, raising=True)
    monkeypatch.setattr(logging_utils, "generate_log_path", lambda: str(log_file), raising=True)
    monkeypatch.setattr(
        logging_utils.config_manager,
        "get_logging_config",
        lambda: {"enable": True, "console_enable": False, "level": "info", "retention_days": 1, "file_prefix": "docwen"},
        raising=True,
    )

    logger = logging_utils.init_logging_system()
    assert logger is not None
    assert any(h.__class__.__name__ == "TimedRotatingFileHandler" for h in logger.handlers)

    logging_utils.clean_old_logs()
    assert old_file.exists() is False

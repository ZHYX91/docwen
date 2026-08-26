"""Focused tests split from test_runtime_config_wiring.py."""

from __future__ import annotations

from ._runtime_config_wiring_support import (
    PROJECT_CONFIGS,
    Any,
    Path,
    logging,
    logging_handlers,
    os,
    pytest,
    tempfile,
)

pytestmark = pytest.mark.unit


class TestLoggingInit:
    """Verify init_logging() sets up full logging from config."""

    def test_init_with_console_only_config(self) -> None:
        from docwen_runtime.logging import init_logging

        config = {
            "logger": {
                "enable": False,  # file disabled
                "console_enable": True,
                "console_level": "WARNING",
                "level": "DEBUG",
            }
        }
        logger = init_logging(config)
        assert logger.name == "docwen"

    def test_init_full_logging(self) -> None:
        from docwen_runtime.logging import get_logging_state, init_logging

        config = {
            "logger": {
                "enable": True,
                "level": "DEBUG",
                "file_prefix": "test_docwen",
                "retention_days": 7,
                "console_enable": True,
                "console_level": "INFO",
                "directory_mode": "temp",
            }
        }
        init_logging(config)
        state = get_logging_state()
        assert state["file_enabled"] is True
        assert state["console_enabled"] is True

    def test_init_with_empty_config_uses_defaults(self) -> None:
        """Empty config should not crash — defaults to console-only."""
        from docwen_runtime.logging import get_logging_state, init_logging

        init_logging({})
        state = get_logging_state()
        assert state["console_enabled"] is True

    def test_console_colorize_always_and_never(self) -> None:
        from docwen_runtime.logging import init_logging

        logger = init_logging(
            {
                "logger": {
                    "enable": False,
                    "console_enable": True,
                    "console_level": "INFO",
                    "console_colorize": "always",
                }
            }
        )
        console_handlers = [h for h in logger.handlers if isinstance(h, logging.StreamHandler)]
        record = logger.makeRecord("docwen", logging.ERROR, __file__, 1, "boom", (), None)
        assert len(console_handlers) == 1
        formatter = console_handlers[0].formatter
        assert formatter is not None
        assert "\x1b[" in formatter.format(record)

        logger = init_logging(
            {
                "logger": {
                    "enable": False,
                    "console_enable": True,
                    "console_level": "INFO",
                    "console_colorize": "never",
                }
            }
        )
        console_handlers = [h for h in logger.handlers if isinstance(h, logging.StreamHandler)]
        record = logger.makeRecord("docwen", logging.ERROR, __file__, 1, "boom", (), None)
        assert len(console_handlers) == 1
        formatter = console_handlers[0].formatter
        assert formatter is not None
        assert "\x1b[" not in formatter.format(record)

    @pytest.mark.parametrize(
        "file_prefix",
        ["../outside", r"..\outside", r"C:\outside\audit", "/absolute/audit", "", "   "],
    )
    def test_log_file_prefix_is_filename_only_and_contained(
        self, tmp_path: Path, monkeypatch, file_prefix: str
    ) -> None:
        from docwen_runtime.logging import resolve_log_file_path

        monkeypatch.delenv("DOCWEN_LOG_DIR", raising=False)
        monkeypatch.delenv("DOCWEN_LOG_TO_TEMP", raising=False)
        log_dir = tmp_path / "logs"

        resolved = Path(
            resolve_log_file_path(
                {
                    "file_prefix": file_prefix,
                    "directory_mode": "custom",
                    "directory": str(log_dir),
                }
            )
        )

        assert resolved.parent == log_dir.resolve()
        assert resolved.name.endswith(".log")
        assert not any(character in resolved.stem for character in '\\/*?:"<>|')

    def test_actual_file_handler_uses_contained_normalized_prefix(self, tmp_path: Path, monkeypatch) -> None:
        from docwen_runtime.logging import get_logging_runtime_state, init_logging

        monkeypatch.delenv("DOCWEN_LOG_DIR", raising=False)
        monkeypatch.delenv("DOCWEN_LOG_TO_TEMP", raising=False)
        log_dir = tmp_path / "logs"
        logger = init_logging(
            {
                "logger": {
                    "enable": True,
                    "console_enable": False,
                    "file_prefix": "../outside",
                    "directory_mode": "custom",
                    "directory": str(log_dir),
                }
            }
        )
        logger.warning("contained-prefix-sentinel")
        state = get_logging_runtime_state()

        assert state.file_enabled is True
        assert state.active_log_file is not None
        active_path = Path(state.active_log_file)
        assert active_path.parent == log_dir.resolve()
        assert "contained-prefix-sentinel" in active_path.read_text(encoding="utf-8")
        assert not (tmp_path / "outside.log").exists()

    def test_retention_is_exact_and_does_not_evaluate_prefix_as_glob(self, tmp_path: Path) -> None:
        from docwen_runtime.logging import _purge_old_logs

        exact_old = tmp_path / "[audit].log.1"
        active_old = tmp_path / "[audit].log"
        unrelated_old = tmp_path / "a.log.1"
        for path in (exact_old, active_old, unrelated_old):
            path.write_text("old", encoding="utf-8")
            os.utime(path, (1, 1))

        _purge_old_logs(tmp_path, "[audit]", retention_days=1)

        assert not exact_old.exists()
        assert active_old.exists()
        assert unrelated_old.exists()

    @pytest.mark.parametrize("retention_days", [None, "7", 0, -1, True])
    def test_invalid_retention_falls_back_without_crashing(
        self, tmp_path: Path, monkeypatch, retention_days: object
    ) -> None:
        import docwen_runtime.logging as runtime_logging

        captured: list[int] = []
        original_purge = runtime_logging._purge_old_logs

        def _capture_purge(log_dir: Path, file_prefix: str, days: int) -> None:
            captured.append(days)
            original_purge(log_dir, file_prefix, days)

        monkeypatch.setattr(runtime_logging, "_purge_old_logs", _capture_purge)
        runtime_logging.init_logging(
            {
                "logger": {
                    "enable": True,
                    "console_enable": False,
                    "retention_days": retention_days,
                    "directory_mode": "custom",
                    "directory": str(tmp_path / "logs"),
                }
            }
        )

        assert captured == [runtime_logging.DEFAULT_RETENTION_DAYS]

    def test_invalid_directory_mode_falls_back_to_user_mode(self, tmp_path: Path, monkeypatch) -> None:
        import docwen_runtime.logging as runtime_logging

        monkeypatch.delenv("DOCWEN_LOG_DIR", raising=False)
        monkeypatch.delenv("DOCWEN_LOG_TO_TEMP", raising=False)
        monkeypatch.setattr(runtime_logging, "_resolve_log_dir", lambda _config: tmp_path / "logs")
        runtime_logging.init_logging(
            {
                "logger": {
                    "enable": True,
                    "console_enable": False,
                    "directory_mode": "outside-contract",
                }
            }
        )

        assert runtime_logging.get_logging_runtime_state().active_directory_mode == "user"

    def test_pre_init_records_replay_to_file_without_console_duplication(
        self, tmp_path: Path, monkeypatch, capsys: pytest.CaptureFixture[str]
    ) -> None:
        import docwen_runtime.logging as runtime_logging

        monkeypatch.delenv("DOCWEN_LOG_DIR", raising=False)
        monkeypatch.delenv("DOCWEN_LOG_TO_TEMP", raising=False)
        logger = runtime_logging.pre_init_logging("INFO")
        logger.warning("pre-init-lifecycle-sentinel")

        runtime_logging.init_logging(
            {
                "logger": {
                    "enable": True,
                    "console_enable": True,
                    "file_prefix": "lifecycle",
                    "directory_mode": "custom",
                    "directory": str(tmp_path / "logs"),
                }
            }
        )

        stderr = capsys.readouterr().err
        log_path = tmp_path / "logs" / "lifecycle.log"
        log_text = log_path.read_text(encoding="utf-8")
        assert stderr.count("pre-init-lifecycle-sentinel") == 1
        assert log_text.count("pre-init-lifecycle-sentinel") == 1
        assert not runtime_logging._pre_init_buffer

        runtime_logging.reconfigure_logging(
            {
                "logger": {
                    "enable": True,
                    "console_enable": True,
                    "file_prefix": "lifecycle",
                    "directory_mode": "custom",
                    "directory": str(tmp_path / "logs"),
                }
            }
        )
        logger.warning("post-reconfigure-sentinel")
        assert log_path.read_text(encoding="utf-8").count("post-reconfigure-sentinel") == 1


class TestLoggingReconfigure:
    """Verify reconfigure_logging() rebuilds handlers."""

    def test_reconfigure_changes_console_level(self) -> None:

        from docwen_runtime.logging import init_logging, reconfigure_logging

        init_logging({"logger": {"enable": False, "console_enable": True, "console_level": "INFO"}})

        # Reconfigure with different level
        logger2 = reconfigure_logging({"logger": {"enable": False, "console_enable": True, "console_level": "ERROR"}})
        assert logger2.name == "docwen"

    def test_reconfigure_from_config_loader(self) -> None:
        """End-to-end: read logging config from ConfigLoader, reconfigure."""
        from docwen_runtime.config.loader import ConfigLoader
        from docwen_runtime.logging import get_logging_state, reconfigure_logging

        with tempfile.TemporaryDirectory() as tmpdir:
            loader = ConfigLoader(base_dir=PROJECT_CONFIGS, user_dir=Path(tmpdir))
            config_dict = loader.config.as_dict()

            reconfigure_logging(config_dict)
            state = get_logging_state()
            # default config has enable=True for file logging
            assert state["file_enabled"] is True
            assert state["console_enabled"] is True

    def test_config_loader_reload_wires_logging(self) -> None:
        """ConfigLoader.reload() must wire logging via _wire_logging().

        After creating a ConfigLoader, the docwen logger must have
        handlers set up from the logger config section — proving
        that the logger.* config section is NOT dead config.
        """
        from docwen_runtime.config.loader import ConfigLoader
        from docwen_runtime.logging import get_logging_state

        with tempfile.TemporaryDirectory() as tmpdir:
            ConfigLoader(base_dir=PROJECT_CONFIGS, user_dir=Path(tmpdir))
            state = get_logging_state()

            # logger config section drives both file and console handlers
            assert state["file_enabled"] is True, "ConfigLoader.reload() must wire logging — file handler missing"
            assert state["console_enabled"] is True, "ConfigLoader.reload() must wire logging — console handler missing"
            assert state["active_log_file"] is not None, "ConfigLoader.reload() must create log file via init_logging"

    def test_runtime_state_reports_env_override_and_active_path(self, tmp_path: Path, monkeypatch) -> None:
        from docwen_runtime.logging import get_logging_runtime_state, init_logging, resolve_log_file_path

        monkeypatch.setenv("DOCWEN_LOG_DIR", str(tmp_path / "override"))
        init_logging(
            {
                "logger": {
                    "enable": True,
                    "console_enable": False,
                    "file_prefix": "audit",
                    "directory_mode": "custom",
                    "directory": str(tmp_path / "ignored"),
                }
            }
        )

        state = get_logging_runtime_state()

        assert state.file_enabled is True
        assert state.console_enabled is False
        assert state.overridden_by_env == "DOCWEN_LOG_DIR"
        assert state.active_directory_mode == "env"
        assert state.active_log_file == str(tmp_path / "override" / "logs" / "audit.log")
        assert resolve_log_file_path({"file_prefix": "audit", "directory_mode": "custom"}) == state.active_log_file

    def test_runtime_state_reports_file_handler_fallback(self, tmp_path: Path, monkeypatch) -> None:
        from docwen_runtime.logging import get_logging_runtime_state, init_logging

        monkeypatch.delenv("DOCWEN_LOG_DIR", raising=False)
        monkeypatch.delenv("DOCWEN_LOG_TO_TEMP", raising=False)

        def _raise_file_handler_error(*_args: object, **_kwargs: object) -> None:
            raise OSError("disk unavailable")

        monkeypatch.setattr(logging_handlers, "RotatingFileHandler", _raise_file_handler_error)

        init_logging(
            {
                "logger": {
                    "enable": True,
                    "console_enable": False,
                    "file_prefix": "audit",
                    "directory_mode": "custom",
                    "directory": str(tmp_path / "logs"),
                }
            }
        )

        state = get_logging_runtime_state()

        assert state.file_enabled is False
        assert state.console_enabled is True
        assert state.active_log_file is None
        assert state.active_directory_mode == "custom"
        assert state.fallback_used is True
        assert state.fallback_reason == "disk unavailable"
        assert state.overridden_by_env is None

    def test_primary_file_failure_falls_back_to_isolated_temp_directory(self, tmp_path: Path, monkeypatch) -> None:
        import docwen_runtime.logging as runtime_logging

        monkeypatch.delenv("DOCWEN_LOG_DIR", raising=False)
        monkeypatch.delenv("DOCWEN_LOG_TO_TEMP", raising=False)
        primary_dir = (tmp_path / "primary").resolve()
        fallback_root = tmp_path / "fallback-root"
        monkeypatch.setattr(runtime_logging.tempfile, "gettempdir", lambda: str(fallback_root))
        original_handler = logging_handlers.RotatingFileHandler

        def _selective_handler(
            filename: str,
            *args: Any,
            **kwargs: Any,
        ) -> logging_handlers.RotatingFileHandler:
            if Path(filename).parent == primary_dir:
                raise OSError("primary unavailable")
            return original_handler(filename, *args, **kwargs)

        monkeypatch.setattr(logging_handlers, "RotatingFileHandler", _selective_handler)
        logger = runtime_logging.init_logging(
            {
                "logger": {
                    "enable": True,
                    "console_enable": False,
                    "file_prefix": "audit",
                    "directory_mode": "custom",
                    "directory": str(primary_dir),
                }
            }
        )
        logger.warning("temp-fallback-sentinel")
        state = runtime_logging.get_logging_runtime_state()

        expected = (fallback_root / "docwen" / "logs" / "audit.log").resolve()
        assert state.file_enabled is True
        assert state.active_log_file == str(expected)
        assert state.active_directory_mode == "fallback_temp"
        assert state.fallback_used is True
        assert state.fallback_reason == "primary unavailable"
        assert "temp-fallback-sentinel" in expected.read_text(encoding="utf-8")

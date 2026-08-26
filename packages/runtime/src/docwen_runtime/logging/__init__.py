"""Runtime logging setup — pre-init, initialize, and reconfiguration.

Three-phase lifecycle:

* **pre_init_logging()** — console-only, minimal format, buffers records
  for later replay into the file handler.
* **init_logging(config)** — full console + file setup from config,
  replays buffered pre-init records into the file.
* **reconfigure_logging(config)** — tear-down and rebuild from changed config.

All three are safe to call multiple times — existing handlers are cleared first.

The ``config`` dict is read from the merged configuration tree
(``ConfigLoader.config.as_dict()``) and looks for the ``logger`` key.
"""

from __future__ import annotations

import logging
import logging.handlers
import os
import re
import sys
import tempfile
import threading
import time
from collections import deque
from dataclasses import dataclass
from pathlib import Path
from typing import Any

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

DEFAULT_LOG_FORMAT: str = "%(asctime)s.%(msecs)03d | %(levelname)s | %(name)s:%(lineno)d - %(message)s"
DEFAULT_DATE_FORMAT: str = "%Y-%m-%d %H:%M:%S"
PRE_INIT_CONSOLE_FORMAT: str = "[%(levelname)s] %(name)s: %(message)s"
DEFAULT_FILE_PREFIX: str = "docwen"
DEFAULT_RETENTION_DAYS: int = 30
MAX_LOG_BYTES: int = 10 * 1024 * 1024  # 10 MB per file
BACKUP_COUNT: int = 5
CONSOLE_COLORIZE_MODES: frozenset[str] = frozenset({"auto", "always", "never"})
DIRECTORY_MODES: frozenset[str] = frozenset({"user", "temp", "custom"})
_INVALID_FILE_PREFIX_CHARS = re.compile(r'[\\/*?:"<>|\x00-\x1f]')
LEVEL_COLORS: dict[int, str] = {
    logging.DEBUG: "\x1b[36m",
    logging.INFO: "\x1b[32m",
    logging.WARNING: "\x1b[33m",
    logging.ERROR: "\x1b[31m",
    logging.CRITICAL: "\x1b[35m",
}
RESET_COLOR: str = "\x1b[0m"

# Pre-init buffer capacity
_PRE_INIT_BUFFER_CAPACITY = 500

# Module-level mutable state — guarded by _state_lock
_state_lock = threading.Lock()
_console_handler: logging.StreamHandler | None = None
_file_handlers: list[logging.Handler] = []
_pre_init_buffer: deque[logging.LogRecord] = deque(maxlen=_PRE_INIT_BUFFER_CAPACITY)
_state: dict[str, Any] = {
    "file_enabled": False,
    "console_enabled": False,
    "active_log_file": None,
    "active_directory_mode": "user",
    "fallback_used": False,
    "fallback_reason": None,
    "overridden_by_env": None,
}


@dataclass(frozen=True)
class LoggingRuntimeState:
    """Read-only snapshot of the active runtime logging destination."""

    file_enabled: bool
    console_enabled: bool
    active_log_file: str | None
    active_directory_mode: str
    fallback_used: bool
    fallback_reason: str | None
    overridden_by_env: str | None


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _get_root() -> logging.Logger:
    """Return the docwen root logger."""
    return logging.getLogger("docwen")


def _clear_handlers(logger: logging.Logger) -> None:
    """Remove all handlers from *logger* and reset module-level references."""
    global _console_handler, _file_handlers, _state
    for h in list(logger.handlers):
        logger.removeHandler(h)
        h.close()
    _console_handler = None
    _file_handlers.clear()
    _state["file_enabled"] = False
    _state["console_enabled"] = False
    _state["active_log_file"] = None
    _state["active_directory_mode"] = "user"
    _state["fallback_used"] = False
    _state["fallback_reason"] = None
    _state["overridden_by_env"] = None


def _ensure_log_dir(log_dir: Path) -> Path:
    """Create *log_dir* if it doesn't exist, return it."""
    log_dir.mkdir(parents=True, exist_ok=True)
    return log_dir


def _env_override_source() -> str | None:
    env_dir = os.environ.get("DOCWEN_LOG_DIR", "").strip()
    if env_dir:
        return "DOCWEN_LOG_DIR"
    if os.environ.get("DOCWEN_LOG_TO_TEMP", "").strip().lower() in {"1", "true", "yes", "on"}:
        return "DOCWEN_LOG_TO_TEMP"
    return None


def _resolve_directory_mode(config: dict[str, Any]) -> str:
    override = _env_override_source()
    if override == "DOCWEN_LOG_DIR":
        return "env"
    if override == "DOCWEN_LOG_TO_TEMP":
        return "temp"
    configured = str(config.get("directory_mode", "user") or "user").strip().lower()
    return configured if configured in DIRECTORY_MODES else "user"


def _resolve_log_dir(config: dict[str, Any]) -> Path:
    """Resolve the logging directory from config or environment."""
    env_dir = os.environ.get("DOCWEN_LOG_DIR", "").strip()
    if env_dir:
        return Path(env_dir) / "logs"

    if os.environ.get("DOCWEN_LOG_TO_TEMP", "").strip().lower() in {"1", "true", "yes", "on"}:
        return Path(tempfile.gettempdir()) / "docwen" / "logs"

    directory_mode = _resolve_directory_mode(config)
    custom_dir = str(config.get("directory", "") or "").strip()

    if directory_mode == "custom" and custom_dir:
        return Path(custom_dir)

    if directory_mode == "temp":
        return Path(tempfile.gettempdir()) / "docwen" / "logs"

    # Default: user log directory via platformdirs
    try:
        import platformdirs

        return Path(platformdirs.user_log_dir("docwen", appauthor=False, ensure_exists=True))
    except ImportError:
        return Path(os.path.expanduser("~")) / ".docwen" / "logs"


def _resolve_log_level(value: object, default: str) -> int:
    """Resolve a log level string to an int, falling back to *default*."""
    if isinstance(value, str) and value.strip():
        return getattr(logging, value.strip().upper(), getattr(logging, default.upper()))
    return getattr(logging, default.upper(), logging.INFO)


def _normalize_console_colorize(value: object) -> str:
    if isinstance(value, str):
        normalized = value.strip().lower()
        if normalized in CONSOLE_COLORIZE_MODES:
            return normalized
    return "auto"


def _normalize_file_prefix(value: object) -> str:
    """Return a filename-only prefix that cannot escape the log directory."""
    normalized = _INVALID_FILE_PREFIX_CHARS.sub("", str(value or "")).strip().rstrip(". ")
    return normalized or DEFAULT_FILE_PREFIX


def _normalize_retention_days(value: object) -> int:
    """Accept only positive integer retention values; booleans are not integers here."""
    if isinstance(value, int) and not isinstance(value, bool) and value > 0:
        return value
    return DEFAULT_RETENTION_DAYS


def _resolve_log_path(log_dir: Path, file_prefix: str) -> Path:
    """Resolve a log path and enforce containment below *log_dir*."""
    resolved_dir = log_dir.resolve(strict=False)
    resolved_path = (resolved_dir / f"{file_prefix}.log").resolve(strict=False)
    try:
        resolved_path.relative_to(resolved_dir)
    except ValueError as exc:  # pragma: no cover - defense after prefix normalization
        raise ValueError("Log file path escapes the configured log directory") from exc
    return resolved_path


def _stream_supports_color(stream: Any) -> bool:
    isatty = getattr(stream, "isatty", None)
    if not callable(isatty):
        return False
    try:
        return bool(isatty())
    except Exception:
        return False


def _should_colorize(mode: str, stream: Any) -> bool:
    if mode == "always":
        return True
    if mode == "never":
        return False
    return _stream_supports_color(stream)


class _AnsiColorFormatter(logging.Formatter):
    """Formatter that wraps console lines in ANSI color by record level."""

    def format(self, record: logging.LogRecord) -> str:
        text = super().format(record)
        color = LEVEL_COLORS.get(record.levelno)
        if not color:
            return text
        return f"{color}{text}{RESET_COLOR}"


def _make_safe_formatter(
    fmt_str: str, datefmt: str = DEFAULT_DATE_FORMAT, *, colorize: bool = False
) -> logging.Formatter:
    """Create a Formatter from *fmt_str*, falling back to DEFAULT_LOG_FORMAT."""
    formatter_class: type[logging.Formatter] = _AnsiColorFormatter if colorize else logging.Formatter
    try:
        return formatter_class(fmt_str, datefmt=datefmt)
    except ValueError:
        return formatter_class(DEFAULT_LOG_FORMAT, datefmt=datefmt)


# ---------------------------------------------------------------------------
# Retention: purge log files older than *retention_days*
# ---------------------------------------------------------------------------


def _purge_old_logs(log_dir: Path, file_prefix: str, retention_days: int) -> None:
    """Delete log files under *log_dir* matching *file_prefix* older than *retention_days* days."""
    if retention_days <= 0 or not log_dir.is_dir():
        return
    cutoff = time.time() - retention_days * 86400
    resolved_dir = log_dir.resolve(strict=False)
    rotated_prefix = f"{file_prefix}.log."
    try:
        for entry in log_dir.iterdir():
            # The active log is never a retention candidate.  Rotated files
            # use the exact ``<prefix>.log.`` family; no user-derived glob is
            # evaluated here.
            if not entry.name.startswith(rotated_prefix):
                continue
            try:
                entry.resolve(strict=False).relative_to(resolved_dir)
            except ValueError:
                continue
            if entry.is_file() and entry.stat().st_mtime < cutoff:
                entry.unlink(missing_ok=True)
    except OSError:
        pass  # Permission errors or other FS issues are not fatal


# ---------------------------------------------------------------------------
# Pre-init buffer
# ---------------------------------------------------------------------------


def _buffer_pre_init_record(record: logging.LogRecord) -> None:
    """Buffer a log record for later replay into the file handler."""
    _pre_init_buffer.append(record)


def _replay_pre_init_buffer(logger: logging.Logger) -> None:
    """Replay buffered pre-init records once, into file handlers only."""
    del logger  # Kept in the signature for the private lifecycle call site.
    records = list(_pre_init_buffer)
    _pre_init_buffer.clear()
    for record in records:
        for handler in list(_file_handlers):
            if record.levelno >= handler.level:
                handler.handle(record)


class _PreInitBufferHandler(logging.Handler):
    """Captures log records into the pre-init buffer."""

    def emit(self, record: logging.LogRecord) -> None:
        _buffer_pre_init_record(record)


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def pre_init_logging(level: str = "INFO") -> logging.Logger:
    """Early console-only logging before configuration is loaded.

    Installs a ``StreamHandler`` on stderr and a buffer handler that
    captures records for later replay into the file handler.

    Args:
        level: Log level name (``"DEBUG"``, ``"INFO"``, etc.).

    Returns:
        The ``"docwen"`` root logger.
    """
    global _console_handler

    with _state_lock:
        logger = _get_root()
        _clear_handlers(logger)
        _pre_init_buffer.clear()

        logger.setLevel(getattr(logging, level.upper(), logging.INFO))

        _console_handler = logging.StreamHandler(sys.stderr)
        _console_handler.setLevel(getattr(logging, level.upper(), logging.INFO))
        _console_handler.setFormatter(logging.Formatter(PRE_INIT_CONSOLE_FORMAT))
        logger.addHandler(_console_handler)

        # Buffer handler — captures records for file replay
        logger.addHandler(_PreInitBufferHandler())

        _install_stdlib_bridge()
        _state["console_enabled"] = True

    logger.info("Logging pre-initialized (console only, level=%s).", level.upper())
    return logger


def init_logging(config: dict[str, Any] | None = None) -> logging.Logger:
    """Initialize full logging from a configuration dictionary.

    Reads the ``logger`` section of *config* to set up console + file
    handlers.  Buffered pre-init records are replayed into the file
    handler after setup.

    Args:
        config: Merged config dictionary (e.g. from
            ``ConfigLoader.config.as_dict()``).

    Returns:
        The ``"docwen"`` root logger.
    """
    return _apply_config(config, replay_pre_init=True)


def reconfigure_logging(config: dict[str, Any] | None = None) -> logging.Logger:
    """Re-read configuration and rebuild all logging handlers.

    Equivalent to :func:`init_logging` but explicitly signals handler
    rebuild from a potentially changed configuration.
    """
    return _apply_config(config, replay_pre_init=False)


def get_logging_state() -> dict[str, Any]:
    """Return a snapshot of the current logging state."""
    with _state_lock:
        return dict(_state)


def get_logging_runtime_state() -> LoggingRuntimeState:
    """Return the current runtime logging state as typed read-only data."""
    state = get_logging_state()
    return LoggingRuntimeState(
        file_enabled=bool(state.get("file_enabled", False)),
        console_enabled=bool(state.get("console_enabled", False)),
        active_log_file=state.get("active_log_file"),
        active_directory_mode=str(state.get("active_directory_mode") or "user"),
        fallback_used=bool(state.get("fallback_used", False)),
        fallback_reason=state.get("fallback_reason"),
        overridden_by_env=state.get("overridden_by_env"),
    )


def resolve_log_file_path(config: dict[str, Any] | None = None) -> str:
    """Resolve the log file path using the same rules as runtime logging."""
    log_cfg = _coerce_log_config(config)
    file_prefix = _normalize_file_prefix(log_cfg.get("file_prefix", DEFAULT_FILE_PREFIX))
    return str(_resolve_log_path(_resolve_log_dir(log_cfg), file_prefix))


# ---------------------------------------------------------------------------
# Internal: shared init / reconfigure logic
# ---------------------------------------------------------------------------


def _apply_config(config: dict[str, Any] | None, *, replay_pre_init: bool) -> logging.Logger:
    global _console_handler, _file_handlers, _state

    log_cfg = _coerce_log_config(config)

    # Resolve levels
    file_level = _resolve_log_level(log_cfg.get("level"), "DEBUG")
    console_level = _resolve_log_level(log_cfg.get("console_level"), "INFO")

    # Resolve log directory and file prefix
    file_prefix = _normalize_file_prefix(log_cfg.get("file_prefix", DEFAULT_FILE_PREFIX))
    retention_days = _normalize_retention_days(log_cfg.get("retention_days", DEFAULT_RETENTION_DAYS))
    fmt_str = str(log_cfg.get("format") or "").strip() or DEFAULT_LOG_FORMAT

    with _state_lock:
        logger = _get_root()
        _clear_handlers(logger)
        _install_stdlib_bridge()

        logger.setLevel(min(file_level, console_level))

        # ── Console handler ────────────────────────────────────────────
        console_enable = bool(log_cfg.get("console_enable", True))
        if console_enable:
            cfmt = str(log_cfg.get("console_format") or "").strip() or fmt_str or DEFAULT_LOG_FORMAT
            colorize_mode = _normalize_console_colorize(log_cfg.get("console_colorize", "auto"))
            _console_handler = logging.StreamHandler(sys.stderr)
            _console_handler.setLevel(console_level)
            _console_handler.setFormatter(
                _make_safe_formatter(cfmt, colorize=_should_colorize(colorize_mode, _console_handler.stream))
            )
            logger.addHandler(_console_handler)
            _state["console_enabled"] = True
        else:
            _state["console_enabled"] = False

        # ── File handler ───────────────────────────────────────────────
        file_enable = bool(log_cfg.get("enable", True))
        if file_enable:
            primary_mode = _resolve_directory_mode(log_cfg)
            primary_error: Exception | None = None
            try:
                log_dir = _resolve_log_dir(log_cfg)
                _ensure_log_dir(log_dir)
                log_path = _resolve_log_path(log_dir, file_prefix)
                _install_file_handler(logger, log_path, file_level, fmt_str)

                # Purge old logs
                _purge_old_logs(log_dir, file_prefix, retention_days)

                _state["file_enabled"] = True
                _state["active_log_file"] = str(log_path)
                _state["active_directory_mode"] = _resolve_directory_mode(log_cfg)
                _state["fallback_used"] = False
                _state["fallback_reason"] = None
                _state["overridden_by_env"] = _env_override_source()
            except Exception as exc:
                primary_error = exc

            if primary_error is not None:
                fallback_dir = Path(tempfile.gettempdir()) / "docwen" / "logs"
                fallback_path = _resolve_log_path(fallback_dir, file_prefix)
                primary_path = _resolve_log_path(_resolve_log_dir(log_cfg), file_prefix)
                fallback_error: Exception | None = None
                if fallback_path != primary_path:
                    try:
                        _ensure_log_dir(fallback_dir)
                        _install_file_handler(logger, fallback_path, file_level, fmt_str)
                        _purge_old_logs(fallback_dir, file_prefix, retention_days)
                        _state["file_enabled"] = True
                        _state["active_log_file"] = str(fallback_path)
                        _state["active_directory_mode"] = "fallback_temp"
                        _state["fallback_used"] = True
                        _state["fallback_reason"] = str(primary_error)
                        _state["overridden_by_env"] = _env_override_source()
                    except Exception as exc:
                        fallback_error = exc
                else:
                    fallback_error = primary_error

                if fallback_error is not None:
                    logger.warning("File log setup failed, continuing with console only: %s", primary_error)
                    _state["file_enabled"] = False
                    _state["active_log_file"] = None
                    _state["active_directory_mode"] = primary_mode
                    _state["fallback_used"] = True
                    _state["fallback_reason"] = str(primary_error)
                    _state["overridden_by_env"] = _env_override_source()
                    if _console_handler is None:
                        _console_handler = logging.StreamHandler(sys.stderr)
                        _console_handler.setLevel(logging.INFO)
                        _console_handler.setFormatter(logging.Formatter(PRE_INIT_CONSOLE_FORMAT))
                        logger.addHandler(_console_handler)
                        _state["console_enabled"] = True
        else:
            _state["file_enabled"] = False
            _state["active_log_file"] = None
            _state["active_directory_mode"] = _resolve_directory_mode(log_cfg)
            _state["fallback_used"] = False
            _state["fallback_reason"] = None
            _state["overridden_by_env"] = _env_override_source()

    # ── Replay pre-init buffer (outside lock to avoid re-entrancy) ─────
    if replay_pre_init:
        _replay_pre_init_buffer(logger)

    logger.info(
        "Logging initialized (file=%s, console=%s).",
        "on" if _state["file_enabled"] else "off",
        "on" if _state["console_enabled"] else "off",
    )
    return logger


def _coerce_log_config(config: dict[str, Any] | None) -> dict[str, Any]:
    raw = config or {}
    if isinstance(raw.get("logger"), dict):
        return dict(raw["logger"])
    return dict(raw)


def _install_file_handler(
    logger: logging.Logger,
    log_path: Path,
    file_level: int,
    fmt_str: str,
) -> None:
    """Create and attach one rotating file handler without leaking a failed handle."""
    handler: logging.handlers.RotatingFileHandler | None = None
    try:
        handler = logging.handlers.RotatingFileHandler(
            str(log_path),
            maxBytes=MAX_LOG_BYTES,
            backupCount=BACKUP_COUNT,
            encoding="utf-8",
        )
        handler.setLevel(file_level)
        handler.setFormatter(_make_safe_formatter(fmt_str))
        logger.addHandler(handler)
        _file_handlers.append(handler)
    except Exception:
        if handler is not None:
            handler.close()
        raise


# ---------------------------------------------------------------------------
# Internal: stdlib bridge
# ---------------------------------------------------------------------------


def _install_stdlib_bridge() -> None:
    """Ensure stdlib logging messages flow to the docwen logger."""
    root = logging.getLogger()
    for h in root.handlers:
        if isinstance(h, _StdlibBridgeHandler):
            return
    root.addHandler(_StdlibBridgeHandler())
    root.setLevel(logging.DEBUG)

    docwen_logger = logging.getLogger("docwen")
    docwen_logger.propagate = False


class _StdlibBridgeHandler(logging.Handler):
    """Forwards stdlib log records to the docwen logger."""

    def emit(self, record: logging.LogRecord) -> None:
        docwen_logger = logging.getLogger("docwen")
        if record.name.startswith("docwen"):
            return
        docwen_logger.handle(record)

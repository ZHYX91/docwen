"""GUI bootstrap: single-instance ownership and runtime/control forwarding."""

from __future__ import annotations

import logging
import os
import sys
import time
from pathlib import Path
from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from docwen_runtime.ipc.single_instance import SingleInstance

logger = logging.getLogger(__name__)

ENV_DISABLE_CONTROL = "DOCWEN_GUI_DISABLE_CONTROL"
ENV_TEST_AUTOCLOSE_MS = "DOCWEN_GUI_TEST_AUTOCLOSE_MS"


class GuiBootstrapDecision:
    """Result of primary-instance ownership and control forwarding."""

    def __init__(
        self,
        should_start_gui: bool = True,
        should_exit: bool = False,
        exit_code: int = 0,
        files_to_add: list[str] | None = None,
        instance_lock: SingleInstance | None = None,
    ) -> None:
        self.should_start_gui = should_start_gui
        self.should_exit = should_exit
        self.exit_code = exit_code
        self.files_to_add = files_to_add or []
        self.instance_lock = instance_lock


def should_disable_control() -> bool:
    """Return whether test/debug mode deliberately bypasses single-instance control."""

    if _env_bool(ENV_DISABLE_CONTROL):
        return True
    if _test_autoclose_ms() > 0:
        return True
    return _is_debug_session()


def parse_startup_files(argv: list[str] | None = None) -> list[str]:
    """Return existing positional files as canonical absolute paths."""

    args = argv if argv is not None else sys.argv
    files: list[str] = []
    for arg in args[1:]:
        if arg.startswith("-"):
            continue
        path = Path(arg).resolve()
        if path.is_file():
            files.append(str(path))
        else:
            logger.debug("Startup argument is not a valid file: %s", arg)
    return files


def bootstrap_gui(app_name: str = "docwen", argv: list[str] | None = None) -> GuiBootstrapDecision:
    """Acquire primary ownership or forward to the existing GUI control endpoint."""

    files = parse_startup_files(argv)
    if should_disable_control():
        _sync_runtime_lock_disabled()
        return GuiBootstrapDecision(files_to_add=files)

    try:
        from docwen_runtime.control import ControlClient, ControlError
        from docwen_runtime.ipc.single_instance import SingleInstance
    except ImportError as exc:
        logger.error("Required runtime/control components are unavailable: %s", exc)
        return GuiBootstrapDecision(should_start_gui=False, should_exit=True, exit_code=1)

    instance_lock = SingleInstance(app_name)
    if instance_lock.acquire():
        return GuiBootstrapDecision(files_to_add=files, instance_lock=instance_lock)

    client = ControlClient(app_name=app_name)
    deadline = time.monotonic() + 5.0
    try:
        if files:
            for file_path in files:
                _send_when_control_ready(client, "open", {"file": file_path}, deadline=deadline)
        else:
            _send_when_control_ready(client, "activate", {}, deadline=deadline)
    except ControlError as exc:
        logger.error("Unable to forward request to the running GUI: %s", exc)
        instance_lock.close()
        return GuiBootstrapDecision(should_start_gui=False, should_exit=True, exit_code=1)

    instance_lock.close()
    return GuiBootstrapDecision(should_start_gui=False, should_exit=True, exit_code=0)


def _send_when_control_ready(client: Any, action: str, payload: dict[str, object], *, deadline: float) -> None:
    """Retry only the primary-startup race before a request is delivered."""

    from docwen_runtime.control import ControlNotRunningError, ControlTimeoutError

    while True:
        remaining = deadline - time.monotonic()
        if remaining <= 0:
            raise ControlTimeoutError("Timed out waiting for the primary GUI control endpoint.")
        attempt_deadline = min(deadline, time.monotonic() + 1.0)
        attempt_timeout = attempt_deadline - time.monotonic()
        if attempt_timeout <= 0:
            raise ControlTimeoutError("Timed out waiting for the primary GUI control endpoint.")
        attempt_payload = dict(payload)
        attempt_payload["_deadline_monotonic"] = attempt_deadline
        try:
            client.request(action, attempt_payload, timeout=attempt_timeout)
            return
        except ControlNotRunningError as exc:
            sleep_remaining = deadline - time.monotonic()
            if sleep_remaining <= 0:
                raise ControlTimeoutError("Timed out waiting for the primary GUI control endpoint.") from exc
            time.sleep(min(0.05, sleep_remaining))


def _env_bool(name: str) -> bool:
    return os.environ.get(name, "").strip().lower() in {"1", "true", "yes", "y", "on"}


def _test_autoclose_ms() -> int:
    raw = os.environ.get(ENV_TEST_AUTOCLOSE_MS, "0") or "0"
    try:
        return int(raw)
    except ValueError:
        logger.warning("%s value is not an integer: %r", ENV_TEST_AUTOCLOSE_MS, raw)
        return 0


def _is_debug_session() -> bool:
    gettrace = getattr(sys, "gettrace", None)
    return bool((callable(gettrace) and gettrace() is not None) or os.environ.get("DEBUGPY_LAUNCHER_PORT"))


def _sync_runtime_lock_disabled() -> None:
    try:
        from docwen_runtime.ipc import disable_ipc

        disable_ipc()
    except ImportError:
        pass


__all__ = [
    "ENV_DISABLE_CONTROL",
    "ENV_TEST_AUTOCLOSE_MS",
    "GuiBootstrapDecision",
    "bootstrap_gui",
    "parse_startup_files",
    "should_disable_control",
]

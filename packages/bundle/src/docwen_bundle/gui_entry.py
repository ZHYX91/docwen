from __future__ import annotations

import logging
import os
import sys
import threading
import time
from collections.abc import Callable
from pathlib import Path
from typing import TYPE_CHECKING, Any, cast

from docwen_runtime.security import NetworkGuardInstallationError, dependency_egress_guard

if TYPE_CHECKING:
    from docwen_core.models import TaskEvent

logger = logging.getLogger(__name__)


def _make_event_callback(event_sink: Callable[[str, dict[str, Any]], None]) -> Callable[[TaskEvent], None]:
    def _callback(event: TaskEvent) -> None:
        payload = {"task_id": event.task_id, **event.payload}
        event_sink(event.event_type, payload)

    return _callback


def _install_gui_control_poll_timer(app: Any, drain_pending_commands: Callable[[], None]) -> object:
    from PySide6.QtCore import QTimer

    timer = QTimer(app)
    timer.setInterval(25)
    timer.timeout.connect(drain_pending_commands)
    timer.start()
    return timer


def _supported_settings_sections(window: Any) -> list[str]:
    provider = getattr(window, "supported_settings_sections", None)
    if not callable(provider):
        return []
    raw_sections = provider()
    if not isinstance(raw_sections, (list, tuple)):
        return []
    return [item for item in raw_sections if isinstance(item, str)]


def _control_wait_timeout(deadline: float | None) -> float:
    """Return the queue wait budget without shortening an explicit deadline."""

    if deadline is None:
        return 15.0
    return max(0.0, deadline - time.monotonic())


def _start_gui_control(window: Any, *, app: Any) -> object:
    """Start runtime/control and marshal every GUI action onto the Qt thread."""

    from docwen_core.version import PRODUCT_VERSION
    from docwen_runtime.control import (
        CONTROL_PROTOCOL_VERSION,
        ControlRequestError,
        ControlServer,
    )

    pending: list[
        tuple[
            str,
            dict[str, Any],
            float | None,
            threading.Event,
            threading.Event,
            dict[str, Any],
        ]
    ] = []
    pending_lock = threading.Lock()
    stopping = threading.Event()

    def _stopping_error() -> ControlRequestError:
        return ControlRequestError("gui_stopping", "DocWen GUI is stopping.")

    def _handle(action: str, payload: dict[str, Any]) -> dict[str, Any]:
        deadline_raw = payload.get("_deadline_monotonic")
        deadline = (
            float(deadline_raw)
            if isinstance(deadline_raw, (int, float)) and not isinstance(deadline_raw, bool)
            else None
        )
        completed = threading.Event()
        cancelled = threading.Event()
        response: dict[str, Any] = {}
        wait_timeout = _control_wait_timeout(deadline)
        if wait_timeout <= 0:
            raise ControlRequestError("control_timeout", "GUI control request expired before it was queued.")
        with pending_lock:
            if stopping.is_set():
                raise _stopping_error()
            pending.append((action, payload, deadline, completed, cancelled, response))
        if not completed.wait(wait_timeout):
            cancelled.set()
            raise ControlRequestError("control_timeout", "GUI thread did not accept the control request in time.")
        error = response.get("error")
        if isinstance(error, ControlRequestError):
            raise error
        data = response.get("data")
        return data if isinstance(data, dict) else {}

    def _drain() -> None:
        with pending_lock:
            commands = list(pending)
            pending.clear()
        for action, payload, deadline, completed, cancelled, response in commands:
            try:
                if stopping.is_set():
                    raise _stopping_error()
                if cancelled.is_set() or (deadline is not None and time.monotonic() >= deadline):
                    raise ControlRequestError(
                        "control_timeout",
                        "GUI control request expired before the GUI thread could apply it.",
                    )
                if action == "status":
                    settings_sections = _supported_settings_sections(window)
                    supported_actions = ["status", "activate", "open"]
                    if settings_sections and callable(getattr(window, "open_settings", None)):
                        supported_actions.append("open_settings")
                    response["data"] = {
                        "state": "running",
                        "running": True,
                        "control_ready": True,
                        "available": True,
                        "control_protocol_version": CONTROL_PROTOCOL_VERSION,
                        "product_version": PRODUCT_VERSION,
                        "pid": os.getpid(),
                        "window_visible": bool(window.isVisible()),
                        "supported_actions": supported_actions,
                        "settings_sections": settings_sections,
                    }
                elif action == "activate":
                    window.handle_ipc_command("activate", None)
                    response["data"] = {"accepted": True, "running": True, "action": "activate"}
                elif action == "open":
                    raw_path = payload.get("file")
                    if not isinstance(raw_path, str) or not raw_path:
                        raise ControlRequestError("invalid_path", "GUI open requires a file path.")
                    path = Path(raw_path)
                    if not path.is_absolute() or not path.is_file():
                        raise ControlRequestError("file_not_found", "The requested GUI file does not exist.")
                    resolved = str(path.resolve())
                    window.handle_ipc_command("open_file", resolved)
                    response["data"] = {
                        "accepted": True,
                        "running": True,
                        "action": "open",
                        "file": resolved,
                    }
                elif action == "open_settings":
                    raw_section = payload.get("section")
                    if not isinstance(raw_section, str) or not raw_section:
                        raise ControlRequestError("invalid_arguments", "GUI open-settings requires a section.")
                    settings_sections = _supported_settings_sections(window)
                    open_settings = getattr(window, "open_settings", None)
                    if not callable(open_settings):
                        raise ControlRequestError(
                            "capability_unavailable",
                            "The running DocWen GUI does not support opening settings.",
                            details={"required_action": "open_settings", "restart_required": True},
                        )
                    if raw_section not in settings_sections:
                        raise ControlRequestError(
                            "settings_section_unavailable",
                            "The requested GUI settings section is unavailable.",
                            details={"section": raw_section, "available_sections": settings_sections},
                        )
                    result = open_settings(raw_section, deadline=deadline)
                    if isinstance(result, dict) and result.get("error_code") == "control_timeout":
                        raise ControlRequestError(
                            "control_timeout",
                            "GUI settings request expired before the dialog could be shown.",
                            details={"section": raw_section},
                        )
                    if not isinstance(result, dict) or result.get("accepted") is not True:
                        raise ControlRequestError(
                            "settings_section_unavailable",
                            "The requested GUI settings section could not be loaded.",
                            details={"section": raw_section, "available_sections": settings_sections},
                        )
                    response["data"] = {
                        "accepted": True,
                        "running": True,
                        "action": "open_settings",
                        "section": raw_section,
                        "reused": bool(result.get("reused")),
                    }
                else:
                    raise ControlRequestError("invalid_control_action", "Unsupported GUI control action.")
            except ControlRequestError as exc:
                response["error"] = exc
            except Exception:
                response["error"] = ControlRequestError(
                    "gui_command_failed", "DocWen GUI could not process the request."
                )
            finally:
                completed.set()

    def _cancel_pending_on_quit() -> None:
        stopping.set()
        with pending_lock:
            commands = list(pending)
            pending.clear()
        for _action, _payload, _deadline, completed, cancelled, response in commands:
            cancelled.set()
            response["error"] = _stopping_error()
            completed.set()

    server = ControlServer(_handle, app_name="docwen")
    server.start()
    timer = _install_gui_control_poll_timer(app, _drain)
    cast(Any, server)._docwen_gui_control_timer = timer
    app.aboutToQuit.connect(_cancel_pending_on_quit)
    app.aboutToQuit.connect(cast(Any, timer).stop)
    app.aboutToQuit.connect(server.begin_stop)
    return server


def _stop_gui_control_and_release_instance_lock(
    control_server: object | None,
    instance_lock: object | None,
) -> None:
    """Finish control cleanup before another GUI process can acquire the lock."""

    try:
        if control_server is not None:
            cast(Any, control_server).stop()
    finally:
        if instance_lock is not None:
            cast(Any, instance_lock).release()


def _main_with_guard_active(argv: list[str] | None = None) -> int:
    from docwen_application.controller import ApplicationController
    from docwen_bundle.config_port import ConfigPortAdapter
    from docwen_bundle.gui_bootstrap import bootstrap_gui
    from docwen_bundle.runtime_factory import create_runtime_port
    from docwen_gui.app import (
        _initialize_application_theme,
        _schedule_test_autoclose,
        _schedule_test_ipc_report,
        _schedule_test_notification_report,
        _schedule_test_ocr_report,
        create_main_window,
        create_qapplication,
    )
    from docwen_gui.i18n import set_locale
    from docwen_gui.qt_bridge.task_event_bridge import TaskEventBridge
    from docwen_gui.release_smoke import _schedule_test_conversion_report
    from docwen_gui.settings_smoke import _schedule_test_settings_report
    from docwen_runtime.config import ConfigLoader

    args = argv if argv is not None else sys.argv
    decision = bootstrap_gui(app_name="docwen", argv=args)
    if decision.should_exit:
        return decision.exit_code

    app = create_qapplication(args)

    task_event_bridge = TaskEventBridge()
    config_loader = ConfigLoader()
    config_port = ConfigPortAdapter(config_loader)
    configured_locale = config_port.get("gui.language.locale", "zh_CN")
    set_locale(configured_locale if isinstance(configured_locale, str) else "zh_CN")
    runtime_port = create_runtime_port(
        config_loader=config_loader,
        event_callback=_make_event_callback(task_event_bridge.enqueue),
    )
    controller = ApplicationController(
        runtime_port=runtime_port,
        config_port=config_port,
    )
    controller.start()
    _initialize_application_theme(app, controller)

    window = create_main_window(
        controller=controller,
        task_event_bridge=task_event_bridge,
        initial_files=decision.files_to_add,
    )

    control_server: object | None = None
    instance_lock = decision.instance_lock
    if instance_lock is not None:
        try:
            control_server = _start_gui_control(window, app=app)
        except Exception:
            instance_lock.release()
            raise

    _schedule_test_autoclose(app, window)
    _schedule_test_notification_report(app, window)
    _schedule_test_ocr_report(app, window)
    _schedule_test_conversion_report(app, window)
    _schedule_test_settings_report(app, window)
    _schedule_test_ipc_report(app, window)
    window.show()
    try:
        return app.exec()
    finally:
        _stop_gui_control_and_release_instance_lock(control_server, instance_lock)


def main(argv: list[str] | None = None) -> int:
    """Run the composed GUI with dependency egress protection enforced."""

    from docwen_runtime.logging import pre_init_logging

    pre_init_logging("INFO")

    try:
        with dependency_egress_guard():
            return _main_with_guard_active(argv)
    except ModuleNotFoundError as exc:
        missing_root = str(exc.name or "").partition(".")[0]
        if missing_root not in {"PySide6", "qfluentwidgets"}:
            raise
        # Keep damaged/incomplete source installs on the same stable dependency
        # category as protocol 3 without importing the CLI package into the GUI
        # executable.  The message is deliberately bounded and traceback-free.
        print(
            f"DocWen GUI cannot start: required dependency '{missing_root}' is missing.",
            file=sys.stderr,
        )
        return 4
    except NetworkGuardInstallationError:
        logger.critical("安全检查失败")
        return 1

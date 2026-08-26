"""DocWen GUI application entry point.

Creates the QApplication, initialises theme, runs security startup checks,
and launches the main window.
"""

from __future__ import annotations

import logging
import sys
from typing import TYPE_CHECKING, cast

if TYPE_CHECKING:
    from PySide6.QtWidgets import QApplication

    from docwen_application.controller import ApplicationController

    from .main_window import MainWindow

# ── QApplication metadata ───────────────────────────────────────────
_APP_NAME = "DocWen"
_OFFLINE_VERSION = "Offline"

logger = logging.getLogger(__name__)


# ── Security startup helpers ──────────────────────────────────────────


def _run_gui_security_startup() -> None:
    """Execute security startup protections for GUI.

    Strict failures raise :exc:`SecurityCheckFailedError`; degraded
    warnings are logged.  Callers should catch the exception and abort
    startup.
    """
    from docwen_runtime.security import resolve_strict_security, run_security_protections

    strict = resolve_strict_security()
    degraded = run_security_protections(logger=logger, strict_security=strict)
    if degraded is not None:
        logger.warning(degraded)
    logger.debug("GUI security startup check complete (strict=%s)", strict)


def create_qapplication(argv: list[str] | None = None) -> QApplication:
    """Create (or return existing) QApplication instance.

    Ensures exactly one QApplication exists for the process lifetime.
    """
    from PySide6.QtWidgets import QApplication

    existing = QApplication.instance()
    if isinstance(existing, QApplication):
        return existing
    if existing is not None:
        raise RuntimeError("Existing application instance is not a QApplication — cannot start GUI.")

    args = argv if argv is not None else sys.argv
    app = QApplication(args)
    app.setApplicationName(_APP_NAME)
    app.setApplicationVersion(_OFFLINE_VERSION)
    from .font_utils import apply_application_font

    apply_application_font(app)
    return app


def _initialize_application_theme(app: object, controller: object | None) -> None:
    """Initialize the global theme from the injected configuration source."""
    from docwen_gui.styles.theme_manager import ThemeManager
    from docwen_gui.styles.theme_semantics import DEFAULT_THEME

    theme_name: object = DEFAULT_THEME
    config_port = getattr(controller, "config_port", None)
    if config_port is not None:
        try:
            theme_name = config_port.get("gui.theme.default_theme", DEFAULT_THEME)
        except Exception:
            logger.warning("Unable to read configured GUI theme; using %s", DEFAULT_THEME)

    if not isinstance(theme_name, str) or theme_name not in ThemeManager.get_available_themes():
        theme_name = DEFAULT_THEME
    ThemeManager.get_instance().initialize(cast("QApplication", app), theme_name)


def create_main_window(
    *,
    controller: ApplicationController | None = None,
    task_event_bridge: object | None = None,
    initial_files: list[str] | None = None,
) -> MainWindow:
    """Create and configure the main window.

    Args:
        controller: An optional pre-configured ApplicationController.
            If None, the window starts without a runtime backend (limited mode).
        initial_files: Optional list of file paths to load on startup
            (e.g. from command-line args or IPC forwarding).

    Returns:
        A configured MainWindow instance (not yet shown).
    """
    from .main_window import MainWindow
    from .qt_bridge.task_event_bridge import TaskEventBridge
    from .view_models.main_window_vm import MainWindowViewModel

    if task_event_bridge is None:
        task_event_bridge = TaskEventBridge()
    bridge = task_event_bridge if isinstance(task_event_bridge, TaskEventBridge) else TaskEventBridge()

    view_model = MainWindowViewModel(controller=controller)
    window = MainWindow(
        view_model=view_model,
        task_event_bridge=bridge,
    )  # setup_ui() is called inside __init__

    if initial_files:
        view_model.add_files(initial_files)

    return window


def _schedule_test_autoclose(app: QApplication, window: MainWindow) -> None:
    """Schedule graceful window close via DOCWEN_GUI_TEST_AUTOCLOSE_MS.

    Used by all entry points to support headless / smoke testing.
    """
    import os

    try:
        auto_close_ms = int(os.environ.get("DOCWEN_GUI_TEST_AUTOCLOSE_MS", "0") or "0")
    except ValueError:
        import logging

        logging.getLogger(__name__).warning("DOCWEN_GUI_TEST_AUTOCLOSE_MS 非法，已忽略")
        return
    if auto_close_ms > 0:
        from PySide6.QtCore import QTimer

        QTimer.singleShot(auto_close_ms, window.close)


def _schedule_test_notification_report(app: QApplication, window: MainWindow) -> None:
    """Write a packaged-notification smoke report when requested by CI.

    This is an internal test hook. It proves that the packaged GUI can reach the
    Qt tray-message path, but it does not claim Windows notification-center
    delivery or user-visible toast persistence. An explicit
    DOCWEN_GUI_TEST_NOTIFICATION_HOLD_MS value keeps the probe-owned tray icon
    alive long enough for a physical acceptance observer without changing the
    normal or default-smoke lifecycle.
    """
    import json
    import os
    from pathlib import Path

    report_path_raw = os.environ.get("DOCWEN_GUI_TEST_NOTIFICATION_REPORT", "").strip()
    if not report_path_raw:
        return

    try:
        probe_hold_ms = max(
            0,
            int(os.environ.get("DOCWEN_GUI_TEST_NOTIFICATION_HOLD_MS", "0") or "0"),
        )
    except ValueError:
        probe_hold_ms = 0
    message_timeout_ms = max(500, probe_hold_ms)

    def _probe() -> None:
        from PySide6.QtWidgets import QSystemTrayIcon

        report: dict[str, object] = {
            "isSystemTrayAvailable": False,
            "supportsMessages": False,
            "defaultTrayIconPresent": False,
            "probeCreatedTrayIcon": False,
            "hasTrayIcon": False,
            "showMessageCalled": False,
            "probeHoldMs": probe_hold_ms,
            "messageTimeoutMs": message_timeout_ms,
            "error": None,
        }
        tray = None
        probe_created_tray = False
        try:
            report["isSystemTrayAvailable"] = bool(QSystemTrayIcon.isSystemTrayAvailable())
            report["supportsMessages"] = bool(QSystemTrayIcon.supportsMessages())
            tray = getattr(window, "_system_tray_icon", None)
            report["defaultTrayIconPresent"] = tray is not None
            if tray is None and report["isSystemTrayAvailable"]:
                tray = QSystemTrayIcon(window.windowIcon(), window)
                tray.setVisible(True)
                probe_created_tray = True
            report["probeCreatedTrayIcon"] = probe_created_tray
            report["hasTrayIcon"] = tray is not None
            if tray is not None and report["isSystemTrayAvailable"]:
                tray.showMessage(
                    "DocWen notification smoke",
                    "Packaged notification smoke",
                    QSystemTrayIcon.MessageIcon.Information,
                    message_timeout_ms,
                )
                report["showMessageCalled"] = True
            app.processEvents()
        except Exception as exc:  # pragma: no cover - environment-specific smoke diagnostics
            report["error"] = f"{type(exc).__name__}: {exc}"
        finally:
            if probe_created_tray and tray is not None:
                if probe_hold_ms > 0:
                    from PySide6.QtCore import QTimer

                    def _cleanup_probe_tray() -> None:
                        tray.hide()
                        tray.deleteLater()

                    QTimer.singleShot(probe_hold_ms, _cleanup_probe_tray)
                else:
                    tray.hide()
                    tray.deleteLater()
            report_path = Path(report_path_raw)
            report_path.parent.mkdir(parents=True, exist_ok=True)
            report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")

    from PySide6.QtCore import QTimer

    QTimer.singleShot(250, _probe)


def _schedule_test_ocr_report(app: QApplication, window: MainWindow) -> None:
    """Run a packaged GUI image-OCR smoke workflow when requested by CI.

    This internal hook drives the same MainWindow/ActionArea route as the GUI
    e2e tests. It is intentionally env-gated so normal GUI startup is unchanged.
    """
    import json
    import os
    import time
    from pathlib import Path

    report_path_raw = os.environ.get("DOCWEN_GUI_TEST_OCR_REPORT", "").strip()
    input_path_raw = os.environ.get("DOCWEN_GUI_TEST_OCR_INPUT", "").strip()
    if not report_path_raw and not input_path_raw:
        return

    report_path = Path(report_path_raw) if report_path_raw else None
    source_path = Path(input_path_raw) if input_path_raw else None
    output_dir_raw = os.environ.get("DOCWEN_GUI_TEST_OCR_OUTPUT_DIR", "").strip()
    expected_text = os.environ.get("DOCWEN_GUI_TEST_OCR_EXPECTED_TEXT", "HELLO DOCWEN OCR")
    try:
        timeout_ms = int(os.environ.get("DOCWEN_GUI_TEST_OCR_TIMEOUT_MS", "60000") or "60000")
    except ValueError:
        timeout_ms = 60000
    timeout_ms = max(timeout_ms, 1000)

    started_at = time.monotonic()
    state: dict[str, object] = {"conversion_started": False, "normalized_path": ""}

    def _write_report(payload: dict[str, object]) -> None:
        if report_path is None:
            return
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    def _finish(payload: dict[str, object]) -> None:
        elapsed_ms = int((time.monotonic() - started_at) * 1000)
        payload.setdefault("elapsedMs", elapsed_ms)
        _write_report(payload)
        window.close()

    def _fail(error: str, **extra: object) -> None:
        _finish(
            {
                "success": False,
                "status": "failed",
                "error": error,
                "expectedText": expected_text,
                **extra,
            }
        )

    def _timed_out() -> bool:
        return (time.monotonic() - started_at) * 1000 > timeout_ms

    def _start_conversion() -> None:
        from PySide6.QtCore import QTimer

        try:
            if source_path is None:
                _fail("missing_input_path")
                return
            if not source_path.is_file():
                _fail("input_file_missing", inputPath=str(source_path))
                return

            controller = getattr(window.view_model, "controller", None)
            config_port = getattr(controller, "config_port", None)
            if output_dir_raw:
                if config_port is None:
                    _fail("config_port_unavailable")
                    return
                output_dir = Path(output_dir_raw)
                output_dir.mkdir(parents=True, exist_ok=True)
                if not config_port.set_many(
                    {
                        "output.directory.mode": "custom",
                        "output.directory.custom_path": str(output_dir),
                        "output.directory.create_date_subfolder": False,
                    }
                ):
                    _fail("output_config_persist_failed")
                    return

            from .main_window import _normalize_path

            normalized = _normalize_path(str(source_path))
            state["normalized_path"] = normalized
            window.view_model.add_files([str(source_path)])
            app.processEvents()
            QTimer.singleShot(100, _poll_conversion)
        except Exception as exc:  # pragma: no cover - release smoke diagnostics
            _fail(f"{type(exc).__name__}: {exc}")

    def _poll_conversion() -> None:
        from PySide6.QtCore import QTimer

        try:
            normalized = str(state.get("normalized_path") or "")
            if not normalized:
                _fail("input_not_registered")
                return

            action_vm = getattr(window, "_action_area_vm", None)
            if not state["conversion_started"]:
                if action_vm is None or not getattr(action_vm, "visible", False):
                    if _timed_out():
                        _fail("action_area_not_ready", inputPath=str(source_path))
                        return
                    QTimer.singleShot(100, _poll_conversion)
                    return
                if getattr(action_vm, "file_type", "") != "image":
                    _fail("unexpected_file_type", fileType=getattr(action_vm, "file_type", ""))
                    return
                if not getattr(action_vm, "extract_ocr", False):
                    action_vm.set_file_to_md_option("extract_ocr", True)
                action_vm.request_conversion("md")
                state["conversion_started"] = True
                app.processEvents()

            batch_vm = getattr(window, "_batch_list_vm", None)
            entry = batch_vm.get_file_entry(normalized) if batch_vm is not None else None
            if entry is not None and entry.status in {"completed", "failed", "cancelled"}:
                output_path = Path(entry.output_path) if entry.output_path else None
                sidecar_path = output_path.with_name(f"{output_path.stem}_ocr.md") if output_path else None
                primary_text = output_path.read_text(encoding="utf-8") if output_path and output_path.is_file() else ""
                sidecar_text = (
                    sidecar_path.read_text(encoding="utf-8") if sidecar_path and sidecar_path.is_file() else ""
                )
                primary_contains_expected = expected_text in primary_text
                sidecar_contains_expected = expected_text in sidecar_text
                _finish(
                    {
                        "success": entry.status == "completed"
                        and (primary_contains_expected or sidecar_contains_expected),
                        "status": entry.status,
                        "inputPath": str(source_path),
                        "outputPath": str(output_path) if output_path else None,
                        "sidecarPath": str(sidecar_path) if sidecar_path else None,
                        "primaryOutputExists": bool(output_path and output_path.is_file()),
                        "sidecarOutputExists": bool(sidecar_path and sidecar_path.is_file()),
                        "primaryReferencesSidecar": bool(sidecar_path and sidecar_path.name in primary_text),
                        "primaryContainsExpectedText": primary_contains_expected,
                        "sidecarContainsExpectedText": sidecar_contains_expected,
                        "expectedText": expected_text,
                        "error": entry.error_message,
                    }
                )
                return

            if _timed_out():
                _fail("conversion_timed_out", inputPath=str(source_path))
                return
            QTimer.singleShot(250, _poll_conversion)
        except Exception as exc:  # pragma: no cover - release smoke diagnostics
            _fail(f"{type(exc).__name__}: {exc}")

    from PySide6.QtCore import QTimer

    QTimer.singleShot(250, _start_conversion)


def _schedule_test_ipc_report(app: QApplication, window: MainWindow) -> None:
    """Write a packaged IPC smoke report after a secondary launch is consumed.

    Unlike ``DOCWEN_GUI_TEST_AUTOCLOSE_MS``, this hook must not disable IPC.
    It is used to prove the packaged primary instance receives file-forwarding
    commands from a second process and then exits under test control.
    """
    import json
    import os
    import time
    from pathlib import Path

    report_path_raw = os.environ.get("DOCWEN_GUI_TEST_IPC_REPORT", "").strip()
    if not report_path_raw:
        return

    expected_file_raw = os.environ.get("DOCWEN_GUI_TEST_IPC_EXPECT_FILE", "").strip()
    expected_file = str(Path(expected_file_raw).resolve()) if expected_file_raw else ""
    try:
        timeout_ms = int(os.environ.get("DOCWEN_GUI_TEST_IPC_TIMEOUT_MS", "30000") or "30000")
    except ValueError:
        timeout_ms = 30000
    timeout_ms = max(timeout_ms, 1000)

    report_path = Path(report_path_raw)
    started_at = time.monotonic()
    received_files: list[str] = []
    activation_count = 0
    finished = False

    def _norm(path: str) -> str:
        return str(Path(path).resolve())

    def _write_report(payload: dict[str, object]) -> None:
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    def _snapshot(*, success: bool, status: str, error: str | None = None) -> dict[str, object]:
        files = [_norm(getattr(file_ref, "path", "")) for file_ref in window.view_model.files]
        selected_file = window.view_model.selected_file
        selected_path = _norm(getattr(selected_file, "path", "")) if selected_file is not None else ""
        expected_received = bool(expected_file and expected_file in [_norm(path) for path in received_files])
        expected_in_files = bool(expected_file and expected_file in files)
        return {
            "success": success,
            "status": status,
            "error": error,
            "expectedFile": expected_file,
            "receivedFiles": [_norm(path) for path in received_files],
            "expectedReceived": expected_received,
            "files": files,
            "expectedInFiles": expected_in_files,
            "selectedFile": selected_path,
            "activationCount": activation_count,
            "elapsedMs": int((time.monotonic() - started_at) * 1000),
        }

    def _finish(success: bool, status: str, error: str | None = None) -> None:
        nonlocal finished
        if finished:
            return
        finished = True
        _write_report(_snapshot(success=success, status=status, error=error))
        # Let the current runtime/control handler return its response before
        # starting MainWindow's execution drain; otherwise the GUI thread and
        # control thread can wait on each other during packaged smoke shutdown.
        QTimer.singleShot(0, window.close)

    def _maybe_finish() -> None:
        if not expected_file:
            return
        snapshot = _snapshot(success=False, status="pending")
        if snapshot["expectedReceived"] and snapshot["expectedInFiles"] and activation_count > 0:
            _finish(True, "completed")

    def _on_ipc_file_received(path: str) -> None:
        received_files.append(path)
        _maybe_finish()

    def _on_activation_requested() -> None:
        nonlocal activation_count
        activation_count += 1
        _maybe_finish()

    def _on_timeout() -> None:
        _finish(False, "timed_out", "ipc_smoke_timed_out")

    window.view_model.ipc_file_received.connect(_on_ipc_file_received)
    window.view_model.window_activation_requested.connect(_on_activation_requested)

    from PySide6.QtCore import QTimer

    QTimer.singleShot(0, lambda: _write_report(_snapshot(success=False, status="waiting")))
    QTimer.singleShot(timeout_ms, _on_timeout)


def run_gui(
    controller: ApplicationController | None = None,
    initial_file: str | None = None,
    argv: list[str] | None = None,
) -> int:
    """Convenience: create QApplication, build main window, show, and exec.

    This is a simplified entry point for development and testing.
    It does NOT perform IPC bootstrap — the production entry point with
    single-instance locking, controller wiring, and the task-event bridge
    lives in ``docwen_bundle.gui_entry.main``.

    Returns:
        The QApplication exit code.
    """
    from docwen_runtime.errors import FAILURE_LABELS, FailureCategory, SecurityCheckFailedError

    try:
        _run_gui_security_startup()
    except SecurityCheckFailedError:
        logger.critical(FAILURE_LABELS[FailureCategory.SECURITY_CHECK])
        return 1

    app = create_qapplication(argv)

    _initialize_application_theme(app, controller)

    initial_files = [initial_file] if initial_file else None
    window = create_main_window(controller=controller, initial_files=initial_files)
    _schedule_test_autoclose(app, window)
    _schedule_test_notification_report(app, window)
    _schedule_test_ocr_report(app, window)
    from .release_smoke import _schedule_test_conversion_report

    _schedule_test_conversion_report(app, window)
    from .settings_smoke import _schedule_test_settings_report

    _schedule_test_settings_report(app, window)
    _schedule_test_ipc_report(app, window)
    window.show()
    return app.exec()


__all__ = [
    "create_main_window",
    "create_qapplication",
    "run_gui",
]

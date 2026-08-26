"""Environment-gated GUI conversion hook used by packaged release verification."""

from __future__ import annotations

import json
import os
import time
from pathlib import Path
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from PySide6.QtWidgets import QApplication

    from .main_window import MainWindow


def _allowlisted_conversion_metrics(result: object) -> dict[str, object]:
    """Return the privacy-bounded conversion facts owned by release evidence."""

    metrics = getattr(result, "metrics", None)
    if metrics is None:
        return {}
    extra = getattr(metrics, "extra", {})
    extra_values = extra if isinstance(extra, dict) else {}
    duration_ms = getattr(metrics, "duration_ms", 0.0)
    input_bytes = getattr(metrics, "input_bytes", 0)
    output_bytes = getattr(metrics, "output_bytes", 0)
    engine = extra_values.get("engine")
    backend = extra_values.get("backend")
    return {
        "durationMs": (
            float(duration_ms) if isinstance(duration_ms, (int, float)) and not isinstance(duration_ms, bool) else None
        ),
        "inputBytes": (
            int(input_bytes) if isinstance(input_bytes, int) and not isinstance(input_bytes, bool) else None
        ),
        "outputBytes": (
            int(output_bytes) if isinstance(output_bytes, int) and not isinstance(output_bytes, bool) else None
        ),
        "engine": engine if isinstance(engine, str) and engine.strip() else None,
        "backend": backend if isinstance(backend, str) and backend.strip() else None,
    }


def _schedule_test_conversion_report(app: QApplication, window: MainWindow) -> None:
    """Drive one real MainWindow conversion route and write a JSON report.

    This is an internal, environment-gated release hook. It deliberately uses
    the same ActionArea/ConversionPanel view-model entry points as a user click;
    it does not import plugins or bypass the application/runtime boundary.
    """

    report_raw = os.environ.get("DOCWEN_GUI_TEST_CONVERSION_REPORT", "").strip()
    input_raw = os.environ.get("DOCWEN_GUI_TEST_CONVERSION_INPUT", "").strip()
    if not report_raw and not input_raw:
        return

    report_path = Path(report_raw) if report_raw else None
    source_path = Path(input_raw) if input_raw else None
    output_dir_raw = os.environ.get("DOCWEN_GUI_TEST_CONVERSION_OUTPUT_DIR", "").strip()
    target_format = os.environ.get("DOCWEN_GUI_TEST_CONVERSION_TARGET", "pdf").strip().lower() or "pdf"
    surface = os.environ.get("DOCWEN_GUI_TEST_CONVERSION_SURFACE", "panel").strip().lower() or "panel"
    action_name = os.environ.get("DOCWEN_GUI_TEST_CONVERSION_ACTION", "").strip().lower()
    expected_warning = os.environ.get("DOCWEN_GUI_TEST_CONVERSION_EXPECT_WARNING", "").strip()
    screenshot_raw = os.environ.get("DOCWEN_GUI_TEST_CONVERSION_SCREENSHOT", "").strip()
    screenshot_path = Path(screenshot_raw) if screenshot_raw else None
    warning_row_screenshot_path = (
        screenshot_path.with_name(f"{screenshot_path.stem}_warning_row{screenshot_path.suffix}")
        if screenshot_path is not None
        else None
    )
    try:
        timeout_ms = int(os.environ.get("DOCWEN_GUI_TEST_CONVERSION_TIMEOUT_MS", "90000") or "90000")
    except ValueError:
        timeout_ms = 90000
    timeout_ms = max(timeout_ms, 1000)

    started_at = time.monotonic()
    state: dict[str, object] = {
        "conversion_started": False,
        "conversion_metrics": {},
        "normalized_path": "",
        "terminal_pending": False,
        "terminal_wait_started": 0.0,
    }

    original_execution_finished = getattr(window, "_on_execution_finished", None)
    if callable(original_execution_finished):

        def _capture_execution_finished(result: object, context: dict[str, object]) -> None:
            # The execution thread connects after this release hook runs. Keep
            # the ordinary terminal handler authoritative while observing the
            # exact result delivered to it.
            original_execution_finished(result, context)
            try:
                state["conversion_metrics"] = _allowlisted_conversion_metrics(result)
            except (TypeError, ValueError):
                state["conversion_metrics"] = {}

        window._on_execution_finished = _capture_execution_finished  # type: ignore[method-assign]

    def _write_report(payload: dict[str, object]) -> None:
        if report_path is None:
            return
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    def _finish(payload: dict[str, object]) -> None:
        payload.setdefault("elapsedMs", int((time.monotonic() - started_at) * 1000))
        payload.setdefault("targetFormat", target_format)
        payload.setdefault("surface", surface)
        payload.setdefault("actionName", action_name)
        _write_report(payload)
        window.close()

    def _fail(error: str, **extra: object) -> None:
        _finish({"success": False, "status": "failed", "error": error, **extra})

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
            if surface not in {"action", "panel"}:
                _fail("unsupported_surface", requestedSurface=surface)
                return

            controller = getattr(window.view_model, "controller", None)
            config_port = getattr(controller, "config_port", None)
            if config_port is None:
                _fail("config_port_unavailable")
                return
            if output_dir_raw:
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

    def _request_conversion() -> bool:
        if surface == "action":
            action_vm = getattr(window, "_action_area_vm", None)
            if action_vm is None or not getattr(action_vm, "visible", False):
                return False
            if action_name:
                action_vm.set_file_to_md_option("optimize_for_type", action_name)
                if getattr(action_vm, "action_name", "") != action_name:
                    return False
            action_vm.request_conversion(target_format)
            return True

        panel_vm = getattr(window, "_conversion_panel_vm", None)
        if panel_vm is None or getattr(panel_vm, "current_file_path", "") != str(source_path):
            return False
        panel_vm.request_conversion(target_format)
        return True

    def _finish_terminal(entry: object) -> None:
        try:
            output_raw = str(getattr(entry, "output_path", "") or "")
            output_path = Path(output_raw) if output_raw else None
            output_exists = bool(output_path and output_path.is_file())
            entry_status = str(getattr(entry, "status", "") or "")
            entry_error = str(getattr(entry, "error_message", "") or "")

            history_payload: list[dict[str, object]] = []
            warning_index = -1
            warning_row_tone = ""
            warning_row_tooltip = ""
            warning_row_visible = False
            warning_row_screenshot_saved = warning_row_screenshot_path is None
            warning_row_screenshot_bytes = 0
            warning_row_screenshot_width = 0
            warning_row_screenshot_height = 0
            summary_payload: dict[str, object] = {}
            status_source = ""
            status_tone = ""
            status_meta_text = ""
            status_summary_text = ""
            info_area_visible = False
            screenshot_saved = screenshot_path is None
            screenshot_bytes = 0
            screenshot_width = 0
            screenshot_height = 0

            info_vm = getattr(window, "_info_area_vm", None)
            info_area = getattr(window, "_info_area", None)
            if info_vm is not None:
                history_rows = list(getattr(info_vm, "history_rows", []))
                for index, row in enumerate(history_rows):
                    message = str(getattr(row, "message", "") or "")
                    message_type = str(getattr(row, "message_type", "") or "")
                    history_payload.append(
                        {
                            "message": message,
                            "messageType": message_type,
                            "filePath": str(getattr(row, "file_path", "") or ""),
                        }
                    )
                    if expected_warning and message == expected_warning and message_type == "warning":
                        warning_index = index

                summary = getattr(info_vm, "task_summary", None)
                if summary is not None:
                    summary_payload = {
                        "state": str(getattr(summary, "state", "") or ""),
                        "tone": str(getattr(summary, "tone", "") or ""),
                        "completedCount": int(getattr(summary, "completed_count", 0) or 0),
                        "totalCount": int(getattr(summary, "total_count", 0) or 0),
                        "failedCount": int(getattr(summary, "failed_count", 0) or 0),
                        "navigatePath": str(getattr(summary, "navigate_path", "") or ""),
                    }
                status_source = str(getattr(info_vm, "status_source", "") or "")
                status_tone = str(getattr(info_vm, "status_tone", "") or "")
                status_meta_text = str(getattr(info_vm, "status_meta_text", "") or "")
                status_summary_text = str(getattr(info_vm, "status_summary_text", "") or "")

            if info_area is not None:
                info_area_visible = bool(info_area.isVisible())
                if warning_index >= 0:
                    row_widget = info_area.get_history_row_widget(warning_index)
                    if row_widget is not None:
                        warning_row_tone = str(row_widget.property("infoStatusTone") or "")
                        warning_row_tooltip = row_widget.toolTip()
                        warning_row_visible = bool(row_widget.isVisible())
                        if warning_row_screenshot_path is not None:
                            warning_row_screenshot_path.parent.mkdir(parents=True, exist_ok=True)
                            row_pixmap = row_widget.grab()
                            warning_row_screenshot_width = row_pixmap.width()
                            warning_row_screenshot_height = row_pixmap.height()
                            warning_row_screenshot_saved = bool(
                                row_pixmap.save(str(warning_row_screenshot_path), "PNG")
                            )
                            if warning_row_screenshot_saved and warning_row_screenshot_path.is_file():
                                warning_row_screenshot_bytes = warning_row_screenshot_path.stat().st_size
                if screenshot_path is not None:
                    screenshot_path.parent.mkdir(parents=True, exist_ok=True)
                    pixmap = info_area.grab()
                    screenshot_width = pixmap.width()
                    screenshot_height = pixmap.height()
                    screenshot_saved = bool(pixmap.save(str(screenshot_path), "PNG"))
                    if screenshot_saved and screenshot_path.is_file():
                        screenshot_bytes = screenshot_path.stat().st_size

            summary_matches = (
                summary_payload.get("state") == "success"
                and summary_payload.get("tone") == "warning"
                and summary_payload.get("completedCount") == 1
                and summary_payload.get("failedCount") == 0
                and summary_payload.get("navigatePath") == output_raw
            )
            warning_ui_matches = not expected_warning or (
                warning_index >= 0
                and warning_row_tone == "warning"
                and warning_row_tooltip == expected_warning
                and warning_row_visible
                and status_source == "task"
                and status_tone == "warning"
                and summary_matches
            )
            screenshot_matches = screenshot_saved and (
                screenshot_path is None or (screenshot_bytes > 0 and screenshot_width > 0 and screenshot_height > 0)
            )
            screenshot_matches = (
                screenshot_matches
                and warning_row_screenshot_saved
                and (
                    warning_row_screenshot_path is None
                    or (
                        warning_row_screenshot_bytes > 0
                        and warning_row_screenshot_width > 0
                        and warning_row_screenshot_height > 0
                    )
                )
            )
            success = entry_status == "completed" and output_exists and warning_ui_matches and screenshot_matches
            error = entry_error
            if entry_status == "completed" and output_exists and not warning_ui_matches:
                error = "successful_warning_ui_mismatch"
            elif entry_status == "completed" and output_exists and not screenshot_matches:
                error = "successful_warning_screenshot_failed"

            _finish(
                {
                    "success": success,
                    "status": entry_status,
                    "inputPath": str(source_path),
                    "outputPath": output_raw or None,
                    "outputExists": output_exists,
                    "outputBytes": output_path.stat().st_size if output_exists and output_path else 0,
                    "conversionMetrics": state.get("conversion_metrics", {}),
                    "expectedWarningMessage": expected_warning,
                    "historyRows": history_payload,
                    "warningHistoryIndex": warning_index,
                    "warningRowTone": warning_row_tone,
                    "warningRowTooltip": warning_row_tooltip,
                    "warningRowVisible": warning_row_visible,
                    "warningRowScreenshotPath": (
                        str(warning_row_screenshot_path) if warning_row_screenshot_path else None
                    ),
                    "warningRowScreenshotSaved": warning_row_screenshot_saved,
                    "warningRowScreenshotBytes": warning_row_screenshot_bytes,
                    "warningRowScreenshotWidth": warning_row_screenshot_width,
                    "warningRowScreenshotHeight": warning_row_screenshot_height,
                    "taskSummary": summary_payload,
                    "statusSource": status_source,
                    "statusTone": status_tone,
                    "statusMetaText": status_meta_text,
                    "statusSummaryText": status_summary_text,
                    "infoAreaVisible": info_area_visible,
                    "screenshotPath": str(screenshot_path) if screenshot_path else None,
                    "screenshotSaved": screenshot_saved,
                    "screenshotBytes": screenshot_bytes,
                    "screenshotWidth": screenshot_width,
                    "screenshotHeight": screenshot_height,
                    "error": error,
                }
            )
        except Exception as exc:  # pragma: no cover - release smoke diagnostics
            _fail(f"{type(exc).__name__}: {exc}")

    def _wait_for_persistent_summary(entry: object) -> None:
        from PySide6.QtCore import QTimer

        info_vm = getattr(window, "_info_area_vm", None)
        status_source = str(getattr(info_vm, "status_source", "") or "")
        wait_started_raw = state.get("terminal_wait_started")
        wait_started = wait_started_raw if isinstance(wait_started_raw, (int, float)) else 0.0
        if status_source == "task" or (time.monotonic() - wait_started) >= 10.0:
            _finish_terminal(entry)
            return
        QTimer.singleShot(250, lambda: _wait_for_persistent_summary(entry))

    def _poll_conversion() -> None:
        from PySide6.QtCore import QTimer

        try:
            normalized = str(state.get("normalized_path") or "")
            if not normalized:
                _fail("input_not_registered")
                return

            if not state["conversion_started"]:
                if not _request_conversion():
                    if _timed_out():
                        _fail("conversion_surface_not_ready", inputPath=str(source_path))
                        return
                    QTimer.singleShot(100, _poll_conversion)
                    return
                state["conversion_started"] = True
                app.processEvents()

            batch_vm = getattr(window, "_batch_list_vm", None)
            entry = batch_vm.get_file_entry(normalized) if batch_vm is not None else None
            if entry is not None and entry.status in {"completed", "failed", "cancelled"}:
                if not state["terminal_pending"]:
                    state["terminal_pending"] = True
                    state["terminal_wait_started"] = time.monotonic()
                    if not expected_warning and screenshot_path is None:
                        _finish_terminal(entry)
                        return
                    # Transients intentionally outrank the task summary. Wait
                    # until they naturally expire so the capture proves the
                    # durable warning summary, not only a completion/progress
                    # transient. The bounded wait still emits a fail-closed
                    # report if the persistent projection never surfaces.
                    QTimer.singleShot(250, lambda: _wait_for_persistent_summary(entry))
                return

            if _timed_out():
                _fail("conversion_timed_out", inputPath=str(source_path))
                return
            QTimer.singleShot(250, _poll_conversion)
        except Exception as exc:  # pragma: no cover - release smoke diagnostics
            _fail(f"{type(exc).__name__}: {exc}")

    from PySide6.QtCore import QTimer

    QTimer.singleShot(250, _start_conversion)


__all__ = ["_schedule_test_conversion_report"]

"""Project terminal execution results into shared GUI task and activity state."""

from __future__ import annotations

import re
from pathlib import Path
from typing import TYPE_CHECKING, Any

from PySide6.QtCore import QObject, Signal, Slot

from docwen_gui.diagnostics import DiagnosticSummary
from docwen_gui.i18n import t as _t
from docwen_gui.path_identity import normalize_path
from docwen_gui.view_models.output_files import result_output_paths

if TYPE_CHECKING:
    from docwen_core.models.result import ConversionResult
    from docwen_gui.view_models.action_area_vm import ActionAreaViewModel
    from docwen_gui.view_models.batch_list_vm import BatchListViewModel
    from docwen_gui.view_models.info_area_vm import InfoAreaViewModel
    from docwen_gui.view_models.main_window_vm import MainWindowViewModel
    from docwen_gui.view_models.task_history import TaskHistory


def _result_warning_messages(result: ConversionResult) -> list[str]:
    """Return user-visible warning diagnostics from a successful result."""
    messages: list[str] = []
    for diagnostic in result.diagnostics:
        if diagnostic.level != "warning":
            continue
        if diagnostic.code == "MD2DOCX-NUMBERING-COMMENTS-UNANCHORED":
            messages.append(_t("main_window.numbering_comments_unanchored"))
            continue
        extension_messages = {
            "docwen.conversion.markdown_extension.typed_endnotes.flattened": _t("main_window.extension_loss_endnotes"),
            "docwen.conversion.markdown_extension.extended_headings.flattened": _t(
                "main_window.extension_loss_headings"
            ),
            "docwen.conversion.markdown_extension.captions_references.flattened": _t(
                "main_window.extension_loss_references"
            ),
            "docwen.conversion.markdown_extension.structural_tables.flattened": _t("main_window.extension_loss_tables"),
        }
        if diagnostic.code in extension_messages:
            messages.append(extension_messages[diagnostic.code])
            continue
        if diagnostic.code == "IMAGE-PNG-16BIT-CLIPPED":
            message = _t("main_window.image_png_16bit_clipped")
            if diagnostic.location:
                message = f"{message} ({diagnostic.location})"
            messages.append(message)
            continue
        if diagnostic.code == "OCR-BEST-EFFORT" or diagnostic.code.startswith("OCR-BEST-EFFORT."):
            status = diagnostic.code.partition(".")[2] or "unknown"
            if status == "no_text":
                message = _t(
                    "main_window.ocr_no_text",
                    "OCR detected no text; verify this image against the source.",
                )
            else:
                message = _t(
                    "main_window.ocr_degraded",
                    "OCR did not complete normally (status={status}); usable results were retained.",
                    status=status,
                )
            if diagnostic.location:
                message = f"{message} ({diagnostic.location})"
            messages.append(message)
            continue
        if diagnostic.code == "OCR-TABLE-UNAVAILABLE":
            message = _t("action_area.table_model_unavailable")
            if diagnostic.location:
                message = f"{message} ({diagnostic.location})"
            messages.append(message)
            continue
        if diagnostic.code == "OCR-TABLE-FALLBACK":
            message = _t(
                "main_window.ocr_table_fallback",
                "Table structure recognition failed; plain OCR text was retained.",
            )
            if diagnostic.location:
                message = f"{message} ({diagnostic.location})"
            messages.append(message)
            continue
        message = diagnostic.message.strip() or diagnostic.code.strip()
        if not message:
            message = _t("main_window.conversion_warning", "Conversion completed with a warning")
        if diagnostic.location:
            message = f"{message} ({diagnostic.location})"
        messages.append(message)
    return messages


def _localized_failure_message(error: object | None = None, *, details: bool = False) -> str:
    """Return localized summary copy while keeping raw diagnostics in details."""

    base = (
        _t("main_window.conversion_failed_detail", "Conversion failed.")
        if details
        else _t("main_window.conversion_failed", "Conversion failed. Open failure details for diagnostic information.")
    )
    if error is None:
        return base
    if isinstance(error, str):
        candidate = error.partition(":")[0].strip()
        stable_code = candidate if re.fullmatch(r"[A-Z][A-Z0-9_-]{1,63}", candidate) else ""
        return f"{base} [{stable_code}]" if stable_code else base
    diagnostic_code = str(getattr(error, "diagnostic_code", "") or "").strip()
    error_type = str(getattr(error, "error_type", "") or "").strip()
    stable_code = diagnostic_code or error_type
    if stable_code == "MD2DOCX-NOTE-SYNTAX-INVALID":
        return _t("main_window.note_syntax_invalid")
    if stable_code == "MD2DOCX-NOTE-PART-INVALID":
        return _t("main_window.note_template_incompatible")
    if stable_code == "OUTPUT_ATOMIC_PUBLISH_UNSUPPORTED":
        return _t("main_window.output_filesystem_unsupported")
    if stable_code == "MERGE-FORMULA-CACHE-MISSING":
        return _t("main_window.merge_formula_cache_missing")
    return f"{base} [{stable_code}]" if stable_code else base


class ExecutionPresenter(QObject):
    """Commit result truth before projecting list rows, summaries and navigation.

    This object reads no widgets or persisted settings and owns no worker.
    The window handles output navigation and desktop notification signals.
    """

    open_output = Signal(str)
    completed = Signal(object)

    def __init__(
        self,
        *,
        view_model: MainWindowViewModel,
        batch_list_vm: BatchListViewModel,
        action_area_vm: ActionAreaViewModel,
        info_area_vm: InfoAreaViewModel,
        task_history: TaskHistory,
        parent: QObject,
    ) -> None:
        super().__init__(parent)
        self._view_model = view_model
        self._batch_list_vm = batch_list_vm
        self._action_area_vm = action_area_vm
        self._info_area_vm = info_area_vm
        self._task_history = task_history

    def _show_ocr_notice(self, results: list[ConversionResult], context: dict[str, Any]) -> None:
        if any(d.code == "OCR-QUALITY-NOTICE" for result in results for d in result.diagnostics):
            self._info_area_vm.add_message(
                _t("main_window.ocr_quality_notice"), "info", operation_id=context.get("request_id", "")
            )

    def started(self, context: dict[str, Any]) -> None:
        """Present the frozen request, without rereading widgets or settings."""
        self._info_area_vm.set_execution_summary(str(context.get("request_summary", "")))

    @Slot(object, object)
    def finished(self, result: object, context: dict[str, Any]) -> None:
        from docwen_core.models.result import ConversionResult

        self._task_history.remember(context)
        task_id = context.get("request_id", "")
        file_path = context.get("file_path", "")
        file_paths = list(context.get("file_paths", []) or ([file_path] if file_path else []))
        total_count = int(context.get("total_count", len(file_paths) or 1))
        self._action_area_vm.hide_cancel()

        if context.get("batch"):
            if not isinstance(result, list):
                self.failed("Invalid batch conversion result", context)
                return
            self._batch_finished(result, context)
            return

        if not isinstance(result, ConversionResult):
            self.failed("Invalid conversion result", context)
            return

        if result.success:
            self._show_ocr_notice([result], context)
            self._show_merge_summary(result, context)
            output_path = self._output_path(result)
            result_paths = result_output_paths(result)
            warning_messages = _result_warning_messages(result)
            completion_tone = "warning" if warning_messages else "success"
            merge_details = [item.message for item in result.diagnostics if item.code == "MERGE-SUMMARY"]
            for index, path in enumerate(file_paths):
                self.file_status(
                    path,
                    "completed",
                    output_path=output_path,
                    output_paths=result_paths,
                    operation_id=task_id,
                    error_message="",
                    warnings=tuple(warning_messages),
                    diagnostic=DiagnosticSummary.from_result(result, output_count=len(result_paths)),
                    local_details=(merge_details[index - 1],) if 0 < index <= len(merge_details) else (),
                )
            self._info_area_vm.add_message(
                _t("info_area.history_completed", name=context.get("display_name", Path(file_path).name)),
                completion_tone,
                show_location=bool(output_path),
                file_path=output_path or file_path,
                navigate_file_path=output_path or "",
                operation_id=task_id,
            )
            for warning_message in warning_messages:
                self._info_area_vm.add_message(
                    warning_message,
                    "warning",
                    show_location=bool(output_path),
                    file_path=output_path or file_path,
                    navigate_file_path=output_path or "",
                    operation_id=task_id,
                )
            guide_actions = self._info_area_vm.compute_guide_actions(
                "success",
            )
            self._info_area_vm.set_task_summary(
                operation_id=task_id,
                current_file=context.get("display_name", Path(file_path).name),
                current_file_path=file_path,
                completed_count=total_count,
                total_count=total_count,
                failed_count=0,
                warning_count=len(warning_messages),
                state="success",
                tone=completion_tone,
                navigate_file_path=output_path or "",
                navigation_kind="output",
                output_path=output_path,
                guide_actions=guide_actions,
                output_paths=result_paths,
                batch=bool(context.get("aggregate")),
            )
            self._publish_summary("completed")
            if context.get("open_after_done") and output_path:
                self.open_output.emit(output_path)
            self.completed.emit(context)

        else:
            self._unsuccessful(result, context)
            result_error = result.error
            cancelled = bool(result_error is not None and result_error.error_type == "cancelled")
            self._publish_summary(
                "cancelled" if cancelled else "failed",
                message=(
                    _t("main_window.task_cancelled_status") if cancelled else _localized_failure_message(result_error)
                ),
            )
            self.completed.emit(context)

    def _show_merge_summary(self, result: ConversionResult, context: dict[str, Any]) -> None:
        summaries = result.metrics.extra.get("merge_summaries", [])
        if not isinstance(summaries, list):
            return
        source_paths = context.get("file_paths", [])
        if not context.get("aggregate") or not isinstance(source_paths, (list, tuple)):
            return
        # Metrics are extensible plugin data. Render at most one summary per
        # actual collection input, and never let malformed optional evidence
        # prevent the task's authoritative completion state from being shown.
        summary_limit = max(0, len(source_paths) - 1)
        messages: list[str] = []
        modes = {
            "row": _t("settings.spreadsheet.merge_modes.by_row"),
            "col": _t("settings.spreadsheet.merge_modes.by_column"),
            "cell": _t("settings.spreadsheet.merge_modes.by_cell"),
        }
        for summary in summaries[:summary_limit]:
            if not isinstance(summary, dict):
                continue
            counts = summary.get("counts")
            source_name, sheet, mode = summary.get("source_name"), summary.get("sheet"), summary.get("mode")
            if (
                not isinstance(source_name, str)
                or not 0 < len(source_name) <= 1024
                or not isinstance(sheet, str)
                or not 0 < len(sheet) <= 256
                or not isinstance(mode, str)
                or mode not in modes
                or not isinstance(counts, dict)
                or any(
                    type(counts.get(key)) is not int or not 0 <= counts[key] <= (2**63 - 1)
                    for key in ("added", "replaced", "skipped", "combined")
                )
            ):
                continue
            message = _t(
                "main_window.merge_result_summary",
                source=source_name,
                sheet=sheet,
                mode=modes[mode],
                added=counts["added"],
                replaced=counts["replaced"],
                skipped=counts["skipped"],
                combined=counts["combined"],
            )
            self._info_area_vm.add_message(message, "info", operation_id=context.get("request_id", ""))
            messages.append(message)
        if messages:
            self._info_area_vm.set_execution_summary(" · ".join(messages))

    @Slot(str, object)
    def failed(self, error_message: str, context: dict[str, Any]) -> None:
        # Admission or reservation can fail before started() projects this request.
        self._info_area_vm.set_execution_summary(str(context.get("request_summary", "")))
        self._task_history.remember(context)
        task_id = context.get("request_id", "")
        file_path = context.get("file_path", "")
        file_paths = list(context.get("file_paths", []) or ([file_path] if file_path else []))
        total_count = int(context.get("total_count", len(file_paths) or 1))
        self._action_area_vm.hide_cancel()
        message = error_message or "Conversion failed"
        for path in file_paths:
            self.file_status(
                path,
                "failed",
                error_message=message,
                operation_id=task_id,
                output_path="",
            )
        self._info_area_vm.add_message(
            _localized_failure_message(message),
            "danger",
            show_location=False,
            operation_id=task_id,
            details_message=_localized_failure_message(message, details=True),
        )
        guide_actions = self._info_area_vm.compute_guide_actions(
            "failed",
            failed_details_path=file_path,
            retry_available=True,
        )
        self._info_area_vm.set_task_summary(
            operation_id=task_id,
            current_file=context.get("display_name", Path(file_path).name),
            current_file_path=file_path,
            completed_count=0,
            total_count=total_count,
            failed_count=total_count,
            state="failed",
            tone="danger",
            navigate_file_path=file_path,
            navigation_kind="failed",
            guide_actions=guide_actions,
        )
        self._publish_summary("failed", message=_localized_failure_message(message))
        self.completed.emit(context)

    def _batch_finished(self, results: list[object], context: dict[str, Any]) -> None:
        from docwen_core.models.result import ConversionResult

        self._task_history.remember(context)

        task_id = context.get("request_id", "")
        file_paths = list(context.get("file_paths", []))
        total_count = int(context.get("total_count", len(file_paths) or len(results)))
        success_count = 0
        failed_count = 0
        skipped_count = 0
        cancelled_count = 0
        output_paths: list[str] = []
        retained_failure_paths: list[str] = []
        warning_rows: list[tuple[str, str, str]] = []
        first_failed_path = ""
        first_error_message = ""
        first_error_summary_source: object | None = None
        first_error_output = ""
        first_retained_failure: tuple[str, str, str] | None = None

        self._show_ocr_notice([r for r in results if isinstance(r, ConversionResult)], context)
        for index, file_path in enumerate(file_paths):
            raw_result = results[index] if index < len(results) else None
            if not isinstance(raw_result, ConversionResult):
                failed_count += 1
                message = "Invalid batch conversion result"
                if not first_failed_path:
                    first_failed_path = file_path
                    first_error_message = message
                    first_error_summary_source = "INVALID-BATCH-RESULT"
                self.file_status(
                    file_path,
                    "failed",
                    output_path="",
                    error_message=message,
                    operation_id=task_id,
                )
                continue

            if raw_result.success:
                success_count += 1
                output_path = self._output_path(raw_result)
                result_paths = result_output_paths(raw_result)
                if output_path:
                    output_paths.append(output_path)
                warning_messages = _result_warning_messages(raw_result)
                for warning_message in warning_messages:
                    warning_rows.append((file_path, output_path, warning_message))
                self.file_status(
                    file_path,
                    "completed",
                    output_path=output_path,
                    output_paths=result_paths,
                    error_message="",
                    operation_id=task_id,
                    warnings=tuple(warning_messages),
                    diagnostic=DiagnosticSummary.from_result(raw_result, output_count=len(result_paths)),
                )
                continue

            error = raw_result.error
            error_type = getattr(error, "error_type", "") if error is not None else ""
            message = error.message if error is not None else "Conversion failed"
            if error_type == "cancelled":
                cancelled_count += 1
                self.file_status(
                    file_path,
                    "cancelled",
                    diagnostic=DiagnosticSummary.from_result(raw_result, output_count=0),
                    output_path="",
                    error_message=message,
                    operation_id=task_id,
                )
            elif error_type == "skipped":
                skipped_count += 1
                self.file_status(
                    file_path,
                    "skipped",
                    diagnostic=DiagnosticSummary.from_result(raw_result, output_count=0),
                    output_path="",
                    skip_reason=message,
                    error_message="",
                    operation_id=task_id,
                )
            else:
                failed_count += 1
                retained_output_path = self.existing_output_path(raw_result)
                if retained_output_path:
                    retained_failure_paths.append(retained_output_path)
                self.file_status(
                    file_path,
                    "failed",
                    output_path=retained_output_path,
                    diagnostic=DiagnosticSummary.from_result(
                        raw_result, output_count=len(result_output_paths(raw_result, existing_only=True))
                    ),
                    output_paths=result_output_paths(raw_result, existing_only=True),
                    error_message=message,
                    operation_id=task_id,
                )
                if retained_output_path and first_retained_failure is None:
                    first_retained_failure = (file_path, retained_output_path, message)
                if not first_failed_path:
                    first_failed_path = file_path
                    first_error_message = message
                    first_error_summary_source = error or message
                    first_error_output = retained_output_path

        completed_count = success_count + failed_count
        if cancelled_count and not failed_count:
            state = "cancelled"
            tone = "warning"
        elif failed_count:
            state = "partial" if success_count else "failed"
            tone = "warning" if success_count else "danger"
        elif skipped_count:
            state = "success" if success_count else "skipped"
            tone = "warning"
        else:
            state = "success"
            tone = "warning" if warning_rows else "success"

        successful_output_path = output_paths[0] if output_paths else ""
        retained_failure_output_path = retained_failure_paths[0] if retained_failure_paths else ""
        guide_output_path = successful_output_path or retained_failure_output_path
        navigate_path = output_paths[0] if state == "success" and output_paths else first_failed_path
        navigation_kind = "output" if state == "success" else ("failed" if first_failed_path else "")
        guide_actions = self._info_area_vm.compute_guide_actions(
            state,
            failed_details_path=first_failed_path,
            retry_available=bool(first_failed_path),
        )
        self._info_area_vm.add_message(
            _t(
                "components.info_area.batch_completed",
                "Batch finished: {success} succeeded, {failed} failed, {skipped} skipped, {cancelled} cancelled",
                success=success_count,
                failed=failed_count,
                skipped=skipped_count,
                cancelled=cancelled_count,
            ),
            tone,
            show_location=bool(guide_output_path),
            file_path=guide_output_path,
            navigate_file_path=guide_output_path,
            operation_id=task_id,
        )
        for warning_file, warning_output, warning_message in warning_rows:
            self._info_area_vm.add_message(
                f"{Path(warning_file).name}: {warning_message}",
                "warning",
                show_location=bool(warning_output),
                file_path=warning_output or warning_file,
                navigate_file_path=warning_output or "",
                operation_id=task_id,
            )
        if first_error_message:
            self._info_area_vm.add_message(
                _localized_failure_message(first_error_summary_source or first_error_message),
                "danger",
                show_location=bool(first_error_output),
                file_path=first_error_output,
                navigate_file_path=first_error_output,
                operation_id=task_id,
                details_message=_localized_failure_message(
                    first_error_summary_source or first_error_message, details=True
                ),
            )
        if first_retained_failure is not None and first_retained_failure[0] != first_failed_path:
            _, retained_output, _retained_message = first_retained_failure
            self._info_area_vm.add_message(
                _localized_failure_message(),
                "danger",
                show_location=True,
                file_path=retained_output,
                navigate_file_path=retained_output,
                details_message=_localized_failure_message(details=True),
                operation_id=task_id,
            )
        self._info_area_vm.set_task_summary(
            operation_id=task_id,
            current_file=context.get("display_name", "Batch conversion"),
            current_file_path=context.get("file_path", ""),
            completed_count=completed_count,
            total_count=total_count,
            failed_count=failed_count,
            skipped_count=skipped_count,
            cancelled_count=cancelled_count,
            warning_count=len(warning_rows),
            state=state,
            tone=tone,
            navigate_file_path=navigate_path,
            navigation_kind=navigation_kind,
            output_path=guide_output_path,
            guide_actions=guide_actions,
            batch=True,
        )
        terminal_message = (
            _localized_failure_message(first_error_summary_source or first_error_message)
            if state in {"failed", "partial"}
            else ""
        )
        self._publish_summary("completed" if state in {"success", "skipped"} else state, message=terminal_message)
        if context.get("open_after_done") and successful_output_path:
            self.open_output.emit(successful_output_path)
        self.completed.emit(context)

    def _publish_summary(self, status: str, *, message: str = "") -> None:
        """Expose the already-committed InfoArea summary to GUI observers."""
        summary = self._info_area_vm.task_summary
        self._view_model.publish_execution_summary(
            status,
            {
                "task_id": summary.operation_id,
                "state": summary.state,
                "completed_count": summary.completed_count,
                "total_count": summary.total_count,
                "failed_count": summary.failed_count,
                "skipped_count": summary.skipped_count,
                "cancelled_count": summary.cancelled_count,
                "message": message,
            },
        )

    def file_status(self, path: str, status: str, **values: Any) -> None:
        """Record task truth before projecting it into the editable list."""
        self._task_history.record(
            str(values.get("operation_id", "")),
            normalize_path(path),
            status,
            str(values.get("error_message") or ""),
            str(values.get("output_path") or ""),
            warnings=tuple(values.pop("warnings", ())),
            skip_reason=str(values.get("skip_reason") or ""),
            output_paths=tuple(values.get("output_paths") or ()),
            diagnostic=values.pop("diagnostic", None),
            local_details=tuple(values.pop("local_details", ())),
        )
        self._batch_list_vm.set_file_status(path, status, **values)

    @staticmethod
    def existing_output_path(result: ConversionResult) -> str:
        """Return a real retained artifact path suitable for failure navigation."""
        paths = result_output_paths(result, existing_only=True)
        return paths[0] if paths else ""

    @staticmethod
    def _output_path(result: ConversionResult) -> str:
        paths = result_output_paths(result)
        return paths[0] if paths else ""

    def _unsuccessful(self, result: ConversionResult, context: dict[str, Any]) -> None:
        error = result.error
        message = error.message if error is not None else "Conversion failed"
        cancelled = bool(error is not None and error.error_type == "cancelled")
        state = "cancelled" if cancelled else "failed"
        tone = "warning" if cancelled else "danger"
        task_id = context.get("request_id", "")
        file_path = context.get("file_path", "")
        file_paths = list(context.get("file_paths", []) or ([file_path] if file_path else []))
        total_count = int(context.get("total_count", len(file_paths) or 1))
        retained_output_path = "" if cancelled else self.existing_output_path(result)
        retained_paths = () if cancelled else result_output_paths(result, existing_only=True)

        entry_status = "cancelled" if cancelled else "failed"
        for path in file_paths:
            self.file_status(
                path,
                entry_status,
                diagnostic=DiagnosticSummary.from_result(result, output_count=len(retained_paths)),
                output_path=retained_output_path,
                output_paths=retained_paths,
                error_message=message,
                operation_id=task_id,
            )
        self._info_area_vm.add_message(
            _t("main_window.task_cancelled_status") if cancelled else _localized_failure_message(error),
            tone,
            show_location=bool(retained_output_path),
            file_path=retained_output_path,
            navigate_file_path=retained_output_path,
            operation_id=task_id,
            details_message=_t("main_window.task_cancelled_status")
            if cancelled
            else _localized_failure_message(error, details=True),
        )
        guide_actions = self._info_area_vm.compute_guide_actions(
            state,
            failed_details_path="" if cancelled else file_path,
            retry_available=not cancelled,
        )
        self._info_area_vm.set_task_summary(
            operation_id=task_id,
            current_file=context.get("display_name", Path(file_path).name),
            current_file_path=file_path,
            completed_count=0,
            total_count=total_count,
            failed_count=0 if cancelled else total_count,
            output_path=retained_output_path,
            cancelled_count=total_count if cancelled else 0,
            output_paths=retained_paths,
            batch=bool(context.get("aggregate")),
            state=state,
            tone=tone,
            navigate_file_path="" if cancelled else file_path,
            navigation_kind="" if cancelled else "failed",
            guide_actions=guide_actions,
        )

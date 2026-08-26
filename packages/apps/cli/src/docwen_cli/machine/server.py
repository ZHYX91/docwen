"""JSON-RPC 2.0 server for DocWen Machine Protocol v1 over stdio."""

from __future__ import annotations

import contextlib
import math
import re
import threading
from collections.abc import Callable
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass
from typing import Any, BinaryIO, Literal, cast

from jsonschema import ValidationError

from docwen_application.conversion_service import (
    ConversionPlanRequest,
    ConversionService,
    ConversionServiceError,
    ConversionTaskOutcome,
    LocalInputHandle,
    StagingOutputTarget,
)
from docwen_cli.machine.contracts import MachineContractValidator
from docwen_cli.machine.framing import FrameWriter, FramingError, read_frame
from docwen_cli.machine.query_service import MachineQueryError, MachineQueryService
from docwen_core.models import ConversionDiagnostic, ConversionErrorInfo
from docwen_core.models.task import TaskEvent
from docwen_core.version import PRODUCT_VERSION

_IDENTIFIER = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._:-]{0,127}$")
_METHODS = [
    "initialize",
    "capability/list",
    "health/check",
    "file/inspect",
    "resource/list",
    "gui/status",
    "gui/activate",
    "gui/open",
    "task/plan",
    "task/execute",
    "task/cancel",
]
_PROGRESS_PHASE = "conversion"
_PROGRESS_TOTAL = 100
_MAX_RUNTIME_PROGRESS = 95
_MAX_PENDING_PLANS = 256
_MAX_DIAGNOSTICS = 64
_MAX_RELATED_RANGES = 16
_MAX_DIAGNOSTIC_FIXES = 8
_MAX_FIX_EDITS = 16
_MAX_FIX_REPLACEMENT_CODE_POINTS = 4096


@dataclass(slots=True)
class _TaskProgressState:
    operation: Literal["conversion", "validation"]
    accepted_inputs: set[tuple[str, str]]
    sequence: int = 0
    completed: int = -1


class MachineProtocolServer:
    """One-session, single-task-concurrency Machine Protocol server."""

    def __init__(
        self,
        service: ConversionService,
        reader: BinaryIO,
        writer: BinaryIO,
        *,
        validator: MachineContractValidator | None = None,
        query_service: MachineQueryService | None = None,
        close_callback: Callable[[], None] | None = None,
    ) -> None:
        self._service = service
        self._reader = reader
        self._writer = FrameWriter(writer)
        self._validator = validator or MachineContractValidator()
        self._query_service = query_service
        self._close_callback = close_callback
        self._executor = ThreadPoolExecutor(max_workers=1, thread_name_prefix="docwen-machine")
        self._initialized = False
        self._plan_operations: dict[
            str,
            tuple[Literal["conversion", "validation"], set[tuple[str, str]]],
        ] = {}
        self._task_progress: dict[str, _TaskProgressState] = {}
        self._progress_lock = threading.Lock()

    def run(self) -> int:
        """Process frames until clean EOF or a fatal framing error."""

        exit_code = 0
        try:
            while True:
                try:
                    message = read_frame(self._reader)
                except FramingError as exc:
                    self._write_error(None, -32700, "Parse error", data={"code": exc.code})
                    exit_code = 2
                    break
                if message is None:
                    break
                self._dispatch(message)
        finally:
            self._executor.shutdown(wait=True, cancel_futures=False)
            with self._progress_lock:
                self._task_progress.clear()
            self._plan_operations.clear()
            if self._close_callback is not None:
                with contextlib.suppress(Exception):
                    self._close_callback()
        return exit_code

    def _dispatch(self, message: dict[str, Any]) -> None:
        request_id = self._response_id(message.get("id"))
        method = message.get("method")
        if method not in _METHODS:
            self._write_error(request_id, -32601, "Method not found")
            return
        try:
            self._validator.validate_message(message)
        except ValidationError as exc:
            self._write_error(
                request_id,
                -32602 if message.get("jsonrpc") == "2.0" else -32600,
                "Invalid params" if message.get("jsonrpc") == "2.0" else "Invalid Request",
                data={"validation": self._bounded_validation_message(exc)},
            )
            return

        if method != "initialize" and not self._initialized:
            self._write_error(request_id, -32600, "initialize must be called first")
            return
        if method == "initialize":
            self._handle_initialize(request_id)
        elif method == "capability/list":
            self._write_result(
                request_id,
                {"capabilities": [item.to_dict() for item in self._service.list_capabilities()]},
            )
        elif method in {"health/check", "file/inspect", "resource/list"}:
            self._handle_query(request_id, str(method), message["params"])
        elif method in {"gui/status", "gui/activate", "gui/open"}:
            self._handle_gui(request_id, str(method), message["params"])
        elif method == "task/plan":
            self._handle_plan(request_id, message["params"])
        elif method == "task/execute":
            self._handle_execute(request_id, str(message["params"]["plan_id"]))
        else:
            self._handle_cancel(request_id, str(message["params"]["task_id"]))

    def _handle_query(self, request_id: str | int | None, method: str, params: dict[str, Any]) -> None:
        service = self._query_service
        if service is None:
            self._write_error(request_id, -32603, "Query service unavailable", data={"code": "query_unavailable"})
            return
        try:
            if method == "health/check":
                result = service.health_check()
            elif method == "file/inspect":
                result = service.inspect_file(params["input"])
            else:
                result = service.list_resources(
                    params["kind"],
                    target=params.get("target"),
                    locale=params.get("locale"),
                )
        except MachineQueryError as exc:
            self._write_error(request_id, -32603, "Query failed", data={"code": exc.code, "message": str(exc)})
            return
        self._write_result(request_id, result)

    def _handle_gui(self, request_id: str | int | None, method: str, params: dict[str, Any]) -> None:
        service = self._query_service
        if service is None:
            self._write_error(request_id, -32603, "GUI service unavailable", data={"code": "gui_unavailable"})
            return
        try:
            result = service.gui_control(
                method.removeprefix("gui/"),
                file_path=params.get("file_path"),
                timeout_seconds=int(params["timeout_seconds"]),
            )
        except MachineQueryError as exc:
            self._write_error(request_id, -32603, "GUI request failed", data={"code": exc.code, "message": str(exc)})
            return
        self._write_result(request_id, result)

    def _handle_initialize(self, request_id: str | int | None) -> None:
        if self._initialized:
            self._write_error(request_id, -32600, "initialize may be called only once")
            return
        self._initialized = True
        self._write_result(
            request_id,
            {
                "protocol": {"name": "docwen.machine", "major": 1, "minor": 0},
                "server": {"name": "DocWen", "version": PRODUCT_VERSION},
                "features": {"progress": True, "cancellation": True},
                "methods": list(_METHODS),
                "artifact_bundle_schema": "docwen.artifact_bundle.v2",
                "max_concurrent_tasks": 1,
            },
        )

    def _handle_plan(self, request_id: str | int | None, params: dict[str, Any]) -> None:
        if len(self._plan_operations) >= _MAX_PENDING_PLANS:
            error = ConversionServiceError(
                "resource_exhausted",
                "pending_plan_limit",
                "Too many unconsumed task plans.",
                retryable=True,
            )
            self._write_error(request_id, -32602, "Task planning failed", data={"task_error": error.to_dict()})
            return
        try:
            request = self._plan_request(params)
            plan = self._service.plan(request)
        except ConversionServiceError as exc:
            self._write_error(request_id, -32602, "Task planning failed", data={"task_error": exc.to_dict()})
            return
        payload = plan.to_dict()
        self._plan_operations[str(payload["plan_id"])] = (
            self._progress_operation(str(params["capability_id"])),
            {(str(item["input_id"]), str(item["sha256"])) for item in params["inputs"]},
        )
        self._write_result(request_id, payload)

    def _handle_execute(self, request_id: str | int | None, plan_id: str) -> None:
        with self._progress_lock:
            task_active = bool(self._task_progress)
        if task_active:
            error = ConversionServiceError(
                "resource_exhausted",
                "max_concurrent_tasks",
                "The Machine v1 session already has an active task.",
                retryable=True,
            )
            self._write_error(request_id, -32602, "Task admission failed", data={"task_error": error.to_dict()})
            return
        try:
            task_id = self._service.accept(plan_id)
        except ConversionServiceError as exc:
            self._plan_operations.pop(plan_id, None)
            self._write_error(request_id, -32602, "Task admission failed", data={"task_error": exc.to_dict()})
            return
        operation, accepted_inputs = self._plan_operations.pop(plan_id, ("conversion", set()))
        with self._progress_lock:
            self._task_progress[task_id] = _TaskProgressState(
                operation=operation,
                accepted_inputs=accepted_inputs,
            )
        self._write_result(request_id, {"task_id": task_id, "state": "accepted"})
        self._executor.submit(self._execute_task, task_id)

    def _handle_cancel(self, request_id: str | int | None, task_id: str) -> None:
        try:
            state = self._service.cancel(task_id)
        except Exception:
            self._write_error(request_id, -32603, "Cancellation failed")
            return
        self._write_result(request_id, {"task_id": task_id, "state": state})

    def _execute_task(self, task_id: str) -> None:
        self._write_progress(task_id, 0, status="started")
        try:
            outcome = self._service.execute_accepted(task_id)
        except ConversionServiceError as exc:
            self._write_terminal(
                "task/failed",
                task_id,
                {
                    "error": exc.to_dict(),
                    "diagnostics": [],
                },
            )
            return
        except Exception:
            self._write_terminal(
                "task/failed",
                task_id,
                {
                    "error": {
                        "category": "internal",
                        "code": "internal_error",
                        "message": "An internal DocWen error occurred.",
                        "retryable": False,
                    },
                    "diagnostics": [],
                },
            )
            return
        try:
            self._write_outcome(outcome)
        except Exception:
            self._write_internal_failure(task_id)

    def _write_outcome(self, outcome: ConversionTaskOutcome) -> None:
        if len(outcome.diagnostics) > _MAX_DIAGNOSTICS:
            raise RuntimeError("runtime diagnostic count exceeds Machine v1 bound")
        with self._progress_lock:
            state = self._task_progress.get(outcome.task_id)
            accepted_inputs = set() if state is None else state.accepted_inputs
        diagnostics = [self._diagnostic(item, accepted_inputs=accepted_inputs) for item in outcome.diagnostics]
        if outcome.state == "completed":
            assert outcome.bundle is not None
            method = "task/completed"
            payload = {
                "bundle": outcome.bundle.to_dict(),
                "diagnostics": diagnostics,
                "metrics": {
                    "duration_ms": max(0, round(outcome.metrics.duration_ms)),
                    "input_bytes": max(0, outcome.metrics.input_bytes),
                    "output_bytes": max(0, outcome.metrics.output_bytes),
                },
            }
        elif outcome.state == "cancelled":
            method = "task/cancelled"
            payload = {
                "reason": outcome.error.message if outcome.error is not None else "Task was cancelled",
                "diagnostics": diagnostics,
            }
        else:
            method = "task/failed"
            payload = {
                "error": self._task_error(outcome.error),
                "diagnostics": diagnostics,
            }

        self._validate_terminal_preview(
            method,
            outcome.task_id,
            payload,
            success_progress=outcome.state == "completed",
        )
        if outcome.state == "completed":
            self._write_progress(outcome.task_id, 100, status="complete")
        self._write_terminal(method, outcome.task_id, payload)

    def _validate_terminal_preview(
        self,
        method: str,
        task_id: str,
        payload: dict[str, Any],
        *,
        success_progress: bool,
    ) -> None:
        with self._progress_lock:
            state = self._task_progress.get(task_id)
            if state is None:
                raise RuntimeError("terminal notification has no accepted task state")
            terminal_sequence = state.sequence + (2 if success_progress and state.completed < 100 else 1)
        self._validator.validate_message(
            {
                "jsonrpc": "2.0",
                "method": method,
                "params": {"task_id": task_id, "sequence": terminal_sequence, **payload},
            }
        )

    def _write_internal_failure(self, task_id: str) -> None:
        payload = {
            "error": {
                "category": "internal",
                "code": "internal_error",
                "message": "An internal DocWen error occurred.",
                "retryable": False,
            },
            "diagnostics": [],
        }
        try:
            self._write_terminal("task/failed", task_id, payload)
        except Exception:
            with self._progress_lock:
                self._task_progress.pop(task_id, None)

    def report_runtime_event(self, event: TaskEvent) -> None:
        """Project only bounded, privacy-safe Runtime progress onto Machine v1."""

        if event.event_type != "task_progress":
            return
        percent = event.payload.get("percent")
        if isinstance(percent, bool) or not isinstance(percent, (int, float)):
            return
        value = float(percent)
        if not math.isfinite(value) or not 0 < value <= 100:
            return
        completed = min(_MAX_RUNTIME_PROGRESS, max(1, round(value)))
        self._write_progress(event.task_id, completed, status="progress")

    def _write_progress(self, task_id: str, completed: int, *, status: str) -> None:
        with self._progress_lock:
            state = self._task_progress.get(task_id)
            if state is None or completed <= state.completed:
                return
            state.sequence += 1
            state.completed = completed
            label = "Validation" if state.operation == "validation" else "Conversion"
            if status == "started":
                message = f"{label} started"
            elif status == "complete":
                message = f"{label} complete"
            else:
                message = f"{label} progress {completed} percent"
            self._write_notification(
                "task/progress",
                {
                    "task_id": task_id,
                    "sequence": state.sequence,
                    "phase": _PROGRESS_PHASE,
                    "completed": completed,
                    "total": _PROGRESS_TOTAL,
                    "unit": "percent",
                    "message": message,
                },
            )

    def _write_terminal(self, method: str, task_id: str, payload: dict[str, Any]) -> None:
        with self._progress_lock:
            state = self._task_progress.get(task_id)
            if state is None:
                raise RuntimeError("terminal notification has no accepted task state")
            state.sequence += 1
            try:
                self._write_notification(
                    method,
                    {
                        "task_id": task_id,
                        "sequence": state.sequence,
                        **payload,
                    },
                )
            finally:
                self._task_progress.pop(task_id, None)

    def _write_result(self, request_id: str | int | None, result: dict[str, Any]) -> None:
        self._write({"jsonrpc": "2.0", "id": request_id, "result": result})

    def _write_error(
        self,
        request_id: str | int | None,
        code: int,
        message: str,
        *,
        data: dict[str, Any] | None = None,
    ) -> None:
        error: dict[str, Any] = {"code": code, "message": message}
        if data:
            error["data"] = data
        self._write({"jsonrpc": "2.0", "id": request_id, "error": error})

    def _write_notification(self, method: str, params: dict[str, Any]) -> None:
        self._write({"jsonrpc": "2.0", "method": method, "params": params})

    def _write(self, payload: dict[str, Any]) -> None:
        self._validator.validate_message(payload)
        self._writer.write(payload)

    @staticmethod
    def _plan_request(params: dict[str, Any]) -> ConversionPlanRequest:
        inputs = tuple(
            LocalInputHandle(
                input_id=str(item["input_id"]),
                kind=cast(Literal["document", "resource"], item["kind"]),
                role=cast(
                    Literal[
                        "source",
                        "linked_resource",
                        "bibliography",
                        "citation_style",
                        "neutral_document",
                        "numbering_export_plan",
                    ],
                    item["role"],
                ),
                logical_path=str(item["logical_path"]),
                path=str(item["locator"]["path"]),
                media_type=str(item["media_type"]),
                size_bytes=int(item["size_bytes"]),
                sha256=str(item["sha256"]),
            )
            for item in params["inputs"]
        )
        output = params["output"]
        return ConversionPlanRequest(
            capability_id=str(params["capability_id"]),
            inputs=inputs,
            output=StagingOutputTarget(
                staging_root=str(output["staging_root"]["path"]),
                staging_policy=str(output["staging_policy"]),
            ),
            options=dict(params["options"]),
        )

    @staticmethod
    def _response_id(value: object) -> str | int | None:
        if isinstance(value, str) and value:
            return value[:128]
        if isinstance(value, int) and not isinstance(value, bool):
            return value
        return None

    @staticmethod
    def _bounded_validation_message(exc: ValidationError) -> str:
        message = str(exc.message)
        return message[:512] or "schema validation failed"

    @staticmethod
    def _progress_operation(capability_id: str) -> Literal["conversion", "validation"]:
        return "validation" if capability_id.startswith("validate.") else "conversion"

    @staticmethod
    def _diagnostic(
        item: ConversionDiagnostic,
        *,
        accepted_inputs: set[tuple[str, str]] | None = None,
    ) -> dict[str, Any]:
        severity = item.level if item.level in {"info", "warning", "error"} else "info"
        code = item.code if _IDENTIFIER.fullmatch(item.code) else "docwen.diagnostic"
        message = item.message.strip()[:4096] or "DocWen diagnostic"
        payload: dict[str, Any] = {"severity": severity, "code": code, "message": message}
        if item.artifact_id is not None and _IDENTIFIER.fullmatch(item.artifact_id):
            payload["artifact_id"] = item.artifact_id
        if item.evidence_schema is not None:
            MachineProtocolServer._validate_diagnostic_evidence(item)
            assert item.source is not None
            if accepted_inputs is not None and (item.source.input_id, item.source.sha256) not in accepted_inputs:
                raise RuntimeError("runtime diagnostic source is outside the accepted Machine inputs")
            evidence = item.to_dict()
            for key in ("evidence_schema", "source", "range", "related_ranges", "fixes"):
                payload[key] = evidence[key]
        elif item.code.startswith("docwen.markdown."):
            raise RuntimeError("Markdown semantic diagnostic lacks Machine source evidence")
        return payload

    @staticmethod
    def _validate_diagnostic_evidence(item: ConversionDiagnostic) -> None:
        source = item.source
        source_range = item.range
        if (
            item.evidence_schema != "docwen.machine.diagnostic_evidence.v1"
            or source is None
            or source_range is None
            or item.level != "error"
            or item.artifact_id is not None
            or not item.code.startswith("docwen.markdown.")
            or not _IDENTIFIER.fullmatch(source.input_id)
            or re.fullmatch(r"[0-9a-f]{64}", source.sha256) is None
            or source.encoding != "utf-8"
            or source.coordinate_system != "unicode_code_point"
            or source.offset_base != 0
            or source.range_end != "exclusive"
            or source_range.start < 0
            or source_range.end <= source_range.start
            or len(item.related_ranges) > _MAX_RELATED_RANGES
            or len(item.fixes) > _MAX_DIAGNOSTIC_FIXES
        ):
            raise RuntimeError("runtime diagnostic evidence violates Machine v1 contract")
        if any(value.start < 0 or value.end <= value.start for value in item.related_ranges):
            raise RuntimeError("runtime related diagnostic range is invalid")
        seen_fix_ids: set[str] = set()
        for fix in item.fixes:
            if (
                not _IDENTIFIER.fullmatch(fix.fix_id)
                or fix.fix_id in seen_fix_ids
                or not fix.edits
                or len(fix.edits) > _MAX_FIX_EDITS
            ):
                raise RuntimeError("runtime diagnostic fix violates Machine v1 contract")
            seen_fix_ids.add(fix.fix_id)
            previous_end = -1
            for edit in fix.edits:
                if (
                    edit.range.start < previous_end
                    or edit.range.end < edit.range.start
                    or len(edit.replacement) > _MAX_FIX_REPLACEMENT_CODE_POINTS
                    or (edit.range.start == edit.range.end and not edit.replacement)
                ):
                    raise RuntimeError("runtime diagnostic fix edit violates Machine v1 contract")
                previous_end = edit.range.end

    @staticmethod
    def _task_error(error: ConversionErrorInfo | None) -> dict[str, Any]:
        if error is None:
            return {
                "category": "conversion_failed",
                "code": "conversion_failed",
                "message": "Conversion failed.",
                "retryable": False,
            }
        category = (
            error.error_type
            if error.error_type
            in {
                "invalid_request",
                "unsupported",
                "unavailable",
                "dependency",
                "security",
                "conflict",
                "timeout",
                "resource_exhausted",
                "conversion_failed",
                "internal",
            }
            else "conversion_failed"
        )
        code = error.diagnostic_code if _IDENTIFIER.fullmatch(error.diagnostic_code) else "conversion_failed"
        return {
            "category": category,
            "code": code,
            "message": error.message[:4096] or "Conversion failed.",
            "retryable": error.recoverable,
        }


__all__ = ["MachineProtocolServer"]

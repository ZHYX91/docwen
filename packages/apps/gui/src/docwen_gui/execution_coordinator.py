"""Coordinate GUI route choice, admission, execution projection and progress."""

from __future__ import annotations

import time
from collections.abc import Callable, Sequence
from pathlib import Path
from typing import TYPE_CHECKING, Any, Literal

from PySide6.QtCore import QObject, Slot

from docwen_gui.execution_admission import ExecutionAdmissionError
from docwen_gui.execution_requests import OutputPolicyConfigError
from docwen_gui.i18n import t as _t
from docwen_gui.view_models._runtime_route_filter import (
    RuntimeRouteChoice,
    RuntimeRouteSource,
    discover_composed_action_route_choices,
    discover_runtime_route_choices,
)

if TYPE_CHECKING:
    from docwen_application.controller import ApplicationController
    from docwen_core.models.request import ConversionRequest
    from docwen_gui.execution_presenter import ExecutionPresenter
    from docwen_gui.execution_requests import ExecutionRequestBuilder
    from docwen_gui.qt_bridge.execution_supervisor import ExecutionSupervisor
    from docwen_gui.view_models.action_area_vm import ActionAreaViewModel
    from docwen_gui.view_models.info_area_vm import InfoAreaViewModel
    from docwen_gui.view_models.main_window_vm import MainWindowViewModel
    from docwen_gui.view_models.task_history import TaskHistory


def _route_operation_label(target_format: str, action_name: str) -> str:
    target = str(target_format or "").strip().upper()
    action = str(action_name or "").strip()
    if action:
        return f"{action} → {target}" if target else action
    return _t("conversion_panel.convert") + (f" → {target}" if target else "")


class ExecutionCoordinator(QObject):
    """One route/admit/reserve/start path for single, batch and aggregate work."""

    def __init__(
        self,
        *,
        view_model: MainWindowViewModel,
        action_area_vm: ActionAreaViewModel,
        info_area_vm: InfoAreaViewModel,
        requests: ExecutionRequestBuilder,
        supervisor: ExecutionSupervisor,
        presenter: ExecutionPresenter,
        history: TaskHistory,
        confirm_request: Callable[[ConversionRequest], bool],
        parent: QObject,
    ) -> None:
        super().__init__(parent)
        self._view_model = view_model
        self._action_area_vm = action_area_vm
        self._info_area_vm = info_area_vm
        self._requests = requests
        self._execution = supervisor
        self._results = presenter
        self._task_history = history
        self._confirm_request = confirm_request
        self._context: dict[str, Any] = {}
        self.started_at: float | None = None
        self._accepting = True
        self._preparing = False

    def stop_accepting(self) -> None:
        self._accepting = False

    def clear_context(self) -> None:
        self._context.clear()

    def single(self, *, file_path: str, target_format: str, action_name: str, options: dict[str, Any]) -> None:
        self._start(
            file_paths=[file_path], mode="single", target_format=target_format, action_name=action_name, options=options
        )

    def batch(
        self, *, file_paths: Sequence[str], target_format: str, action_name: str, options: dict[str, Any]
    ) -> None:
        self._start(
            file_paths=file_paths, mode="batch", target_format=target_format, action_name=action_name, options=options
        )

    def aggregate(
        self, *, file_paths: Sequence[str], target_format: str, action_name: str, options: dict[str, Any]
    ) -> None:
        self._start(
            file_paths=file_paths,
            mode="aggregate",
            target_format=target_format,
            action_name=action_name,
            options=options,
        )

    def _start(
        self,
        *,
        file_paths: Sequence[str],
        mode: Literal["single", "batch", "aggregate"],
        target_format: str,
        action_name: str,
        options: dict[str, Any],
    ) -> None:
        if not self._accepting or not file_paths:
            return
        controller = self._view_model.controller
        if controller is None or not controller.has_runtime:
            self._info_area_vm.add_message(
                _t("main_window.runtime_unavailable", "Runtime is unavailable; conversion cannot start."),
                "warning",
            )
            return
        if self._preparing or self._execution.busy:
            self._info_area_vm.add_message(
                _t(
                    "main_window.task_already_running",
                    "A task is already running. Cancel it before starting another one.",
                ),
                "warning",
            )
            return
        self._preparing = True
        try:
            self._prepare_and_start(
                controller=controller,
                file_paths=file_paths,
                mode=mode,
                target_format=target_format,
                action_name=action_name,
                options=options,
            )
        finally:
            self._preparing = False

    def _prepare_and_start(
        self,
        *,
        controller: ApplicationController,
        file_paths: Sequence[str],
        mode: Literal["single", "batch", "aggregate"],
        target_format: str,
        action_name: str,
        options: dict[str, Any],
    ) -> None:
        resolved = self._resolve_route(file_paths=file_paths, target_format=target_format, action_name=action_name)
        if resolved is None:
            return
        target_format, choice = resolved
        try:
            if mode == "single":
                request, context = self._requests.single(
                    file_path=file_paths[0],
                    target_format=target_format,
                    action_name=action_name,
                    options=options,
                    route_options=choice.options,
                )
            else:
                build = self._requests.batch if mode == "batch" else self._requests.aggregate
                request, context = build(
                    file_paths=file_paths,
                    target_format=target_format,
                    action_name=action_name,
                    options=options,
                    route_options=choice.options,
                )
        except OutputPolicyConfigError:
            self._report_output_policy_config_error()
            return
        except ValueError as exc:
            self._info_area_vm.add_message(str(exc), "warning")
            return
        if not self._admit(request, context) or not self._accepting:
            return
        task_id = request.request_id

        def project_reserved_execution() -> None:
            self.started_at = time.monotonic()
            runtime_ids = (
                tuple(f"{task_id}-{index}" for index in range(len(file_paths))) if mode == "batch" else (task_id,)
            )
            self._view_model.begin_execution_telemetry(task_id, runtime_ids)
            for path in file_paths:
                self._results.file_status(path, "processing", operation_id=task_id)
            self._action_area_vm.show_cancel()
            self._info_area_vm.add_message(
                _t("info_area.history_started", name=context["display_name"]),
                "info",
                show_location=False,
                operation_id=task_id,
            )

        self.launch(
            controller=controller,
            request=request,
            context=context,
            project_reserved_execution=project_reserved_execution,
            aggregate_action_name=action_name if mode == "aggregate" else "",
            batch_execution=mode == "batch",
        )

    def _resolve_route(
        self,
        *,
        file_paths: Sequence[str],
        target_format: str,
        action_name: str,
    ) -> tuple[str, RuntimeRouteChoice] | None:
        """Resolve one canonical route for every input before building a request."""

        controller = self._view_model.controller
        sources: list[RuntimeRouteSource] = []
        for file_path in file_paths:
            context = self._requests.file_context(file_path)
            if context is None:
                self._info_area_vm.add_message(
                    _t("main_window.route_unavailable", "No compatible operation is available for this file."),
                    "warning",
                    show_location=True,
                    file_path=file_path,
                    operation=_route_operation_label(target_format, action_name),
                )
                return None
            detected_format, source_category = context
            sources.append(RuntimeRouteSource(detected_format, source_category))
        normalized_target = str(target_format or "").strip().lower()
        if action_name and normalized_target:
            result = discover_composed_action_route_choices(
                controller,
                sources=tuple(sources),
                target=normalized_target,
                action_name=action_name,
            )
        else:
            result = discover_runtime_route_choices(
                controller,
                sources=tuple(sources),
                operation="action" if action_name else "conversion",
                action_name=action_name,
            )
        if result.status == "failed":
            self._info_area_vm.add_message(
                _t(
                    "main_window.route_catalog_failed",
                    "Available operations could not be loaded; the request was not started.",
                ),
                "warning",
            )
            return None
        choice = result.get(normalized_target) if normalized_target else None
        if choice is None and not normalized_target and len(result.choices) == 1:
            choice = result.choices[0]
        if choice is None:
            source_path = file_paths[0] if len(file_paths) == 1 else ""
            self._info_area_vm.add_message(
                _t("main_window.route_unavailable", "No compatible operation is available for this file."),
                "warning",
                show_location=bool(source_path),
                file_path=source_path or None,
                operation=_route_operation_label(target_format, action_name),
            )
            return None
        return choice.target, choice

    def launch(
        self,
        *,
        controller: ApplicationController,
        request: ConversionRequest,
        context: dict[str, Any],
        project_reserved_execution: Callable[[], None],
        aggregate_action_name: str = "",
        batch_execution: bool = False,
    ) -> bool:
        """Project an owned request; the supervisor owns reservation and cleanup."""
        if not self._accepting:
            return False

        def on_reserved() -> None:
            self._task_history.remember(context)
            project_reserved_execution()
            self._context = dict(context)
            self._info_area_vm.begin_task(
                operation_id=request.request_id,
                current_file=context.get("display_name", Path(context.get("file_path", "")).name),
                total_count=int(context.get("total_count", 1)),
            )

        return self._execution.launch(
            controller=controller,
            request=request,
            context=context,
            on_reserved=on_reserved,
            aggregate_action_name=aggregate_action_name,
            batch_execution=batch_execution,
        )

    def _admit(self, request: ConversionRequest, context: dict[str, Any]) -> bool:
        """Project rejected attempts through the same result and history as worker failures."""
        try:
            return self._confirm_request(request)
        except ExecutionAdmissionError as exc:
            self.started_at = time.monotonic()
            self._context = dict(context)
            self._info_area_vm.begin_task(
                operation_id=request.request_id,
                current_file=context.get("display_name", Path(context.get("file_path", "")).name),
                total_count=int(context.get("total_count", 1)),
            )
            self._results.failed(str(exc), context)
            return False

    def _report_output_policy_config_error(self) -> None:
        self._info_area_vm.add_message(
            _t(
                "main_window.output_settings_unavailable",
                "Output settings could not be loaded; the request was not started.",
            ),
            "error",
        )

    @Slot(dict)
    def progress(self, payload: dict[str, Any]) -> None:
        context = self._context
        operation_id = str(payload.get("operation_id", ""))
        if operation_id != context.get("request_id"):
            return
        paths = list(context.get("file_paths", []))
        current_file = ""
        if context.get("batch"):
            task_id = str(payload.get("task_id", ""))
            for index, path in enumerate(paths):
                if task_id == f"{operation_id}-{index}":
                    current_file = Path(path).name
                    break
        percent = payload.get("percent")
        completed = int(payload.get("completed_count", 0))
        if context.get("batch") and paths:
            from docwen_core.events.task_events import TASK_PROGRESS

            fraction = (
                max(0.0, min(100.0, float(percent))) / 100
                if isinstance(percent, (int, float)) and payload.get("event_type") == TASK_PROGRESS
                else 0.0
            )
            percent = 100 * (completed + fraction) / len(paths)
        self._info_area_vm.update_task_progress(
            operation_id,
            current_file=current_file,
            message=str(payload.get("message", "")),
            percent=float(percent) if isinstance(percent, (int, float)) else None,
            completed_count=int(payload.get("completed_count", 0)) if not context.get("aggregate") else 0,
        )

    def cancel(self) -> None:
        active_parent_ids = tuple(self._execution.threads)
        task_id = active_parent_ids[0] if active_parent_ids else self._view_model.current_task_id
        if not task_id:
            return
        try:
            if active_parent_ids:
                self._execution.cancel(task_id)
            else:
                controller = self._view_model.controller
                if controller is None or not controller.has_runtime:
                    return
                controller.cancel(task_id)
            self._info_area_vm.mark_cancelling(task_id)
        except Exception as exc:
            self._info_area_vm.add_message(str(exc), "warning")

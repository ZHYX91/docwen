"""TaskManager — orchestrates conversion tasks through the runtime pipeline.

The TaskManager is the central coordinator that wires together:
  route_resolver → plugin_registry → workspace_manager → plugin → output_finalizer

It handles both single-file and batch conversions.
"""

from __future__ import annotations

import contextlib
import threading
import time
from collections.abc import Callable, Mapping
from dataclasses import dataclass
from typing import TYPE_CHECKING, Any

from docwen_core.cancellation import CancellationToken
from docwen_core.detection import (
    freeze_ooxml_signature_info,
    signature_derived_output_diagnostic,
    signature_info_for_ref,
    signature_validation_diagnostic,
)
from docwen_core.errors import CancellationRequested
from docwen_core.events.task_events import (
    make_task_cancelled,
    make_task_completed,
    make_task_failed,
    make_task_progress,
    make_task_started,
)
from docwen_core.models.file_ref import FileRef
from docwen_core.models.request import PRECONVERSION_INTERMEDIATES_OPTION, ConversionRequest
from docwen_core.models.result import (
    ConversionDiagnostic,
    ConversionErrorInfo,
    ConversionMetrics,
    ConversionResult,
)
from docwen_core.models.task import TaskEvent
from docwen_runtime._request_admission import admit_markdown_ocr_options
from docwen_runtime.config.document_styles import DocumentStyleCatalogError
from docwen_runtime.engine.route_resolver import RouteResolutionError
from docwen_runtime.security import NetworkAccessBlockedError
from docwen_runtime.templates import (
    TemplateNotFoundError,
    TemplateRegistry,
    TemplateResolutionError,
    is_canonical_template_id,
    validate_template_path,
)

if TYPE_CHECKING:
    from docwen_runtime._execution_context import RuntimeExecutionContext
    from docwen_runtime.engine.route_resolver import RouteResolver
    from docwen_runtime.output.finalizer import OutputFinalizer
    from docwen_runtime.plugin_registry.registry import PluginRegistry
    from docwen_runtime.workspace.manager import WorkspaceManager

# Callback signature for event streaming
EventListener = Callable[[TaskEvent], None]


@dataclass(slots=True)
class _SingleTaskState:
    request: ConversionRequest
    input_ref: FileRef
    task_id: str
    token: CancellationToken
    on_event: EventListener | None
    sequence: list[int]
    runtime_context: RuntimeExecutionContext | None = None
    plugin_result: ConversionResult | None = None
    plugin_started_at: float | None = None
    plugin_finished: bool = False
    duration_ms: float = 0.0
    finalization_attempted: bool = False

    def next_sequence(self) -> int:
        value = self.sequence[0]
        self.sequence[0] += 1
        return value


_DOCUMENT_STYLE_TARGETS = frozenset({"docx", "doc", "odt", "rtf", "wps", "pdf"})


class RouteOptionsError(ValueError):
    """Raised when a runtime request contains options outside its route schema."""

    diagnostic_code = "ROUTE_OPTIONS_UNSUPPORTED"

    def __init__(self, option_keys: list[str]) -> None:
        self.option_keys = tuple(option_keys)
        super().__init__(f"Route does not accept option(s): {', '.join(option_keys)}")


def _route_requires_document_styles(route_spec: Any) -> bool:
    """Return whether this exact resolved route renders through DOCX."""

    return (
        route_spec.source_format == "markdown"
        and route_spec.target_format in _DOCUMENT_STYLE_TARGETS
        and not route_spec.action_name
    )


class TaskManager:
    """Orchestrates the full conversion pipeline.

    Does NOT:
    - Manage caller-level thread/process pools.
    - Own GUI/CLI presentation logic.
    - Handle IPC directly.

    Events from plugins (progress, diagnostic, artifact_ready) are
    forwarded through the *on_event* callback so they reach the
    application layer.  The callback is passed into
    ``RuntimeExecutionContext`` → ``_RuntimeProgressSink``, which
    constructs proper ``TaskEvent`` instances via the standard
    factories.
    """

    def __init__(
        self,
        plugin_registry: PluginRegistry,
        route_resolver: RouteResolver,
        workspace_manager: WorkspaceManager,
        output_finalizer: OutputFinalizer,
        *,
        numbering_registry: Any = None,
        proofread_rules: Any = None,
    ) -> None:
        self._plugins = plugin_registry
        self._resolver = route_resolver
        self._workspaces = workspace_manager
        self._finalizer = output_finalizer
        self._numbering_registry = numbering_registry
        self._proofread_rules = proofread_rules

        # Active cancellation tokens keyed by task_id
        self._tokens: dict[str, CancellationToken] = {}
        self._active_task_ids: set[str] = set()
        self._reserved_cancellations: set[str] = set()
        # Pending cancellation requests (for tasks not yet started)
        self._pending_cancellations: set[str] = set()
        self._cancellation_lock = threading.Lock()

    # ── Public API ──────────────────────────────────────────────────

    def execute_single(
        self,
        request: ConversionRequest,
        *,
        on_event: EventListener | None = None,
    ) -> ConversionResult:
        """Execute a single-file conversion synchronously.

        Args:
            request: The conversion request.
            on_event: Optional callback for each ``TaskEvent`` emitted.
                This callback receives events from BOTH the TaskManager
                itself AND the plugin (via ``_RuntimeProgressSink``).

        Returns:
            A ``ConversionResult``.

        Events flow:
            TaskManager._emit() ──→ on_event callback
            Plugin → _RuntimeProgressSink ──→ on_event callback
                                           ──→ Application layer
        """
        # TaskManager is the internal runtime engine. ApplicationController is
        # the public execution boundary and performs file admission before the
        # runtime adapter reaches this method. Runtime-owned option projection
        # remains idempotent so all internal callers share one policy.
        request = admit_markdown_ocr_options(request, request.config_snapshot)
        request = freeze_ooxml_signature_info(request)
        input_ref = next(
            (item for item in request.input_refs if item.input_role in {"source", "neutral_document"}),
            request.input_refs[0],
        )
        task_id = request.request_id

        with self._cancellation_lock:
            if task_id in self._active_task_ids:
                raise RuntimeError(f"Task id is already active: {task_id!r}")
            is_reserved_cancellation = task_id in self._reserved_cancellations
            token = self._tokens.get(task_id)
            if token is None:
                token = CancellationToken()
                self._tokens[task_id] = token
            self._active_task_ids.add(task_id)

            # Apply any pre-existing cancellation request atomically with
            # registration so cancel cannot strand a late pending entry.
            apply_pending_cancellation = task_id in self._pending_cancellations
            if apply_pending_cancellation:
                self._pending_cancellations.discard(task_id)
        if apply_pending_cancellation:
            token.cancel(reason="user_cancelled")

        state = _SingleTaskState(
            request=request,
            input_ref=input_ref,
            task_id=task_id,
            token=token,
            on_event=on_event,
            sequence=[0],
        )
        try:
            self._run_plugin(state, reserved_cancellation=is_reserved_cancellation)
            plugin_result = state.plugin_result
            assert plugin_result is not None

            plugin_reported_cancellation = (
                plugin_result.error is not None and plugin_result.error.error_type == "cancelled"
            )
            if token.is_cancelled or plugin_reported_cancellation:
                return self._plugin_cancelled_result(
                    state,
                    plugin_reported_cancellation=plugin_reported_cancellation,
                )

            if not plugin_result.success or plugin_result.error is not None:
                return self._plugin_failure_result(state)

            return self._plugin_success_result(state)

        except CancellationRequested:
            return self._cancelled_result(state)
        except RouteResolutionError as exc:
            return self._known_failure_result(
                state,
                ConversionErrorInfo(
                    error_type="unsupported_route",
                    message=str(exc),
                    diagnostic_code="ROUTE_UNSUPPORTED",
                ),
                metrics=ConversionMetrics(input_bytes=input_ref.size_bytes),
            )
        except RouteOptionsError as exc:
            return self._known_failure_result(
                state,
                ConversionErrorInfo(
                    error_type="invalid_input",
                    message=str(exc),
                    diagnostic_code=exc.diagnostic_code,
                ),
                metrics=ConversionMetrics(input_bytes=input_ref.size_bytes),
            )
        except TemplateResolutionError as exc:
            return self._known_failure_result(
                state,
                ConversionErrorInfo(
                    error_type=(exc.error_type if isinstance(exc, DocumentStyleCatalogError) else "invalid_input"),
                    message=str(exc),
                    diagnostic_code=exc.diagnostic_code,
                ),
                include_runtime_diagnostics=True,
            )
        except NetworkAccessBlockedError as exc:
            return self._known_failure_result(
                state,
                ConversionErrorInfo(
                    error_type=exc.error_type,
                    message=str(exc),
                    diagnostic_code="NETWORK_ACCESS_BLOCKED",
                ),
                include_runtime_diagnostics=True,
                metrics=ConversionMetrics(
                    duration_ms=(time.perf_counter() - state.plugin_started_at) * 1000.0
                    if state.plugin_started_at is not None
                    else 0.0,
                    input_bytes=input_ref.size_bytes,
                ),
            )
        except Exception as exc:
            return self._unexpected_failure_result(state, exc)
        finally:
            with self._cancellation_lock:
                self._active_task_ids.discard(task_id)
                if task_id not in self._reserved_cancellations:
                    self._tokens.pop(task_id, None)
            # Cleanup workspace after finalization
            with contextlib.suppress(Exception):
                self._workspaces.cleanup(task_id)

    def _run_plugin(
        self,
        state: _SingleTaskState,
        *,
        reserved_cancellation: bool,
    ) -> None:
        if reserved_cancellation:
            state.token.check()

        self._emit(
            state.on_event,
            make_task_started(
                state.task_id,
                state.next_sequence(),
                input_path=state.input_ref.path,
            ),
        )
        plugin_id, route_spec = self._resolver.resolve(
            state.input_ref,
            state.request.target_format,
            state.request.action_name,
        )
        plugin = self._plugins.get(plugin_id)
        if plugin is None:
            raise RuntimeError(f"Plugin {plugin_id!r} not found in registry")

        self._validate_route_options(state.request, route_spec)
        state.request = self._resolve_template_options(state.request)
        workspace = self._workspaces.create(
            state.task_id,
            state.input_ref.path,
            tuple(state.request.input_refs),
        )
        state.runtime_context = self._build_execution_context(
            state,
            route_spec=route_spec,
            workspace=workspace,
        )
        self._emit(
            state.on_event,
            make_task_progress(
                state.task_id,
                state.next_sequence(),
                0.0,
                f"Starting conversion via {plugin_id}",
            ),
        )

        state.plugin_started_at = time.perf_counter()
        state.plugin_result = plugin.convert(state.runtime_context)
        state.duration_ms = (time.perf_counter() - state.plugin_started_at) * 1000.0
        state.plugin_finished = True

    def _build_execution_context(
        self,
        state: _SingleTaskState,
        *,
        route_spec: Any,
        workspace: Any,
    ) -> RuntimeExecutionContext:
        from docwen_core.export_semantics import MarkdownExportSemantics
        from docwen_runtime._execution_context import RuntimeExecutionContext
        from docwen_runtime.config import (
            build_document_style_catalog,
            build_heading_cleanup_rules,
            build_ocr_blockquote_title,
            build_proofread_rules,
        )

        request = state.request
        proofread_rules = self._proofread_rules
        numbering_registry = self._numbering_registry
        heading_cleanup_rules = ()
        if request.config_snapshot:
            proofread_rules = build_proofread_rules(request.config_snapshot)
            if numbering_registry is not None:
                numbering_registry = numbering_registry.with_config_snapshot(
                    request.config_snapshot,
                    locale=request.options.get("locale"),
                )
            heading_cleanup_rules = build_heading_cleanup_rules(request.config_snapshot)

        document_style_catalog = None
        if _route_requires_document_styles(route_spec):
            document_style_catalog = build_document_style_catalog(
                request.config_snapshot,
                request_options=request.options,
            )
        return RuntimeExecutionContext(
            request=request,
            workspace=workspace,
            config_snapshot=request.config_snapshot,
            cancellation_token=state.token,
            on_event=state.on_event,
            shared_seq=state.sequence,
            numbering_registry=numbering_registry,
            heading_cleanup_rules=heading_cleanup_rules,
            proofread_rules=proofread_rules,
            ocr_blockquote_title=build_ocr_blockquote_title(
                request.config_snapshot,
                requested_locale=request.options.get("locale"),
            ),
            document_style_catalog=document_style_catalog,
            markdown_export_semantics=MarkdownExportSemantics.from_config_snapshot(
                request.config_snapshot,
                requested_locale=request.options.get("locale"),
            ),
        )

    def _plugin_cancelled_result(
        self,
        state: _SingleTaskState,
        *,
        plugin_reported_cancellation: bool,
    ) -> ConversionResult:
        assert state.plugin_result is not None
        assert state.runtime_context is not None
        cancellation_error = (
            state.plugin_result.error
            if plugin_reported_cancellation
            else ConversionErrorInfo(
                error_type="cancelled",
                message="Task was cancelled",
            )
        )
        assert cancellation_error is not None
        terminal_diagnostics = self._emit_terminal(
            state.on_event,
            make_task_cancelled(state.task_id, state.next_sequence()),
        )
        return ConversionResult(
            task_id=state.task_id,
            success=False,
            diagnostics=self._merge_diagnostics(
                state.plugin_result.diagnostics,
                state.runtime_context.reported_diagnostics,
                terminal_diagnostics,
            ),
            error=cancellation_error,
            metrics=ConversionMetrics(
                duration_ms=state.duration_ms,
                input_bytes=state.input_ref.size_bytes,
                output_bytes=0,
                extra=self._metrics_extra(state.plugin_result.metrics),
            ),
        )

    def _plugin_failure_result(self, state: _SingleTaskState) -> ConversionResult:
        assert state.plugin_result is not None
        assert state.runtime_context is not None
        preserved = None
        if self._preconversion_intermediate_artifacts(
            state.request,
            state.input_ref.path,
            state.task_id,
        ):
            self._emit(
                state.on_event,
                make_task_progress(
                    state.task_id,
                    state.next_sequence(),
                    90.0,
                    "Finalizing output",
                ),
            )
            state.finalization_attempted = True
            preserved = self._finalize_failure_intermediates(
                state.request,
                state.input_ref.path,
                state.task_id,
                duration_ms=state.duration_ms,
                input_bytes=state.input_ref.size_bytes,
                cancellation=state.token.view(),
            )
        preserved_diagnostics = preserved.diagnostics if preserved is not None else []
        preserved_metrics = preserved.metrics if preserved is not None else ConversionMetrics()
        plugin_error = state.plugin_result.error or ConversionErrorInfo(
            error_type="conversion_failed",
            message="Plugin reported conversion failure",
            diagnostic_code="PLUGIN_REPORTED_FAILURE",
        )
        terminal_diagnostics = self._emit_terminal(
            state.on_event,
            make_task_failed(
                state.task_id,
                state.next_sequence(),
                plugin_error.error_type,
                plugin_error.message,
            ),
        )
        return ConversionResult(
            task_id=state.task_id,
            success=False,
            artifacts=preserved.artifacts if preserved is not None else [],
            diagnostics=self._merge_diagnostics(
                state.plugin_result.diagnostics,
                state.runtime_context.reported_diagnostics,
                preserved_diagnostics,
                terminal_diagnostics,
            ),
            error=plugin_error,
            metrics=ConversionMetrics(
                duration_ms=state.duration_ms,
                input_bytes=state.input_ref.size_bytes,
                output_bytes=preserved_metrics.output_bytes,
                extra={
                    **self._metrics_extra(state.plugin_result.metrics),
                    **self._metrics_extra(preserved_metrics),
                },
            ),
        )

    def _plugin_success_result(self, state: _SingleTaskState) -> ConversionResult:
        assert state.plugin_result is not None
        assert state.runtime_context is not None
        artifacts = list(state.plugin_result.artifacts)
        artifacts.extend(
            self._preconversion_intermediate_artifacts(
                state.request,
                state.input_ref.path,
                state.task_id,
            )
        )
        intentional_empty_success = not artifacts and self._is_intentional_no_output_success(
            state.request,
            state.plugin_result,
            state.runtime_context.reported_diagnostics,
        )
        if intentional_empty_success or not state.request.output_policy.write_artifacts:
            return self._unpublished_success_result(state)
        return self._finalized_success_result(state, artifacts)

    def _unpublished_success_result(self, state: _SingleTaskState) -> ConversionResult:
        assert state.plugin_result is not None
        assert state.runtime_context is not None
        terminal_diagnostics = self._emit_terminal(
            state.on_event,
            make_task_completed(state.task_id, state.next_sequence()),
        )
        return ConversionResult(
            task_id=state.task_id,
            success=True,
            artifacts=[],
            diagnostics=self._merge_diagnostics(
                state.plugin_result.diagnostics,
                state.runtime_context.reported_diagnostics,
                self._ooxml_signature_diagnostics(
                    state.request,
                    delivered_artifact=False,
                ),
                terminal_diagnostics,
            ),
            metrics=ConversionMetrics(
                duration_ms=state.duration_ms,
                input_bytes=state.input_ref.size_bytes,
                output_bytes=0,
                extra=self._metrics_extra(state.plugin_result.metrics),
            ),
        )

    def _finalized_success_result(
        self,
        state: _SingleTaskState,
        artifacts: list[Any],
    ) -> ConversionResult:
        assert state.plugin_result is not None
        assert state.runtime_context is not None
        self._emit(
            state.on_event,
            make_task_progress(
                state.task_id,
                state.next_sequence(),
                90.0,
                "Finalizing output",
            ),
        )
        state.finalization_attempted = True
        finalizer_result = self._finalizer.finalize(
            task_id=state.task_id,
            artifacts=artifacts,
            policy=state.request.output_policy,
            input_path=state.input_ref.path,
            duration_ms=state.duration_ms,
            input_bytes=state.input_ref.size_bytes,
            cancellation=state.token.view(),
        )
        final_error = finalizer_result.error
        if not finalizer_result.success and final_error is None:
            final_error = ConversionErrorInfo(
                error_type="output_finalization_failed",
                message="Output finalization failed",
                diagnostic_code="FINALIZER_FAILED",
            )
        result = ConversionResult(
            task_id=state.task_id,
            success=finalizer_result.success,
            artifacts=finalizer_result.artifacts,
            diagnostics=self._merge_diagnostics(
                state.plugin_result.diagnostics,
                state.runtime_context.reported_diagnostics,
                finalizer_result.diagnostics,
            ),
            error=final_error or state.plugin_result.error,
            metrics=ConversionMetrics(
                duration_ms=finalizer_result.metrics.duration_ms,
                input_bytes=finalizer_result.metrics.input_bytes,
                output_bytes=finalizer_result.metrics.output_bytes,
                extra={
                    **self._metrics_extra(state.plugin_result.metrics),
                    **self._metrics_extra(finalizer_result.metrics),
                },
            ),
        )
        if result.success:
            terminal_event = make_task_completed(state.task_id, state.next_sequence())
        else:
            assert result.error is not None
            terminal_event = make_task_failed(
                state.task_id,
                state.next_sequence(),
                result.error.error_type,
                result.error.message,
            )
        terminal_diagnostics = self._emit_terminal(state.on_event, terminal_event)
        return ConversionResult(
            task_id=result.task_id,
            success=result.success,
            artifacts=result.artifacts,
            diagnostics=self._merge_diagnostics(
                result.diagnostics,
                self._ooxml_signature_diagnostics(
                    state.request,
                    delivered_artifact=bool(result.artifacts),
                )
                if result.success
                else [],
                terminal_diagnostics,
            ),
            error=result.error,
            metrics=result.metrics,
        )

    def _cancelled_result(self, state: _SingleTaskState) -> ConversionResult:
        terminal_diagnostics = self._emit_terminal(
            state.on_event,
            make_task_cancelled(state.task_id, state.next_sequence()),
        )
        runtime_diagnostics = state.runtime_context.reported_diagnostics if state.runtime_context is not None else []
        return ConversionResult(
            task_id=state.task_id,
            success=False,
            diagnostics=self._merge_diagnostics(runtime_diagnostics, terminal_diagnostics),
            error=ConversionErrorInfo(
                error_type="cancelled",
                message="Task was cancelled",
            ),
        )

    def _known_failure_result(
        self,
        state: _SingleTaskState,
        error: ConversionErrorInfo,
        *,
        include_runtime_diagnostics: bool = False,
        metrics: ConversionMetrics | None = None,
    ) -> ConversionResult:
        terminal_diagnostics = self._emit_terminal(
            state.on_event,
            make_task_failed(
                state.task_id,
                state.next_sequence(),
                error.error_type,
                error.message,
            ),
        )
        runtime_diagnostics = (
            state.runtime_context.reported_diagnostics
            if include_runtime_diagnostics and state.runtime_context is not None
            else []
        )
        return ConversionResult(
            task_id=state.task_id,
            success=False,
            diagnostics=self._merge_diagnostics(
                [
                    ConversionDiagnostic(
                        level="error",
                        message=error.message,
                        code=error.diagnostic_code,
                    )
                ],
                runtime_diagnostics,
                terminal_diagnostics,
            ),
            error=error,
            metrics=metrics or ConversionMetrics(),
        )

    def _unexpected_failure_result(
        self,
        state: _SingleTaskState,
        exc: Exception,
    ) -> ConversionResult:
        if state.token.is_cancelled:
            return self._cancelled_result(state)
        if state.plugin_started_at is not None and not state.plugin_finished:
            state.duration_ms = (time.perf_counter() - state.plugin_started_at) * 1000.0

        preserved = None
        if not state.finalization_attempted:
            state.finalization_attempted = True
            preserved = self._finalize_failure_intermediates(
                state.request,
                state.input_ref.path,
                state.task_id,
                duration_ms=state.duration_ms,
                input_bytes=state.input_ref.size_bytes,
                cancellation=state.token.view(),
            )
        preserved_diagnostics = preserved.diagnostics if preserved is not None else []
        preserved_metrics = preserved.metrics if preserved is not None else ConversionMetrics()
        plugin_diagnostics = state.plugin_result.diagnostics if state.plugin_result is not None else []
        plugin_metrics = state.plugin_result.metrics if state.plugin_result is not None else ConversionMetrics()
        runtime_diagnostics = state.runtime_context.reported_diagnostics if state.runtime_context is not None else []
        runtime_error = ConversionErrorInfo(
            error_type="conversion_failed",
            message=str(exc),
        )
        terminal_diagnostics = self._emit_terminal(
            state.on_event,
            make_task_failed(
                state.task_id,
                state.next_sequence(),
                runtime_error.error_type,
                runtime_error.message,
            ),
        )
        return ConversionResult(
            task_id=state.task_id,
            success=False,
            artifacts=preserved.artifacts if preserved is not None else [],
            diagnostics=self._merge_diagnostics(
                plugin_diagnostics,
                runtime_diagnostics,
                preserved_diagnostics,
                terminal_diagnostics,
            ),
            error=runtime_error,
            metrics=ConversionMetrics(
                duration_ms=state.duration_ms,
                input_bytes=state.input_ref.size_bytes,
                output_bytes=preserved_metrics.output_bytes,
                extra={
                    **self._metrics_extra(plugin_metrics),
                    **self._metrics_extra(preserved_metrics),
                },
            ),
        )

    def execute_batch(
        self,
        request: ConversionRequest,
        *,
        on_event: EventListener | None = None,
        continue_on_error: bool = True,
    ) -> list[ConversionResult]:
        """Execute a batch conversion for all input files.

        Each input file is converted independently via ``execute_single()``.
        Results are returned in the same order as ``request.input_refs``.

        Args:
            request: The conversion request with multiple ``input_refs``.
            on_event: Optional callback forwarded to each ``execute_single``.
            continue_on_error: If ``True`` (default), a single file failure
                does not stop the batch.  If ``False``, the batch stops
                at the first error and marks remaining items as skipped.

        Returns:
            A list of ``ConversionResult``, one per input file.
        """
        if any(input_ref.input_role != "source" for input_ref in request.input_refs):
            raise ValueError("batch conversion accepts only independent source inputs")

        results: list[ConversionResult] = []

        for i, input_ref in enumerate(request.input_refs):
            single_request = ConversionRequest(
                request_id=f"{request.request_id}-{i}",
                input_refs=[input_ref],
                target_format=request.target_format,
                action_name=request.action_name,
                options=dict(request.options),
                output_policy=request.output_policy,
                config_snapshot=dict(request.config_snapshot),
                manifest_context=request.manifest_context,
            )

            result = self.execute_single(single_request, on_event=on_event)
            results.append(result)

            # Stop on first failure if continue_on_error is False
            if not result.success and not continue_on_error:
                for _remaining in request.input_refs[i + 1 :]:
                    results.append(
                        ConversionResult(
                            task_id=f"{request.request_id}-skipped",
                            success=False,
                            error=ConversionErrorInfo(
                                error_type="skipped",
                                message="Skipped due to previous error",
                            ),
                        )
                    )
                break

        return results

    def cancel(self, task_id: str) -> bool:
        """Request cancellation of a running task.

        Args:
            task_id: The task to cancel.

        Returns:
            ``True`` if the task was found and cancellation was requested,
            ``False`` if the task was not found.
        """
        with self._cancellation_lock:
            token = self._tokens.get(task_id)
            if token is None:
                # Task not yet started — record pending cancellation
                self._pending_cancellations.add(task_id)
                return True
        token.cancel(reason="user_cancelled")
        return True

    def reserve_cancellation(self, task_id: str) -> None:
        """Reserve a token across an external synchronous execute handoff."""
        with self._cancellation_lock:
            if task_id in self._reserved_cancellations or task_id in self._active_task_ids:
                raise RuntimeError(f"Cancellation task id is already reserved: {task_id!r}")
            token = CancellationToken()
            self._tokens[task_id] = token
            self._reserved_cancellations.add(task_id)
            apply_pending_cancellation = task_id in self._pending_cancellations
            if apply_pending_cancellation:
                self._pending_cancellations.discard(task_id)
        if apply_pending_cancellation:
            token.cancel(reason="user_cancelled")

    def release_cancellation(self, task_id: str) -> None:
        """Release an external reservation once its admission window closes."""
        with self._cancellation_lock:
            self._reserved_cancellations.discard(task_id)
            if task_id not in self._active_task_ids:
                self._tokens.pop(task_id, None)

    def cancel_all(self) -> int:
        """Cancel all running tasks.

        Returns:
            Number of tasks cancelled.
        """
        count = 0
        with self._cancellation_lock:
            tokens = list(self._tokens.values())
        for token in tokens:
            if not token.is_cancelled:
                token.cancel(reason="user_cancelled")
                count += 1
        return count

    @property
    def active_tasks(self) -> list[str]:
        """Return ids of currently active tasks."""
        with self._cancellation_lock:
            return list(self._active_task_ids)

    # ── Helpers ─────────────────────────────────────────────────────

    @staticmethod
    def _emit(on_event: EventListener | None, event: TaskEvent) -> None:
        """Emit a task event to the listener (if any)."""
        if on_event is not None:
            on_event(event)

    @staticmethod
    def _emit_terminal(
        on_event: EventListener | None,
        event: TaskEvent,
    ) -> list[ConversionDiagnostic]:
        """Emit a terminal event without letting presentation replace task truth."""
        try:
            TaskManager._emit(on_event, event)
        except Exception as exc:
            return [
                ConversionDiagnostic(
                    level="warning",
                    message=f"Task event listener rejected {event.event_type}: {exc}",
                    code="TASK_EVENT_LISTENER_ERROR",
                )
            ]
        return []

    @staticmethod
    def _metrics_extra(metrics: ConversionMetrics) -> dict[str, Any]:
        """Return a defensive copy of plugin/finalizer metrics extras."""
        extra: Any = metrics.extra
        return dict(extra) if isinstance(extra, Mapping) else {}

    @staticmethod
    def _is_intentional_no_output_success(
        request: ConversionRequest,
        plugin_result: ConversionResult,
        reported_diagnostics: list[ConversionDiagnostic],
    ) -> bool:
        """Recognize the explicit proofread-disabled diagnostics-only result."""
        if request.action_name != "validate":
            return False
        return any(
            diagnostic.code == "PROOFREAD-SKIPPED" for diagnostic in (*plugin_result.diagnostics, *reported_diagnostics)
        )

    @staticmethod
    def _merge_diagnostics(
        *groups: list[ConversionDiagnostic],
    ) -> list[ConversionDiagnostic]:
        """Merge diagnostics in producer order while removing exact repeats."""
        merged: list[ConversionDiagnostic] = []
        seen: set[tuple[str, str, str, str, str | None]] = set()
        for group in groups:
            for diagnostic in group:
                key = (
                    diagnostic.level,
                    diagnostic.message,
                    diagnostic.code,
                    diagnostic.location,
                    diagnostic.artifact_id,
                )
                if key in seen:
                    continue
                seen.add(key)
                merged.append(diagnostic)
        return merged

    @staticmethod
    def _ooxml_signature_diagnostics(
        request: ConversionRequest,
        *,
        delivered_artifact: bool,
    ) -> list[ConversionDiagnostic]:
        """Project frozen presence-only signature facts onto a success."""
        diagnostics: list[ConversionDiagnostic] = []
        for source_ref in (item for item in request.input_refs if item.input_role == "source"):
            info = signature_info_for_ref(source_ref)
            validation = signature_validation_diagnostic(info)
            if validation is not None:
                diagnostics.append(validation)
            if delivered_artifact:
                derived = signature_derived_output_diagnostic(info)
                if derived is not None:
                    diagnostics.append(derived)
        return TaskManager._merge_diagnostics(diagnostics)

    @staticmethod
    def _validate_route_options(request: ConversionRequest, route_spec: Any) -> None:
        if route_spec.options_schema.get("additionalProperties") is not False:
            return
        properties = route_spec.options_schema.get("properties", {})
        allowed = set(properties) if isinstance(properties, Mapping) else set()
        allowed.add(PRECONVERSION_INTERMEDIATES_OPTION)
        unsupported = sorted(set(request.options) - allowed)
        if unsupported:
            raise RouteOptionsError(unsupported)

    @staticmethod
    def _resolve_template_options(request: ConversionRequest) -> ConversionRequest:
        """Resolve one exact canonical template ID to its validated path.

        Markdown converters intentionally do not import runtime registries.
        The runtime owns bundled resource discovery, so it normalizes
        ``template_name`` at the boundary and lets plugins consume a plain file
        path through the internal option contract. Consumers must pass the
        canonical ID published by ``resources list templates`` exactly.
        """
        raw_template = request.options.get("template_name")
        if raw_template is None:
            return request
        if not isinstance(raw_template, str) or not is_canonical_template_id(raw_template):
            raise TemplateResolutionError(
                "template_name must be an exact canonical template resource ID",
                diagnostic_code="TEMPLATE_ID_INVALID",
            )

        document_template_targets = {"docx", "doc", "odt", "rtf", "wps", "pdf"}
        # CSV export can intentionally use an XLSX template as an intermediate
        # workbook and emit one CSV artifact per populated worksheet.
        spreadsheet_template_targets = {"xlsx", "xls", "ods", "csv"}
        if request.target_format in document_template_targets:
            expected_template_target = "docx"
        elif request.target_format in spreadsheet_template_targets:
            expected_template_target = "xlsx"
        else:
            raise TemplateResolutionError(
                f"Templates are not supported for target format: {request.target_format}",
                diagnostic_code="TEMPLATE_TARGET_UNSUPPORTED",
            )

        try:
            template = TemplateRegistry.default().get_template(raw_template, target_type=expected_template_target)
        except TemplateNotFoundError as exc:
            raise TemplateResolutionError(
                str(exc),
                diagnostic_code="TEMPLATE_NOT_FOUND",
            ) from exc
        resolved_path = str(validate_template_path(template.path, expected_target=expected_template_target))

        return ConversionRequest(
            request_id=request.request_id,
            input_refs=list(request.input_refs),
            target_format=request.target_format,
            action_name=request.action_name,
            options={**request.options, "template_name": resolved_path},
            output_policy=request.output_policy,
            config_snapshot=dict(request.config_snapshot),
        )

    def _finalize_failure_intermediates(
        self,
        request: ConversionRequest,
        input_path: str,
        task_id: str,
        *,
        duration_ms: float = 0.0,
        input_bytes: int = 0,
        cancellation: Any = None,
    ) -> ConversionResult | None:
        """Place requested pre-conversion auxiliaries without masking failure.

        Plugin-owned artifacts remain excluded on failure.  This helper only
        handles the application-recorded hub intermediate and converts an
        output-directory/finalizer exception into an additional diagnostic so
        the original plugin or runtime error remains authoritative.
        """
        try:
            artifacts = self._preconversion_intermediate_artifacts(request, input_path, task_id)
            if not artifacts:
                return None
            return self._finalizer.finalize(
                task_id=task_id,
                artifacts=artifacts,
                policy=request.output_policy,
                input_path=input_path,
                duration_ms=duration_ms,
                input_bytes=input_bytes,
                cancellation=cancellation,
            )
        except CancellationRequested:
            raise
        except Exception as exc:
            return ConversionResult(
                task_id=task_id,
                success=False,
                diagnostics=[
                    ConversionDiagnostic(
                        level="error",
                        message=f"Failed to preserve pre-conversion intermediate: {exc}",
                        code="PRECONVERSION_INTERMEDIATE_FINALIZE_ERROR",
                    )
                ],
                metrics=ConversionMetrics(
                    duration_ms=duration_ms,
                    input_bytes=input_bytes,
                ),
            )

    @staticmethod
    def _preconversion_intermediate_artifacts(
        request: ConversionRequest,
        input_path: str,
        task_id: str,
    ) -> list[Any]:
        """Return auxiliary artifacts requested by application pre-conversion."""
        from docwen_core.models.artifact import ARTIFACT_KIND_AUXILIARY, ArtifactManifest
        from docwen_core.models.request import PRECONVERSION_INTERMEDIATES_OPTION

        records = request.options.get(PRECONVERSION_INTERMEDIATES_OPTION, [])
        if not isinstance(records, list):
            return []

        artifacts: list[ArtifactManifest] = []
        for index, record in enumerate(records):
            if not isinstance(record, dict):
                continue
            applies_to = str(record.get("applies_to_input_path", "") or "")
            if applies_to and applies_to != input_path:
                continue
            staging_path = str(record.get("staging_path", "") or "")
            suggested_name = str(record.get("suggested_name", "") or "")
            if not staging_path or not suggested_name:
                continue
            artifacts.append(
                ArtifactManifest(
                    artifact_id=f"{task_id}-preconversion-{index}",
                    kind=ARTIFACT_KIND_AUXILIARY,
                    staging_path=staging_path,
                    suggested_name=suggested_name,
                    metadata={
                        "source": "preconversion",
                        "source_format": record.get("source_format", ""),
                        "target_format": record.get("target_format", ""),
                        "backend": record.get("backend", ""),
                    },
                    is_primary=False,
                )
            )
        return artifacts

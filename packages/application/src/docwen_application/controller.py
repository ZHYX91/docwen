"""Application controller — the main entry point for GUI/CLI to interact with the app.

The controller translates user intents (from GUI or CLI) into application commands,
selects workflows, publishes application events, and delegates to the runtime port.

It does NOT:
- Load plugins directly
- Create runtime workspaces or physically place final output artifacts
- Schedule workers
- Update GUI widgets or format CLI output
"""

from __future__ import annotations

import tempfile
import threading
from dataclasses import dataclass, field, replace
from pathlib import Path
from typing import TYPE_CHECKING, Any, cast

from docwen_core.cancellation import CancellationToken
from docwen_core.errors import DocWenError

if TYPE_CHECKING:
    from docwen_application.ports.runtime import ConfigPort, PresenterPort, RuntimePort
    from docwen_core.protocols import CancellationTokenView


class ControllerError(DocWenError):
    """Raised when the controller encounters an error."""


class CapabilityUnavailableError(ControllerError):
    """Raised when the injected runtime cannot reflect its capabilities."""

    code = "capability_unavailable"


@dataclass(frozen=True)
class _PreconversionBatchPlan:
    """Internal batch plan after preconversion.

    ``result_slots`` is aligned to the original input order.  ``None`` means
    the corresponding result must be filled from runtime execution; a
    ``ConversionResult`` means preconversion already failed for that input.
    ``input_indices`` keeps each runnable ref tied to that original order.
    """

    request: Any
    result_slots: list[Any | None]
    output_policies: list[Any]
    input_indices: list[int]


@dataclass(frozen=True)
class _ManagedPreconversion:
    """Prepared payload plus the temporary-directory owner keeping it alive."""

    payload: Any
    temp_owner: Any
    manifest_request: Any | None = None

    def cleanup(self) -> None:
        """Release all application-owned pre-conversion staging."""
        self.temp_owner.cleanup()


@dataclass(frozen=True)
class _PreconversionTerminal:
    """Terminal results produced before Runtime admission."""

    results: list[Any]


@dataclass
class _ExecutionCancellationScope:
    """One operation-owned cancellation source spanning Application/Runtime."""

    canonical_task_id: str
    aliases: tuple[str, ...]
    runtime_task_ids: tuple[str, ...]
    token: CancellationToken = field(default_factory=CancellationToken)
    retain_until_release: bool = False
    phase: str = "preconversion"
    claimed: bool = False
    active_runtime_task_id: str | None = None
    runtime_cancel_inflight_ids: set[str] = field(default_factory=set)
    runtime_cancelled_ids: set[str] = field(default_factory=set)
    runtime_release_pending_ids: set[str] = field(default_factory=set)


class ApplicationController:
    """Top-level application controller.

    All external dependencies (runtime, config, presenter) are injected
    at construction time. This ensures:
    - The controller never directly imports runtime/plugin internals.
    - Testability: fake ports can be injected in tests.
    - The dependency direction is explicit and auditable.
    """

    def __init__(
        self,
        runtime_port: RuntimePort | None = None,
        config_port: ConfigPort | None = None,
        presenter_port: PresenterPort | None = None,
    ) -> None:
        """Initialize the application controller.

        Args:
            runtime_port: Delegates execution to the runtime layer.
                If None, the controller runs in a limited mode (no
                conversions can be executed until a port is set).
            config_port: Provides typed configuration values.
            presenter_port: Sends results/errors for display.
        """
        self._runtime_port = runtime_port
        self._config_port = config_port
        self._presenter_port = presenter_port
        self._started = False
        self._cancellation_lock = threading.Lock()
        self._cancellation_scopes: dict[str, _ExecutionCancellationScope] = {}

    def start(self) -> None:
        """Start the application controller."""
        self._started = True

    def stop(self) -> None:
        """Stop the controller and release its injected runtime."""
        self._started = False
        runtime_port = self._runtime_port
        self._runtime_port = None
        if runtime_port is not None:
            runtime_port.shutdown()

    @property
    def is_running(self) -> bool:
        """Whether the controller is currently running."""
        return self._started

    @property
    def has_runtime(self) -> bool:
        """Whether a runtime port has been injected."""
        return self._runtime_port is not None

    def describe_runtime_capabilities(self) -> dict[str, Any]:
        """Return runtime composition facts without exposing the runtime port."""

        from docwen_application.ports.runtime import CapabilityDiscoveryPort

        runtime_port = self._runtime_port
        if runtime_port is None:
            raise CapabilityUnavailableError("Runtime capability discovery is unavailable: no runtime is configured.")
        if not isinstance(runtime_port, CapabilityDiscoveryPort):
            raise CapabilityUnavailableError("Runtime capability discovery is unavailable in this assembly.")
        return runtime_port.describe_capabilities()

    # ── Operation cancellation ownership ───────────────────────────

    def prepare_execution_cancellation(
        self,
        request: Any,
        *,
        batch: bool = False,
    ) -> _ExecutionCancellationScope:
        """Reserve an operation scope before an asynchronous worker starts.

        GUI callers use this immediately before exposing the cancel action so
        a cancel click can never race ahead of Application registration.
        ``release_execution_cancellation`` releases the retained terminal
        scope after the worker's completion signal has been projected.
        """
        return self._obtain_cancellation_scope(request, batch=batch, claim=False, retain=True)

    def release_execution_cancellation(self, task_id: str, reservation: object) -> None:
        """Release a scope retained for an asynchronous presentation owner."""
        with self._cancellation_lock:
            scope = self._cancellation_scopes.get(task_id)
            if scope is not None and reservation is scope:
                self._remove_cancellation_scope_locked(scope)

    def cancel(self, task_id: str) -> None:
        """Cancel one canonical Application operation.

        Before Runtime admission this only trips the operation-owned token,
        so a request that never reaches Runtime cannot leave a pending Runtime
        cancellation behind.  During Runtime execution only the currently
        admitted child is forwarded; future batch children are cancelled by
        Application before they enter Runtime.
        """
        runtime_target: str | None = None
        tracked_scope: _ExecutionCancellationScope | None = None
        with self._cancellation_lock:
            scope = self._cancellation_scopes.get(task_id)
            if scope is None:
                return
            elif scope.phase != "committed":
                scope.token.cancel(reason="user_cancelled")
                active_task_id = scope.active_runtime_task_id
                if (
                    scope.phase == "runtime"
                    and active_task_id is not None
                    and active_task_id not in scope.runtime_cancel_inflight_ids
                    and active_task_id not in scope.runtime_cancelled_ids
                ):
                    scope.runtime_cancel_inflight_ids.add(active_task_id)
                    runtime_target = active_task_id
                    tracked_scope = scope

        if runtime_target is None:
            return
        if self._runtime_port is None:
            if tracked_scope is not None:
                with self._cancellation_lock:
                    tracked_scope.runtime_cancel_inflight_ids.discard(runtime_target)
            return

        try:
            self._runtime_port.cancel(runtime_target)
        except BaseException:
            release_target = False
            if tracked_scope is not None:
                with self._cancellation_lock:
                    tracked_scope.runtime_cancel_inflight_ids.discard(runtime_target)
                    if runtime_target in tracked_scope.runtime_release_pending_ids:
                        tracked_scope.runtime_release_pending_ids.discard(runtime_target)
                        release_target = True
            if release_target:
                self._release_runtime_cancellation(runtime_target)
            raise
        else:
            release_target = False
            if tracked_scope is not None:
                with self._cancellation_lock:
                    tracked_scope.runtime_cancel_inflight_ids.discard(runtime_target)
                    tracked_scope.runtime_cancelled_ids.add(runtime_target)
                    if runtime_target in tracked_scope.runtime_release_pending_ids:
                        tracked_scope.runtime_release_pending_ids.discard(runtime_target)
                        release_target = True
            if release_target:
                self._release_runtime_cancellation(runtime_target)

    @staticmethod
    def _cancellation_scope_shape(request: Any, *, batch: bool) -> tuple[str, tuple[str, ...], tuple[str, ...]]:
        canonical_task_id = str(request.request_id)
        if batch:
            runtime_task_ids = tuple(f"{canonical_task_id}-{index}" for index, _ in enumerate(request.input_refs))
        else:
            runtime_task_ids = (canonical_task_id,)
        aliases = tuple(dict.fromkeys((canonical_task_id, *runtime_task_ids)))
        return canonical_task_id, aliases, runtime_task_ids

    def _obtain_cancellation_scope(
        self,
        request: Any,
        *,
        batch: bool,
        claim: bool,
        retain: bool,
    ) -> _ExecutionCancellationScope:
        canonical_task_id, aliases, runtime_task_ids = self._cancellation_scope_shape(request, batch=batch)
        with self._cancellation_lock:
            existing = self._cancellation_scopes.get(canonical_task_id)
            if existing is not None and existing.phase == "committed":
                if existing.retain_until_release:
                    raise ControllerError(f"Cancellation scope is awaiting release: {canonical_task_id!r}")
                self._remove_cancellation_scope_locked(existing)
                existing = None
            if existing is not None:
                if existing.aliases != aliases or existing.runtime_task_ids != runtime_task_ids:
                    raise ControllerError(f"Cancellation scope shape changed for {canonical_task_id!r}")
                if claim and existing.claimed:
                    raise ControllerError(f"Execution is already active for {canonical_task_id!r}")
                existing.retain_until_release = existing.retain_until_release or retain
                if claim:
                    existing.claimed = True
                return existing

            for alias in aliases:
                if alias in self._cancellation_scopes:
                    raise ControllerError(f"Cancellation task id is already active: {alias!r}")
            scope = _ExecutionCancellationScope(
                canonical_task_id=canonical_task_id,
                aliases=aliases,
                runtime_task_ids=runtime_task_ids,
                retain_until_release=retain,
                claimed=claim,
            )
            for alias in aliases:
                self._cancellation_scopes[alias] = scope
            return scope

    def _remove_cancellation_scope_locked(self, scope: _ExecutionCancellationScope) -> None:
        for alias in scope.aliases:
            if self._cancellation_scopes.get(alias) is scope:
                self._cancellation_scopes.pop(alias, None)

    def _begin_runtime_task(self, scope: _ExecutionCancellationScope, task_id: str) -> bool:
        """Linearize cancellation against one imminent Runtime call."""
        reservation = self._runtime_cancellation_reservation()
        if reservation is not None:
            reservation.reserve_cancellation(task_id)
        with self._cancellation_lock:
            if scope.token.is_cancelled:
                admitted = False
            else:
                scope.phase = "runtime"
                scope.active_runtime_task_id = task_id
                admitted = True
        if not admitted and reservation is not None:
            reservation.release_cancellation(task_id)
        return admitted

    def _finish_runtime_task(self, scope: _ExecutionCancellationScope, task_id: str) -> None:
        """Close one Runtime call's Application-owned admission window."""
        release_now = False
        with self._cancellation_lock:
            if scope.active_runtime_task_id == task_id:
                scope.active_runtime_task_id = None
            if task_id in scope.runtime_cancel_inflight_ids:
                scope.runtime_release_pending_ids.add(task_id)
            else:
                release_now = True
        if release_now:
            self._release_runtime_cancellation(task_id)

    def _runtime_cancellation_reservation(self) -> Any | None:
        """Return the optional production handoff capability when present."""
        from docwen_application.ports.runtime import CancellationReservationPort

        runtime_port = self._runtime_port
        return runtime_port if isinstance(runtime_port, CancellationReservationPort) else None

    def _release_runtime_cancellation(self, task_id: str) -> None:
        reservation = self._runtime_cancellation_reservation()
        if reservation is not None:
            reservation.release_cancellation(task_id)

    def _commit_without_runtime(self, scope: _ExecutionCancellationScope) -> bool:
        with self._cancellation_lock:
            cancelled = scope.token.is_cancelled
            scope.phase = "committed"
            return cancelled

    def _complete_cancellation_scope(self, scope: _ExecutionCancellationScope) -> None:
        with self._cancellation_lock:
            scope.phase = "committed"
            scope.claimed = False
            if not scope.retain_until_release:
                self._remove_cancellation_scope_locked(scope)

    def _freeze_manifest_context(self, request: Any) -> Any:
        """Freeze one config snapshot and original-input manifest context."""
        from copy import deepcopy

        from docwen_core.models.conversion_manifest import ConversionManifestContext

        config_snapshot = deepcopy(getattr(request, "config_snapshot", {}))
        if not config_snapshot and self._config_port is not None:
            captured = cast(object, self._config_port.snapshot())
            config_snapshot = deepcopy(captured) if isinstance(captured, dict) else {}
        context = getattr(request, "manifest_context", None)
        if context is None:
            candidate = ConversionManifestContext.from_request_inputs(request.input_refs, config_snapshot)
            context = candidate if candidate.policy.save_to_output or self._config_port is not None else None
        if config_snapshot == getattr(request, "config_snapshot", {}) and context is getattr(
            request, "manifest_context", None
        ):
            return request
        return replace(request, config_snapshot=config_snapshot, manifest_context=context)

    def _persist_output_manifests(self, request: Any, result: Any) -> Any:
        """Invoke the optional Runtime sidecar capability without changing terminal truth."""
        context = getattr(request, "manifest_context", None)
        if context is None or not context.policy.save_to_output:
            return result
        from docwen_application.ports.runtime import OutputManifestPersistencePort

        runtime_port = self._runtime_port
        if runtime_port is None or not isinstance(runtime_port, OutputManifestPersistencePort):
            return result
        try:
            return runtime_port.persist_output_manifests(request, result)
        except Exception:
            return self._mark_manifest_write_failure(result)

    @staticmethod
    def _mark_manifest_write_failure(result: Any) -> Any:
        """Append a bounded warning while preserving success/error/artifacts."""
        from docwen_core.models.result import ConversionDiagnostic, ConversionResult

        if isinstance(result, list):
            return [ApplicationController._mark_manifest_write_failure(item) for item in result]
        if not isinstance(result, ConversionResult):
            return result
        if result.error is not None and result.error.error_type == "cancelled":
            return result
        if any(item.code == "OUTPUT_MANIFEST_WRITE_FAILED" for item in result.diagnostics):
            return result
        return replace(
            result,
            diagnostics=[
                *result.diagnostics,
                ConversionDiagnostic(
                    level="warning",
                    code="OUTPUT_MANIFEST_WRITE_FAILED",
                    message="The configured output manifest could not be written.",
                ),
            ],
        )

    # ── Internal command assembly ───────────────────────────────────

    def _convert_command(self) -> Any:
        """Build the internal single-file command behind admitted execution."""
        if self._runtime_port is None:
            raise ControllerError("No runtime port configured — cannot create ConvertCommand")
        from docwen_application.commands.convert import ConvertCommand

        return ConvertCommand(self._runtime_port)

    def _batch_command(self, *, continue_on_error: bool = True) -> Any:
        """Build the internal batch command behind admitted execution."""
        if self._runtime_port is None:
            raise ControllerError("No runtime port configured — cannot create BatchCommand")
        from docwen_application.commands.batch import BatchCommand

        return BatchCommand(self._runtime_port, continue_on_error=continue_on_error)

    def _aggregate_command(self, action_name: str) -> Any:
        """Build the internal aggregate command behind admitted execution."""
        if self._runtime_port is None:
            raise ControllerError("No runtime port configured — cannot create AggregateCommand")
        from docwen_application.commands.batch import AggregateCommand

        return AggregateCommand(self._runtime_port, action_name=action_name)

    # ── Convenience: direct execution ───────────────────────────────

    def execute_single(self, request: Any) -> Any:
        """Convenience: execute a single-file conversion directly.

        Before dispatching to the runtime, frozen Core admission facts are
        enforced and OOXML signature presence is recorded. If the admitted
        format is a non-hub format (e.g. ``doc``, ``wps``, ``rtf``, ``odt``),
        it is pre-converted to the hub format (e.g. ``docx``) via the office
        bridge before the plugin runs.

        Args:
            request: A ``ConversionRequest`` with exactly one source and any
                typed resources declared by that route.

        Returns:
            A ``ConversionResult``.
        """
        from docwen_core.detection import enforce_file_admission, freeze_ooxml_signature_info

        request = enforce_file_admission(request)
        request = freeze_ooxml_signature_info(request)
        request = self._freeze_manifest_context(request)
        original_request = request
        scope = self._obtain_cancellation_scope(request, batch=False, claim=True, retain=False)
        managed: _ManagedPreconversion | None = None
        try:
            prepared = self._maybe_preconvert(request, cancellation=scope.token.view(), batch=False)
            managed = prepared if isinstance(prepared, _ManagedPreconversion) else None
            request = prepared.payload if isinstance(prepared, _ManagedPreconversion) else prepared
            manifest_request = managed.manifest_request if managed is not None else None
            manifest_request = manifest_request or request
            from docwen_core.models.result import ConversionResult

            if isinstance(request, _PreconversionTerminal):
                self._commit_without_runtime(scope)
                return self._persist_output_manifests(manifest_request, request.results[0])
            if isinstance(request, ConversionResult):
                if self._commit_without_runtime(scope):
                    result = self._cancelled_results(original_request, batch=False)[0]
                    return self._persist_output_manifests(manifest_request, result)
                return self._persist_output_manifests(manifest_request, request)
            cmd = self._convert_command()
            task_id = str(request.request_id)
            if not self._begin_runtime_task(scope, task_id):
                result = self._cancelled_results(original_request, batch=False)[0]
                return self._persist_output_manifests(manifest_request, result)
            try:
                result = cmd.execute(request)
                return self._persist_output_manifests(request, result)
            finally:
                self._finish_runtime_task(scope, task_id)
        finally:
            try:
                if managed is not None:
                    managed.cleanup()
            finally:
                self._complete_cancellation_scope(scope)

    def execute_aggregate(self, request: Any, action_name: str) -> Any:
        """Execute one many-to-one request inside the shared cancellation scope."""

        from docwen_core.detection import enforce_file_admission, freeze_ooxml_signature_info

        request = enforce_file_admission(request)
        request = freeze_ooxml_signature_info(request)
        request = self._freeze_manifest_context(request)
        scope = self._obtain_cancellation_scope(request, batch=False, claim=True, retain=False)
        task_id = str(request.request_id)
        try:
            command = self._aggregate_command(action_name)
            if not self._begin_runtime_task(scope, task_id):
                result = self._cancelled_results(request, batch=False)[0]
                return self._persist_output_manifests(request, result)
            try:
                result = command.execute(request)
                return self._persist_output_manifests(request, result)
            finally:
                self._finish_runtime_task(scope, task_id)
        finally:
            self._complete_cancellation_scope(scope)

    def execute_batch(self, request: Any) -> list[Any]:
        """Convenience: execute a batch conversion directly.

        Each input file is pre-converted to its hub format if needed
        before dispatching to the runtime.

        Args:
            request: A ``ConversionRequest`` containing only independent
                source inputs. Typed resources belong to ``execute_single``.

        Returns:
            A ``list[ConversionResult]``.
        """
        from docwen_core.detection import enforce_file_admission, freeze_ooxml_signature_info

        if any(ref.input_role != "source" for ref in request.input_refs):
            raise ValueError("batch conversion accepts only independent source inputs")

        request = enforce_file_admission(request)
        request = freeze_ooxml_signature_info(request)
        request = self._freeze_manifest_context(request)
        original_request = request
        scope = self._obtain_cancellation_scope(request, batch=True, claim=True, retain=False)
        managed: _ManagedPreconversion | None = None
        try:
            prepared = self._maybe_preconvert(request, cancellation=scope.token.view(), batch=True)
            managed = prepared if isinstance(prepared, _ManagedPreconversion) else None
            request = prepared.payload if isinstance(prepared, _ManagedPreconversion) else prepared
            manifest_request = managed.manifest_request if managed is not None else None
            manifest_request = manifest_request or (
                request.request if isinstance(request, _PreconversionBatchPlan) else request
            )
            from docwen_core.models.result import ConversionResult

            if isinstance(request, _PreconversionTerminal):
                self._commit_without_runtime(scope)
                return self._persist_output_manifests(manifest_request, request.results)
            if isinstance(request, ConversionResult):
                if self._commit_without_runtime(scope):
                    results = self._cancelled_results(original_request, batch=True)
                    return self._persist_output_manifests(manifest_request, results)
                return self._persist_output_manifests(manifest_request, [request])
            if isinstance(request, _PreconversionBatchPlan):
                if not request.request.input_refs:
                    self._commit_without_runtime(scope)
                    results = [slot for slot in request.result_slots if slot is not None]
                    return self._persist_output_manifests(manifest_request, results)
                runtime_results = self._execute_preconverted_batch(request, scope)
                runtime_iter = iter(runtime_results)
                results = [next(runtime_iter) if slot is None else slot for slot in request.result_slots]
                return self._persist_output_manifests(manifest_request, results)
            results = self._execute_runtime_batch(request, scope)
            return self._persist_output_manifests(manifest_request, results)
        finally:
            try:
                if managed is not None:
                    managed.cleanup()
            finally:
                self._complete_cancellation_scope(scope)

    @staticmethod
    def _cancelled_results(
        request: Any,
        *,
        batch: bool,
        existing_slots: list[Any | None] | None = None,
    ) -> list[Any]:
        """Build aligned cancellation results without rewriting terminal slots."""
        results: list[Any] = []
        for index, _ref in enumerate(request.input_refs):
            existing = existing_slots[index] if existing_slots is not None and index < len(existing_slots) else None
            if existing is not None:
                results.append(existing)
                continue
            task_id = f"{request.request_id}-{index}" if batch else str(request.request_id)
            results.append(ApplicationController._cancelled_result(task_id))
        return results

    @staticmethod
    def _cancelled_result(task_id: str) -> Any:
        """Build one structured cancellation result."""
        from docwen_core.models.result import ConversionErrorInfo, ConversionResult

        return ConversionResult(
            task_id=task_id,
            success=False,
            error=ConversionErrorInfo(error_type="cancelled", message="Task was cancelled"),
        )

    def _execute_runtime_batch(self, request: Any, scope: _ExecutionCancellationScope) -> list[Any]:
        """Execute a regular batch while Application owns future cancellation."""
        if not hasattr(request, "input_refs") or len(request.input_refs) == 0:
            raise ValueError("ConversionRequest must have at least one input file")

        from docwen_core.models.request import ConversionRequest

        runtime_results: list[Any] = []
        cmd = self._convert_command()
        for index, input_ref in enumerate(request.input_refs):
            task_id = f"{request.request_id}-{index}"
            child_request = ConversionRequest(
                request_id=task_id,
                input_refs=[input_ref],
                target_format=request.target_format,
                action_name=getattr(request, "action_name", ""),
                options=dict(getattr(request, "options", {})),
                output_policy=request.output_policy,
                config_snapshot=dict(getattr(request, "config_snapshot", {})),
                manifest_context=(
                    request.manifest_context.for_input(index) if request.manifest_context is not None else None
                ),
            )
            if not self._begin_runtime_task(scope, task_id):
                runtime_results.append(self._cancelled_result(task_id))
                continue
            try:
                runtime_results.append(cmd.execute(child_request))
            finally:
                self._finish_runtime_task(scope, task_id)
        return runtime_results

    def _execute_preconverted_batch(
        self,
        plan: _PreconversionBatchPlan,
        scope: _ExecutionCancellationScope,
    ) -> list[Any]:
        """Execute prepared refs sequentially with per-input output anchors.

        This is the narrow preconversion counterpart of the default
        ``BatchWorkflow``.  It keeps the same continue-on-error ordering while
        allowing each physically staged input to retain its own source policy.
        """
        from docwen_core.models.request import ConversionRequest

        runtime_results: list[Any] = []
        cmd = self._convert_command()
        for input_ref, output_policy, original_index in zip(
            plan.request.input_refs,
            plan.output_policies,
            plan.input_indices,
            strict=True,
        ):
            task_id = f"{plan.request.request_id}-{original_index}"
            child_request = ConversionRequest(
                request_id=task_id,
                input_refs=[input_ref],
                target_format=plan.request.target_format,
                action_name=plan.request.action_name,
                options=dict(plan.request.options),
                output_policy=output_policy,
                config_snapshot=dict(plan.request.config_snapshot),
                manifest_context=(
                    plan.request.manifest_context.for_input(original_index)
                    if plan.request.manifest_context is not None
                    else None
                ),
            )
            if not self._begin_runtime_task(scope, task_id):
                runtime_results.append(self._cancelled_result(task_id))
                continue
            try:
                runtime_results.append(cmd.execute(child_request))
            finally:
                self._finish_runtime_task(scope, task_id)
        return runtime_results

    def _maybe_preconvert(
        self,
        request: Any,
        *,
        cancellation: CancellationTokenView,
        batch: bool,
    ) -> Any:
        """Pre-convert admitted non-hub formats from ``FileRef.format``.

        For each input ref, if the frozen content-derived format is a non-hub format
        (e.g. ``doc``, ``wps``, ``rtf``, ``odt``), it is pre-converted
        to the hub format (e.g. ``docx``) via the office bridge.  A managed
        payload owns any temporary staging until the public execute method
        synchronously consumes the updated request or result.

        If pre-conversion fails for any input, an error ``ConversionResult``
        is returned directly (for single-file) or the failed ref is
        replaced with an error result for batch.

        Args:
            request: A ``ConversionRequest`` with one or more ``input_refs``.

        Returns:
            The original request when no preconversion is needed; otherwise a
            managed converted request, batch plan, or failure result.
        """
        from docwen_application.preconversion.chain_resolver import resolve_chain
        from docwen_application.preconversion.intermediate_policy import build_intermediate_record_if_enabled
        from docwen_application.preconversion.pre_converter import (
            PreConversionFailure,
            backend_priority_spec,
            pre_convert,
        )
        from docwen_core.formats import get_media_type
        from docwen_core.models.request import PRECONVERSION_INTERMEDIATES_OPTION

        # Fast path: nothing to pre-convert.  FileRef.format is the canonical
        # content-derived identity frozen by the ingress admission boundary;
        # never reopen the path or fall back to its suffix here.
        needs_pre = False
        for ref in request.input_refs:
            if ref.input_role != "source":
                continue
            chain = resolve_chain(
                ref.format,
                request.target_format,
                action_name=request.action_name,
            )
            if len(chain) > 1:
                needs_pre = True
                break

        if not needs_pre:
            return request

        if cancellation.is_cancelled:
            return _PreconversionTerminal(self._cancelled_results(request, batch=batch))

        # Pre-convert each input that needs it
        from copy import deepcopy

        from docwen_core.detection.ooxml_signature import OOXML_SIGNATURE_INFO_METADATA_KEY
        from docwen_core.models.conversion_manifest import ConversionManifestContext, PreconversionStep
        from docwen_core.models.file_inspection import (
            FILE_ADMISSION_ACCEPTANCE_METADATA_KEY,
            FILE_INSPECTION_METADATA_KEY,
        )
        from docwen_core.models.file_ref import FileRef
        from docwen_core.models.request import ConversionRequest
        from docwen_core.models.result import ConversionDiagnostic, ConversionErrorInfo, ConversionResult

        config_snapshot = deepcopy(getattr(request, "config_snapshot", {}))
        if not config_snapshot and self._config_port is not None and getattr(request, "manifest_context", None) is None:
            captured_snapshot = self._config_port.snapshot()
            config_snapshot = deepcopy(captured_snapshot)
        manifest_context = getattr(request, "manifest_context", None)
        if manifest_context is None:
            manifest_context = ConversionManifestContext.from_request_inputs(request.input_refs, config_snapshot)

        def current_manifest_request() -> Any:
            return replace(
                request,
                config_snapshot=deepcopy(config_snapshot),
                manifest_context=manifest_context,
            )

        temp_owner = tempfile.TemporaryDirectory(prefix="docwen_pre_", ignore_cleanup_errors=True)
        new_refs: list[FileRef] = []
        errors: list[tuple[int, ConversionResult]] = []
        result_slots: list[ConversionResult | None] = []
        intermediate_records: list[dict[str, Any]] = []
        output_policies: list[Any] = []
        input_indices: list[int] = []

        try:
            for idx, ref in enumerate(request.input_refs):
                if cancellation.is_cancelled:
                    terminal = _PreconversionTerminal(
                        self._cancelled_results(request, batch=batch, existing_slots=result_slots)
                    )
                    return _ManagedPreconversion(
                        payload=terminal,
                        temp_owner=temp_owner,
                        manifest_request=current_manifest_request(),
                    )
                if ref.input_role != "source":
                    new_refs.append(ref)
                    output_policies.append(request.output_policy)
                    input_indices.append(idx)
                    result_slots.append(None)
                    continue

                actual_format = ref.format
                chain = resolve_chain(
                    actual_format,
                    request.target_format,
                    action_name=request.action_name,
                )
                if len(chain) <= 1:
                    new_refs.append(ref)
                    output_policies.append(request.output_policy)
                    input_indices.append(idx)
                    result_slots.append(None)
                    continue

                staging_dir = Path(temp_owner.name) / str(idx)
                staging_dir.mkdir(parents=True, exist_ok=True)
                priority_key, default_priority = backend_priority_spec(actual_format)
                backend_priority = self._configured_priority(config_snapshot, priority_key, default_priority)
                pre_result = pre_convert(
                    ref.path,
                    actual_format,
                    staging_dir=str(staging_dir),
                    cancel=cancellation,
                    backend_priority=backend_priority,
                )
                cleanup_diagnostics = (
                    [
                        ConversionDiagnostic(
                            level="warning",
                            code="OFFICE_CLEANUP_FAILED",
                            message=pre_result.cleanup_message,
                        )
                    ]
                    if isinstance(pre_result, PreConversionFailure)
                    and pre_result.cleanup_failed
                    and pre_result.cleanup_message
                    else []
                )
                bridge_cancelled = isinstance(pre_result, PreConversionFailure) and pre_result.cancelled
                if cancellation.is_cancelled or bridge_cancelled:
                    manifest_context = manifest_context.with_step(
                        PreconversionStep(
                            input_index=idx,
                            source_format=actual_format,
                            target_format=chain[0],
                            status="cancelled",
                            diagnostic_code=(
                                pre_result.diagnostic_code if isinstance(pre_result, PreConversionFailure) else ""
                            ),
                        )
                    )
                    cancelled_results = self._cancelled_results(
                        request,
                        batch=batch,
                        existing_slots=result_slots,
                    )
                    if cleanup_diagnostics:
                        cancelled_results[idx].diagnostics.extend(cleanup_diagnostics)
                    terminal = _PreconversionTerminal(cancelled_results)
                    return _ManagedPreconversion(
                        payload=terminal,
                        temp_owner=temp_owner,
                        manifest_request=current_manifest_request(),
                    )
                if pre_result is None or isinstance(pre_result, PreConversionFailure):
                    bridge_message = (
                        pre_result.message
                        if pre_result is not None
                        else (
                            "No external office backend succeeded. Install Microsoft Office/WPS "
                            "(Windows COM) or LibreOffice."
                        )
                    )
                    manifest_context = manifest_context.with_step(
                        PreconversionStep(
                            input_index=idx,
                            source_format=actual_format,
                            target_format=chain[0],
                            status="failed",
                            diagnostic_code=pre_result.diagnostic_code if pre_result is not None else "",
                        )
                    )
                    err = ConversionResult(
                        task_id=request.request_id if not batch else f"{request.request_id}-{idx}",
                        success=False,
                        diagnostics=cleanup_diagnostics,
                        error=ConversionErrorInfo(
                            error_type=pre_result.error_type if pre_result is not None else "dependency_missing",
                            message=(f"Cannot pre-convert {actual_format.upper()} to {chain[0]}: {bridge_message}"),
                            diagnostic_code=pre_result.diagnostic_code if pre_result is not None else "",
                        ),
                    )
                    errors.append((idx, err))
                    result_slots.append(err)
                else:
                    manifest_context = manifest_context.with_step(
                        PreconversionStep(
                            input_index=idx,
                            source_format=actual_format,
                            target_format=chain[0],
                            status="completed",
                            backend=pre_result.backend,
                        )
                    )
                    intermediate_record = build_intermediate_record_if_enabled(
                        pre_result.pre_converted_path,
                        Path(ref.path).stem,
                        actual_format,
                        target_format=chain[0],
                        backend=pre_result.backend,
                        config_snapshot=config_snapshot,
                    )
                    if intermediate_record is not None:
                        intermediate_record["applies_to_input_path"] = pre_result.pre_converted_path
                        intermediate_records.append(intermediate_record)
                    derived_metadata = deepcopy(ref.metadata)
                    source_inspection = derived_metadata.pop(FILE_INSPECTION_METADATA_KEY, None)
                    derived_metadata.pop(FILE_ADMISSION_ACCEPTANCE_METADATA_KEY, None)
                    derived_metadata.pop(OOXML_SIGNATURE_INFO_METADATA_KEY, None)
                    derived_metadata["_docwen_preconversion_source"] = {
                        "path": ref.path,
                        "format": actual_format,
                        "category": ref.category,
                        "warning_message": ref.warning_message,
                        "inspection": source_inspection if isinstance(source_inspection, dict) else None,
                    }
                    new_refs.append(
                        FileRef(
                            path=pre_result.pre_converted_path,
                            format=chain[0],  # hub format
                            category=ref.category,
                            encoding=ref.encoding,
                            warning_message="",
                            size_bytes=Path(pre_result.pre_converted_path).stat().st_size,
                            input_kind=ref.input_kind,
                            input_role=ref.input_role,
                            logical_path=ref.logical_path,
                            media_type=get_media_type(chain[0]),
                            metadata=derived_metadata,
                        )
                    )
                    output_policies.append(self._source_anchored_output_policy(request.output_policy, ref.path))
                    input_indices.append(idx)
                    result_slots.append(None)

            if cancellation.is_cancelled:
                terminal = _PreconversionTerminal(
                    self._cancelled_results(request, batch=batch, existing_slots=result_slots)
                )
                return _ManagedPreconversion(
                    payload=terminal,
                    temp_owner=temp_owner,
                    manifest_request=current_manifest_request(),
                )

            if errors and not batch:
                return _ManagedPreconversion(
                    payload=errors[0][1],
                    temp_owner=temp_owner,
                    manifest_request=current_manifest_request(),
                )

            # Build the new request before handing temporary ownership to the
            # public execute method.  Any failure during reconstruction must
            # release staging here rather than relying on garbage collection.
            converted_options = deepcopy(request.options)
            if intermediate_records:
                converted_options[PRECONVERSION_INTERMEDIATES_OPTION] = intermediate_records

            converted_request = ConversionRequest(
                request_id=request.request_id,
                input_refs=new_refs,
                target_format=request.target_format,
                action_name=request.action_name,
                options=converted_options,
                output_policy=(
                    next(
                        (
                            policy
                            for ref, policy in zip(request.input_refs, output_policies, strict=True)
                            if ref.input_role == "source"
                        ),
                        request.output_policy,
                    )
                    if not batch
                    else request.output_policy
                ),
                config_snapshot=deepcopy(config_snapshot),
                manifest_context=manifest_context,
            )
            if batch:
                payload = _PreconversionBatchPlan(
                    request=converted_request,
                    result_slots=result_slots,
                    output_policies=output_policies,
                    input_indices=input_indices,
                )
            else:
                payload = converted_request
            return _ManagedPreconversion(
                payload=payload,
                temp_owner=temp_owner,
                manifest_request=current_manifest_request(),
            )
        except BaseException:
            temp_owner.cleanup()
            raise

    @staticmethod
    def _source_anchored_output_policy(output_policy: Any, source_path: str) -> Any:
        """Preserve same-as-source semantics after the physical input moves."""
        if output_policy.output_path or output_policy.output_dir:
            return output_policy
        return replace(output_policy, output_dir=str(Path(source_path).parent))

    def _configured_priority(
        self,
        config_snapshot: dict[str, Any],
        key: str,
        default: list[str],
    ) -> list[str]:
        """Resolve one ordered backend list from the admitted snapshot."""
        current: Any = config_snapshot
        found = True
        for part in key.split("."):
            if not isinstance(current, dict) or part not in current:
                found = False
                break
            current = current[part]
        if not found:
            current = default
        if isinstance(current, (list, tuple)):
            return [str(item) for item in current if isinstance(item, str)]
        return list(default)

    # ── Non-runtime dependency accessors ───────────────────────────

    @property
    def config_port(self) -> ConfigPort | None:
        return self._config_port

    @property
    def presenter_port(self) -> PresenterPort | None:
        return self._presenter_port

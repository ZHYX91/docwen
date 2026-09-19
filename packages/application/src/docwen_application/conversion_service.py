"""Plan-first application service shared by interactive and machine entry points."""

from __future__ import annotations

import hashlib
import logging
import sys
import threading
from copy import deepcopy
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Literal, Protocol
from uuid import uuid4

from docwen_application.bundle_mapping import (
    BundleMappingError,
    build_bundle_draft,
    validate_physical_page_diagnostics,
)
from docwen_application.composed_capabilities import composed_capability_bindings
from docwen_application.conversion_capabilities import (
    CAPABILITY_BINDINGS,
    CAPABILITY_BY_ID,
    CapabilityBinding,
)
from docwen_application.conversion_contracts import (
    DOCX_TO_MARKDOWN_CAPABILITY_ID,
    MARKDOWN_TO_DOCX_CAPABILITY_ID,
    ConversionPlan,
    ConversionPlanRequest,
    ConversionServiceError,
    ConversionTaskOutcome,
    LocalInputHandle,
    MachineCapability,
)
from docwen_application.conversion_options import resolve_conversion_options
from docwen_application.conversion_requests import build_conversion_request
from docwen_application.conversion_routes import resolve_conversion_route_plan
from docwen_application.optimization_catalog import parse_optimization_catalog
from docwen_application.ports.runtime import ArtifactBundleCommitPort
from docwen_application.runtime_capability_catalog import RuntimeCapabilityCatalog
from docwen_core.detection import FileAdmissionPathError
from docwen_core.models import (
    ArtifactBundleValidationError,
    ConversionRequest,
    ConversionResult,
    validate_artifact_bundle_draft,
)
from docwen_core.models.resolved_numbering import (
    ResolvedNumberingPortError,
    load_resolved_numbering_bytes,
)
from docwen_core.paths import filesystem_path

logger = logging.getLogger(__name__)

_HASH_CHUNK_BYTES = 1024 * 1024


class _ExecutionController(Protocol):
    @property
    def has_runtime(self) -> bool: ...

    def describe_runtime_capabilities(self) -> dict[str, Any]: ...

    def prepare_execution_cancellation(self, request: Any, *, batch: bool = False) -> object: ...

    def release_execution_cancellation(self, task_id: str, reservation: object) -> None: ...

    def execute_single(self, request: Any) -> Any: ...

    def execute_aggregate(self, request: Any, action_name: str) -> Any: ...

    def cancel(self, task_id: str) -> None: ...


@dataclass(frozen=True, slots=True)
class _RuntimeCapabilityState:
    availability: Literal["available", "limited", "unavailable"]
    dependencies: tuple[dict[str, Any], ...] = ()
    limitations: tuple[dict[str, Any], ...] = ()


@dataclass(frozen=True, slots=True)
class _RuntimeDiscovery:
    gates: dict[str, bool]
    routes: dict[str, dict[str, Any]]
    error: dict[str, Any] | None = None
    bindings: tuple[CapabilityBinding, ...] = CAPABILITY_BINDINGS
    catalog: RuntimeCapabilityCatalog | None = None


@dataclass(slots=True)
class _PlanRecord:
    public: ConversionPlan
    request: ConversionPlanRequest
    binding: CapabilityBinding


@dataclass(slots=True)
class _TaskRecord:
    request: ConversionRequest
    public_options: dict[str, Any]
    staging_root: str
    reservation: object
    binding: CapabilityBinding
    state: str = "accepted"


class ConversionService:
    """Own capability discovery, immutable planning, execution, and cancellation."""

    def __init__(self, controller: _ExecutionController, bundle_committer: ArtifactBundleCommitPort) -> None:
        self._controller = controller
        self._bundle_committer = bundle_committer
        self._lock = threading.Lock()
        self._plans: dict[str, _PlanRecord] = {}
        self._tasks: dict[str, _TaskRecord] = {}

    def list_capabilities(self) -> tuple[MachineCapability, ...]:
        discovery = self._discover_runtime()
        capabilities: list[MachineCapability] = []
        for binding in discovery.bindings:
            state = self._runtime_capability_state(binding, discovery)
            capabilities.append(
                MachineCapability(
                    capability_id=binding.capability_id,
                    operation=binding.operation,
                    input_shape=binding.input_shape,
                    output_media_types=(binding.output_media_type,),
                    output_shape=binding.output_shape,
                    options_schema=deepcopy(binding.options_schema),
                    availability=state.availability,
                    dependencies=state.dependencies,
                    limitations=deepcopy((*binding.limitations, *state.limitations)),
                    optimization_id=binding.optimization_id,
                )
            )
        return tuple(capabilities)

    def plan(self, request: ConversionPlanRequest) -> ConversionPlan:
        request = deepcopy(request)
        binding, effective_options = self._validate_plan_request(request)
        plan = ConversionPlan(
            plan_id=f"plan.{uuid4().hex}",
            capability_id=request.capability_id,
            effective_options=effective_options,
            output_shape=binding.output_shape,
            limitations=deepcopy(binding.limitations),
        )
        with self._lock:
            self._plans[plan.plan_id] = _PlanRecord(public=deepcopy(plan), request=request, binding=deepcopy(binding))
        return plan

    def accept(self, plan_id: str, task_id: str | None = None) -> str:
        with self._lock:
            record = self._plans.pop(plan_id, None)
        if record is None:
            raise ConversionServiceError(
                "conflict", "plan_not_found", f"plan is unknown or already consumed: {plan_id}"
            )

        task_id = task_id or f"task.{uuid4().hex}"
        self._validate_identifier(task_id, field_name="task_id")
        binding, effective_options = self._validate_plan_request(record.request)
        if binding != record.binding:
            raise ConversionServiceError("conflict", "capability_changed", "capability changed after planning")
        if effective_options != record.public.effective_options:
            raise ConversionServiceError(
                "conflict",
                "plan_options_changed",
                "capability options changed after planning",
            )
        conversion_request = build_conversion_request(
            task_id,
            record.request,
            binding,
            effective_options,
        )
        reservation = self._controller.prepare_execution_cancellation(conversion_request, batch=False)
        task = _TaskRecord(
            request=conversion_request,
            public_options=dict(effective_options),
            staging_root=record.request.output.staging_root,
            reservation=reservation,
            binding=binding,
        )
        with self._lock:
            if task_id in self._tasks:
                self._controller.release_execution_cancellation(task_id, reservation)
                raise ConversionServiceError("conflict", "task_id_in_use", f"task id is already in use: {task_id}")
            self._tasks[task_id] = task
        return task_id

    def execute_accepted(self, task_id: str) -> ConversionTaskOutcome:
        with self._lock:
            task = self._tasks.get(task_id)
            if task is None:
                raise ConversionServiceError("conflict", "task_not_found", f"task is not accepted: {task_id}")
            if task.state != "accepted":
                raise ConversionServiceError("conflict", "task_not_accepted", f"task cannot execute from {task.state}")
            task.state = "running"

        try:
            self._validate_execution_snapshot(task.request, task.staging_root)
            try:
                if task.binding.input_cardinality == "many":
                    raw_result = self._controller.execute_aggregate(task.request, task.binding.action_name)
                else:
                    raw_result = self._controller.execute_single(task.request)
            except FileAdmissionPathError as exc:
                raise ConversionServiceError("security", exc.error_type, str(exc)) from exc
            if not isinstance(raw_result, ConversionResult):
                raise ConversionServiceError(
                    "internal",
                    "invalid_runtime_result",
                    "runtime returned an unsupported result type",
                )
            outcome = self._outcome_from_result(task, raw_result)
            return outcome
        finally:
            self._controller.release_execution_cancellation(task_id, task.reservation)
            with self._lock:
                task.state = "terminal"

    def cancel(self, task_id: str) -> str:
        with self._lock:
            task = self._tasks.get(task_id)
            if task is None:
                return "not_found"
            if task.state == "terminal":
                return "already_terminal"
        self._controller.cancel(task_id)
        return "cancellation_requested"

    def _outcome_from_result(self, task: _TaskRecord, result: ConversionResult) -> ConversionTaskOutcome:
        if result.task_id != task.request.request_id:
            raise ConversionServiceError(
                "internal",
                "runtime_task_mismatch",
                "runtime result belongs to a different task",
            )
        if not result.success:
            self._discard_result_artifacts(task.staging_root, result)
            bound_diagnostics = sorted(
                {diagnostic.artifact_id for diagnostic in result.diagnostics if diagnostic.artifact_id is not None}
            )
            if bound_diagnostics:
                raise ConversionServiceError(
                    "conversion_failed",
                    "dangling_diagnostic_artifact",
                    "failed or cancelled runtime result cannot bind diagnostics to output artifacts",
                    details={"artifact_ids": bound_diagnostics},
                )
            state: Literal["failed", "cancelled"] = (
                "cancelled" if result.error is not None and result.error.error_type == "cancelled" else "failed"
            )
            return ConversionTaskOutcome(
                task_id=result.task_id,
                state=state,
                bundle=None,
                diagnostics=tuple(result.diagnostics),
                metrics=result.metrics,
                error=result.error,
            )

        try:
            draft = build_bundle_draft(
                profile=task.binding.bundle_profile,
                output_media_type=task.binding.output_media_type,
                artifacts=result.artifacts,
            )
        except BundleMappingError as exc:
            self._discard_result_artifacts(task.staging_root, result)
            raise ConversionServiceError(
                exc.category,
                exc.code,
                str(exc),
                details=exc.details,
            ) from exc
        try:
            validate_artifact_bundle_draft(draft)
        except ArtifactBundleValidationError as exc:
            self._discard_result_artifacts(task.staging_root, result)
            raise ConversionServiceError(
                "conversion_failed",
                exc.code,
                str(exc),
            ) from exc
        artifact_ids = {artifact.artifact_id for artifact in draft.artifacts}
        dangling_diagnostics = sorted(
            {
                diagnostic.artifact_id
                for diagnostic in result.diagnostics
                if diagnostic.artifact_id is not None and diagnostic.artifact_id not in artifact_ids
            }
        )
        if dangling_diagnostics:
            self._discard_result_artifacts(task.staging_root, result)
            raise ConversionServiceError(
                "conversion_failed",
                "dangling_diagnostic_artifact",
                "runtime diagnostic references an artifact outside the output bundle",
                details={"artifact_ids": dangling_diagnostics},
            )
        if task.binding.runtime_route_id == CAPABILITY_BY_ID[DOCX_TO_MARKDOWN_CAPABILITY_ID].runtime_route_id:
            expected_recognition = task.public_options.get("recognize_text")
            expected_resources = task.public_options.get("preserve_resources")
            expected_placement = task.public_options.get("ocr_placement")
            image_relations = [
                relation for relation in draft.relations if relation.type == "resource_of" and relation.role == "image"
            ]
            ocr_relations = [
                relation
                for relation in draft.relations
                if relation.type == "fragment_of" and relation.role == "ocr_text"
            ]
            invalid_public_contract = (
                not isinstance(expected_recognition, bool)
                or not isinstance(expected_resources, bool)
                or expected_placement not in {"image_md", "main_md"}
            )
            producer_drift = (not expected_resources and bool(image_relations)) or (
                (not expected_recognition or expected_placement == "main_md") and bool(ocr_relations)
            )
            if invalid_public_contract or producer_drift:
                self._discard_result_artifacts(task.staging_root, result)
                raise ConversionServiceError(
                    "internal",
                    "document_fidelity_option_mismatch",
                    "DOCX producer artifacts do not match the accepted fidelity options",
                    details={
                        "image_resource_count": len(image_relations),
                        "ocr_fragment_count": len(ocr_relations),
                    },
                )
        if task.binding.bundle_profile == "physical_page_ocr":
            try:
                validate_physical_page_diagnostics(draft, result.diagnostics)
            except BundleMappingError as exc:
                self._discard_result_artifacts(task.staging_root, result)
                raise ConversionServiceError(
                    exc.category,
                    exc.code,
                    str(exc),
                    details=exc.details,
                ) from exc
            preferred = [artifact for artifact in result.artifacts if artifact.is_primary]
            expected_ocr = task.public_options.get("recognize_text")
            expected_images = task.public_options.get("preserve_resources")
            if (
                len(preferred) != 1
                or not isinstance(expected_ocr, bool)
                or not isinstance(expected_images, bool)
                or preferred[0].metadata.get("ocr_enabled") is not expected_ocr
                or preferred[0].metadata.get("keep_images") is not expected_images
            ):
                self._discard_result_artifacts(task.staging_root, result)
                raise ConversionServiceError(
                    "internal",
                    "physical_page_option_mismatch",
                    "physical-page producer modes do not match the accepted Machine options",
                )
        try:
            bundle = self._bundle_committer.commit(
                task_id=result.task_id,
                staging_root=task.staging_root,
                draft=draft,
            )
        except Exception as exc:
            logger.exception("Artifact Bundle commit rejected for task %s", result.task_id)
            self._discard_result_artifacts(task.staging_root, result)
            raise ConversionServiceError(
                "security",
                "bundle_commit_failed",
                "runtime rejected the output bundle",
            ) from exc
        return ConversionTaskOutcome(
            task_id=result.task_id,
            state="completed",
            bundle=bundle,
            diagnostics=tuple(result.diagnostics),
            metrics=result.metrics,
        )

    def _discard_result_artifacts(self, staging_root: str, result: ConversionResult) -> None:
        paths = [artifact.staging_path for artifact in result.artifacts]
        if not paths:
            return
        try:
            self._bundle_committer.discard(staging_root=staging_root, artifact_paths=paths)
        except Exception:
            # Preserve the authoritative task failure. The machine adapter never
            # projects rejected paths as deliverables even if local cleanup is denied.
            return

    def _validate_plan_request(self, request: ConversionPlanRequest) -> tuple[CapabilityBinding, dict[str, Any]]:
        discovery = self._discover_runtime()
        binding = next((item for item in discovery.bindings if item.capability_id == request.capability_id), None)
        if binding is None:
            raise ConversionServiceError(
                "unsupported",
                "capability_not_found",
                f"unsupported capability: {request.capability_id}",
            )
        if not self._controller.has_runtime:
            raise ConversionServiceError("unavailable", "runtime_unavailable", "conversion runtime is unavailable")
        runtime_state = self._runtime_capability_state(binding, discovery)
        if runtime_state.availability == "unavailable":
            missing = [
                dependency["dependency_id"]
                for dependency in runtime_state.dependencies
                if dependency["required"] and not dependency["available"]
            ]
            raise ConversionServiceError(
                "unavailable",
                "capability_unavailable",
                f"capability is unavailable in the active runtime: {binding.capability_id}",
                details={"missing_required_dependencies": missing},
            )
        input_ids = [handle.input_id for handle in request.inputs]
        if len(input_ids) != len(set(input_ids)):
            raise ConversionServiceError(
                "invalid_request",
                "duplicate_input_id",
                "task inputs must use unique input identifiers",
            )
        for handle in request.inputs:
            self._validate_input_logical_path(handle.logical_path)
        logical_paths = [handle.logical_path for handle in request.inputs]
        if len(logical_paths) != len(set(logical_paths)):
            raise ConversionServiceError(
                "invalid_request",
                "duplicate_input_logical_path",
                "task inputs must use unique logical paths",
            )
        slots = {slot.role: slot for slot in binding.input_shape.slots}
        for handle in request.inputs:
            slot = slots.get(handle.role)
            if slot is None:
                raise ConversionServiceError(
                    "invalid_request",
                    "undeclared_input_role",
                    f"capability does not declare input role: {handle.role}",
                )
        for handle in request.inputs:
            slot = slots[handle.role]
            if handle.kind != slot.kind:
                raise ConversionServiceError(
                    "invalid_request",
                    "input_slot_kind_mismatch",
                    f"input role {handle.role} requires kind {slot.kind}",
                )
        for handle in request.inputs:
            slot = slots[handle.role]
            if handle.media_type not in slot.media_types:
                raise ConversionServiceError(
                    "unsupported",
                    "input_slot_media_type_mismatch",
                    f"input role {handle.role} does not accept {handle.media_type}",
                )
        if binding.capability_id == MARKDOWN_TO_DOCX_CAPABILITY_ID:
            neutral_count = sum(handle.role == "neutral_document" for handle in request.inputs)
            plan_count = sum(handle.role == "numbering_export_plan" for handle in request.inputs)
            if neutral_count == 0:
                raise ConversionServiceError(
                    "invalid_request",
                    "docwen.resolved_document.missing",
                    "resolved-document input is required",
                )
            if plan_count == 0:
                raise ConversionServiceError(
                    "invalid_request",
                    "docwen.numbering_export_plan.missing",
                    "numbering-export-plan input is required",
                )
            if neutral_count != 1:
                raise ConversionServiceError(
                    "invalid_request",
                    "docwen.resolved_document.invalid",
                    "resolved-document input must occur exactly once",
                )
            if plan_count != 1:
                raise ConversionServiceError(
                    "invalid_request",
                    "docwen.numbering_export_plan.invalid",
                    "numbering-export-plan input must occur exactly once",
                )
        for slot in binding.input_shape.slots:
            count = sum(handle.role == slot.role for handle in request.inputs)
            if count < slot.min_items or (slot.max_items is not None and count > slot.max_items):
                raise ConversionServiceError(
                    "invalid_request",
                    "input_slot_cardinality_mismatch",
                    f"input role {slot.role} has invalid cardinality",
                )
        effective_options = resolve_conversion_options(request.options, binding)
        if request.output.staging_policy != "require_empty":
            raise ConversionServiceError(
                "invalid_request",
                "unsupported_staging_policy",
                "staging_policy must be require_empty",
            )
        for handle in request.inputs:
            self._validate_input(handle)
        if binding.capability_id == MARKDOWN_TO_DOCX_CAPABILITY_ID:
            self._validate_resolved_numbering_inputs(request)
        self._validate_empty_staging_root(request.output.staging_root)
        return binding, effective_options

    @staticmethod
    def _validate_resolved_numbering_inputs(request: ConversionPlanRequest) -> None:
        neutral = next(handle for handle in request.inputs if handle.role == "neutral_document")
        plan = next(handle for handle in request.inputs if handle.role == "numbering_export_plan")
        try:
            neutral_bytes = Path(neutral.path).read_bytes()
        except OSError as exc:
            raise ConversionServiceError(
                "security",
                "docwen.resolved_document.invalid",
                "resolved-document input cannot be read",
            ) from exc
        try:
            plan_bytes = Path(plan.path).read_bytes()
        except OSError as exc:
            raise ConversionServiceError(
                "security",
                "docwen.numbering_export_plan.invalid",
                "numbering-export-plan input cannot be read",
            ) from exc
        if len(neutral_bytes) != neutral.size_bytes or hashlib.sha256(neutral_bytes).hexdigest() != neutral.sha256:
            raise ConversionServiceError(
                "conflict",
                "docwen.resolved_document.invalid",
                "resolved-document changed during admission",
            )
        if len(plan_bytes) != plan.size_bytes or hashlib.sha256(plan_bytes).hexdigest() != plan.sha256:
            raise ConversionServiceError(
                "conflict",
                "docwen.numbering_export_plan.invalid",
                "numbering-export-plan changed during admission",
            )
        try:
            load_resolved_numbering_bytes(neutral_bytes, plan_bytes)
        except ResolvedNumberingPortError as exc:
            category = (
                "unsupported"
                if exc.code == "docwen.numbering_export_plan.unsupported_materialization"
                else "invalid_request"
            )
            raise ConversionServiceError(category, exc.code, str(exc)) from exc

    def _discover_runtime(self) -> _RuntimeDiscovery:
        if not self._controller.has_runtime:
            return _RuntimeDiscovery(
                gates={},
                routes={},
                error={
                    "severity": "error",
                    "code": "runtime_unavailable",
                    "message": "The conversion runtime is not configured.",
                },
            )

        try:
            description = self._controller.describe_runtime_capabilities()
            optimizations = parse_optimization_catalog(description)
            gates = {
                str(gate.get("id")): bool(gate.get("available"))
                for gate in description.get("gates", [])
                if isinstance(gate, dict) and isinstance(gate.get("id"), str)
            }
            routes = {
                str(candidate["id"]): candidate
                for source in description.get("sources", [])
                if isinstance(source, dict)
                for candidate in source.get("routes", [])
                if isinstance(candidate, dict) and isinstance(candidate.get("id"), str)
            }
            bindings = composed_capability_bindings(optimizations, routes)
        except Exception:
            return _RuntimeDiscovery(
                gates={},
                routes={},
                error={
                    "severity": "error",
                    "code": "runtime_discovery_failed",
                    "message": "The active runtime could not describe its capabilities.",
                },
            )
        return _RuntimeDiscovery(gates=gates, routes=routes, bindings=bindings, catalog=optimizations.runtime_catalog)

    @staticmethod
    def _runtime_capability_state(
        binding: CapabilityBinding,
        discovery: _RuntimeDiscovery,
    ) -> _RuntimeCapabilityState:
        if discovery.error is not None:
            return _RuntimeCapabilityState(
                availability="unavailable",
                limitations=(discovery.error,),
            )

        plan = (
            resolve_conversion_route_plan(
                discovery.catalog,
                source_format=binding.input_format,
                source_category=binding.input_category,
                target_format=binding.target_format,
                action_name=binding.action_name,
            )
            if discovery.catalog is not None
            else None
        )
        if plan is None or plan.final_route.id != binding.runtime_route_id:
            return _RuntimeCapabilityState(
                availability="unavailable",
                limitations=(
                    {
                        "severity": "error",
                        "code": "runtime_route_missing",
                        "message": f"The active runtime does not expose the complete route for {binding.capability_id}.",
                    },
                ),
            )

        routes = tuple(discovery.routes[route.id] for route in plan.routes)

        required_ids = tuple(
            dict.fromkeys(
                (
                    *(str(item) for route in routes for item in route.get("required_capabilities", [])),
                    *binding.required_dependency_ids,
                )
            )
        )
        optional_ids = tuple(
            dict.fromkeys(
                str(item)
                for route in routes
                for item in route.get("optional_capabilities", [])
                if str(item) not in required_ids
            )
        )
        if binding.dependency_ids:
            required_ids = tuple(item for item in required_ids if item in binding.dependency_ids)
            optional_ids = tuple(item for item in optional_ids if item in binding.dependency_ids)
        dependencies = tuple(
            {
                "dependency_id": dependency_id,
                "required": required,
                "available": discovery.gates.get(dependency_id, False),
            }
            for required, dependency_ids in (
                (True, required_ids),
                (False, optional_ids),
            )
            for dependency_id in dependency_ids
        )
        route_limitations = (
            tuple(
                {
                    "severity": "warning",
                    "code": "runtime_route_limitation",
                    "message": str(message),
                }
                for route in routes
                for message in route.get("limitations", [])
                if str(message).strip()
            )
            if binding.project_runtime_limitations
            else ()
        )
        route_available = plan.available and not any(
            dependency["required"] and not dependency["available"] for dependency in dependencies
        )
        if not route_available:
            availability: Literal["available", "limited", "unavailable"] = "unavailable"
        elif any(not dependency["required"] and not dependency["available"] for dependency in dependencies):
            availability = "limited"
        else:
            availability = "available"
        return _RuntimeCapabilityState(
            availability=availability,
            dependencies=dependencies,
            limitations=route_limitations,
        )

    def _validate_execution_snapshot(self, request: ConversionRequest, staging_root: str) -> None:
        for input_ref in request.input_refs:
            expected_sha = str(input_ref.metadata["machine_input_sha256"])
            expected_size = int(input_ref.metadata["machine_input_size_bytes"])
            path = filesystem_path(input_ref.path, force_extended=sys.platform == "win32")
            if self._path_traverses_link_or_junction(path):
                raise ConversionServiceError(
                    "security",
                    "input_is_link",
                    "input must not be a link or junction",
                )
            size_bytes, sha256 = self._file_integrity(path, code_prefix="input")
            if size_bytes != expected_size or sha256 != expected_sha:
                raise ConversionServiceError(
                    "conflict",
                    "input_changed_after_plan",
                    "input content changed after planning",
                )
        self._validate_empty_staging_root(staging_root)

    def _validate_input(self, handle: LocalInputHandle) -> None:
        self._validate_identifier(handle.input_id, field_name="input_id")
        path = Path(handle.path)
        if not path.is_absolute():
            raise ConversionServiceError("invalid_request", "input_path_not_absolute", "input path must be absolute")
        io_path = filesystem_path(path, force_extended=sys.platform == "win32")
        if self._path_traverses_link_or_junction(io_path):
            raise ConversionServiceError(
                "security",
                "input_is_link",
                "input must not be a link or junction",
            )
        size_bytes, sha256 = self._file_integrity(io_path, code_prefix="input")
        if size_bytes != handle.size_bytes or sha256 != handle.sha256:
            raise ConversionServiceError(
                "invalid_request",
                "input_integrity_mismatch",
                "declared input size or sha256 does not match the local file",
            )

    @staticmethod
    def _validate_input_logical_path(logical_path: str) -> None:
        segments = logical_path.split("/")
        if (
            not logical_path
            or len(logical_path) > 1024
            or logical_path.startswith("/")
            or "\\" in logical_path
            or "\x00" in logical_path
            or ":" in segments[0]
            or any(segment in {"", ".", ".."} for segment in segments)
        ):
            raise ConversionServiceError(
                "invalid_request",
                "invalid_input_logical_path",
                "input logical_path must be a normalized relative POSIX path",
            )

    @classmethod
    def _validate_empty_staging_root(cls, staging_root: str) -> None:
        path = Path(staging_root)
        if not path.is_absolute():
            raise ConversionServiceError(
                "invalid_request",
                "staging_root_not_absolute",
                "staging root must be absolute",
            )
        io_path = filesystem_path(path, force_extended=sys.platform == "win32")
        if cls._is_link_or_junction(io_path):
            raise ConversionServiceError(
                "security",
                "staging_root_is_link",
                "staging root must not be a link or junction",
            )
        try:
            if not io_path.is_dir():
                raise ConversionServiceError(
                    "invalid_request",
                    "staging_root_not_directory",
                    "staging root must be an existing directory",
                )
            if next(io_path.iterdir(), None) is not None:
                raise ConversionServiceError(
                    "conflict",
                    "staging_root_not_empty",
                    "staging root must be empty",
                )
        except OSError as exc:
            raise ConversionServiceError(
                "security",
                "staging_root_unreadable",
                "staging root cannot be inspected",
            ) from exc

    @staticmethod
    def _validate_identifier(value: str, *, field_name: str) -> None:
        if not value or len(value) > 128 or not value[0].isalnum():
            raise ConversionServiceError(
                "invalid_request",
                "invalid_identifier",
                f"{field_name} is not a valid identifier",
            )
        if any(not (character.isascii() and (character.isalnum() or character in "._:-")) for character in value):
            raise ConversionServiceError(
                "invalid_request",
                "invalid_identifier",
                f"{field_name} is not a valid identifier",
            )

    @staticmethod
    def _is_link_or_junction(path: Path) -> bool:
        if path.is_symlink():
            return True
        is_junction = getattr(path, "is_junction", None)
        return bool(is_junction()) if callable(is_junction) else False

    @classmethod
    def _path_traverses_link_or_junction(cls, path: Path) -> bool:
        current = path
        while True:
            if cls._is_link_or_junction(current):
                return True
            parent = current.parent
            if parent == current:
                return False
            current = parent

    @staticmethod
    def _file_integrity(path: Path, *, code_prefix: str) -> tuple[int, str]:
        if not path.is_file():
            raise ConversionServiceError(
                "invalid_request",
                f"{code_prefix}_not_regular_file",
                f"{code_prefix} must be an existing regular file",
            )
        digest = hashlib.sha256()
        size_bytes = 0
        try:
            with path.open("rb") as stream:
                while chunk := stream.read(_HASH_CHUNK_BYTES):
                    size_bytes += len(chunk)
                    digest.update(chunk)
        except OSError as exc:
            raise ConversionServiceError(
                "security",
                f"{code_prefix}_unreadable",
                f"{code_prefix} cannot be read",
            ) from exc
        return size_bytes, digest.hexdigest()


__all__ = ["ConversionService"]

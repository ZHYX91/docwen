"""Runtime outcome and terminal-event honesty contracts."""

from __future__ import annotations

import hashlib
from collections.abc import Callable
from pathlib import Path
from typing import Any

import pytest

from docwen_core.models.artifact import ARTIFACT_KIND_PRIMARY, ArtifactManifest
from docwen_core.models.file_ref import FileRef
from docwen_core.models.manifest import PluginManifest, RouteSpec
from docwen_core.models.request import ConversionRequest, OutputPolicy
from docwen_core.models.resolved_numbering import (
    NUMBERING_EXPORT_PLAN_MEDIA_TYPE,
    RESOLVED_DOCUMENT_MEDIA_TYPE,
)
from docwen_core.models.result import (
    ConversionDiagnostic,
    ConversionErrorInfo,
    ConversionMetrics,
    ConversionResult,
)
from docwen_core.models.task import TaskEvent
from docwen_runtime.engine.route_resolver import RouteResolver
from docwen_runtime.engine.task_manager import TaskManager
from docwen_runtime.output.finalizer import OutputFinalizer
from docwen_runtime.plugin_registry.registry import PluginRegistry
from docwen_runtime.workspace.manager import WorkspaceManager

pytestmark = pytest.mark.unit

ConvertFn = Callable[[Any], ConversionResult]


class _OutcomePlugin:
    def __init__(self, convert_fn: ConvertFn) -> None:
        self._convert_fn = convert_fn
        self._manifest = PluginManifest(
            plugin_id="outcome_honesty",
            name="Outcome honesty probe",
            version="0.1.0",
            description="Deterministic Runtime outcome probe",
            routes=[
                RouteSpec(source_format="markdown", target_format="docx", label="probe"),
                RouteSpec(
                    source_format="markdown",
                    target_format="markdown",
                    action_name="validate",
                    label="proofread probe",
                ),
            ],
        )

    @property
    def manifest(self) -> PluginManifest:
        return self._manifest

    def can_handle(self, source_format: str, target_format: str, action_name: str = "") -> bool:
        return source_format == "markdown" and (
            (target_format == "docx" and not action_name) or (target_format == "markdown" and action_name == "validate")
        )

    def convert(self, context: Any) -> ConversionResult:
        return self._convert_fn(context)


def _build_manager(
    tmp_path: Path,
    convert_fn: ConvertFn,
    *,
    finalizer: OutputFinalizer | None = None,
) -> TaskManager:
    registry = PluginRegistry()
    registry.register(_OutcomePlugin(convert_fn))
    return TaskManager(
        registry,
        RouteResolver(registry),
        WorkspaceManager(root_dir=str(tmp_path / "runtime")),
        finalizer or OutputFinalizer(),
    )


def _request(
    tmp_path: Path,
    task_id: str,
    *,
    target_format: str = "docx",
    action_name: str = "",
) -> ConversionRequest:
    input_path = tmp_path / f"{task_id}.md"
    input_path.write_text("# input", encoding="utf-8")
    return ConversionRequest(
        request_id=task_id,
        input_refs=[
            FileRef(
                path=str(input_path),
                format="markdown",
                category="markdown",
                size_bytes=input_path.stat().st_size,
            )
        ],
        target_format=target_format,
        action_name=action_name,
        output_policy=OutputPolicy(output_dir=str(tmp_path / "output")),
    )


def _terminal_events(events: list[TaskEvent]) -> list[TaskEvent]:
    return [event for event in events if event.event_type in {"task_completed", "task_failed", "task_cancelled"}]


def _successful_artifact(context: Any, *, content: str = "placed") -> ArtifactManifest:
    staging_path = context.workspace.create_artifact_path("primary", ".docx")
    Path(staging_path).write_text(content, encoding="utf-8")
    return ArtifactManifest(
        artifact_id="primary",
        kind=ARTIFACT_KIND_PRIMARY,
        staging_path=staging_path,
        suggested_name="result.docx",
        is_primary=True,
    )


def test_unsupported_route_is_typed_and_fails_before_plugin_execution(tmp_path: Path) -> None:
    plugin_called = False

    def convert(context: Any) -> ConversionResult:
        nonlocal plugin_called
        plugin_called = True
        return ConversionResult(task_id=context.request.request_id, success=True)

    manager = _build_manager(tmp_path, convert)
    events: list[TaskEvent] = []

    result = manager.execute_single(
        _request(tmp_path, "unsupported-route", target_format="pdf"),
        on_event=events.append,
    )

    assert result.success is False
    assert result.error is not None
    assert result.error.error_type == "unsupported_route"
    assert result.error.diagnostic_code == "ROUTE_UNSUPPORTED"
    assert plugin_called is False
    assert [event.event_type for event in events] == ["task_started", "task_failed"]
    assert events[-1].payload["error_type"] == "unsupported_route"


def test_resolved_numbering_uses_neutral_document_as_runtime_primary(tmp_path: Path) -> None:
    neutral = tmp_path / "resolved-document.json"
    plan = tmp_path / "numbering-export-plan.json"
    neutral.write_bytes(b'{"authored_markdown":"# Authored"}')
    plan.write_bytes(b'{"targets":[]}')

    def ref(path: Path, *, kind: str, role: str, media_type: str) -> FileRef:
        data = path.read_bytes()
        return FileRef(
            path=str(path),
            format="markdown" if kind == "document" else "resource",
            category="markdown" if kind == "document" else "other",
            size_bytes=len(data),
            input_kind=kind,
            input_role=role,
            logical_path=f"request/{role}.json",
            media_type=media_type,
            metadata={
                "machine_input_size_bytes": len(data),
                "machine_input_sha256": hashlib.sha256(data).hexdigest(),
            },
        )

    neutral_ref = ref(
        neutral,
        kind="document",
        role="neutral_document",
        media_type=RESOLVED_DOCUMENT_MEDIA_TYPE,
    )
    plan_ref = ref(
        plan,
        kind="resource",
        role="numbering_export_plan",
        media_type=NUMBERING_EXPORT_PLAN_MEDIA_TYPE,
    )
    observed: dict[str, Any] = {}

    def convert(context: Any) -> ConversionResult:
        observed["input_path"] = context.workspace.input_path
        observed["roles"] = [item.input_role for item in context.workspace.input_resources()]
        neutral_copy = context.workspace.input_resources("neutral_document")[0]
        observed["neutral_path"] = neutral_copy.path
        observed["neutral_bytes"] = Path(neutral_copy.path).read_bytes()
        observed["plan_bytes"] = Path(context.workspace.input_resources("numbering_export_plan")[0].path).read_bytes()
        return ConversionResult(
            task_id=context.request.request_id,
            success=True,
            artifacts=[_successful_artifact(context)],
        )

    manager = _build_manager(tmp_path, convert)
    events: list[TaskEvent] = []
    request = ConversionRequest(
        request_id="resolved-numbering-runtime",
        input_refs=[neutral_ref, plan_ref],
        target_format="docx",
        output_policy=OutputPolicy(output_dir=str(tmp_path / "output")),
    )

    result = manager.execute_single(request, on_event=events.append)

    assert result.success is True
    assert result.metrics.input_bytes == neutral_ref.size_bytes
    assert events[0].event_type == "task_started"
    assert events[0].payload["input_path"] == str(neutral)
    assert observed["input_path"] == observed["neutral_path"]
    assert observed["roles"] == ["neutral_document", "numbering_export_plan"]
    assert observed["neutral_bytes"] == neutral.read_bytes()
    assert observed["plan_bytes"] == plan.read_bytes()


def test_reported_plugin_failure_emits_failed_without_spurious_finalizing_progress(tmp_path: Path) -> None:
    error = ConversionErrorInfo(
        error_type="conversion_failed",
        message="plugin rejected input",
        diagnostic_code="PLUGIN-FAILED",
    )

    def convert(context: Any) -> ConversionResult:
        context.progress.report_diagnostic("error", error.message, error.diagnostic_code)
        return ConversionResult(
            task_id=context.request.request_id,
            success=False,
            diagnostics=[ConversionDiagnostic(level="error", message=error.message, code=error.diagnostic_code)],
            error=error,
            metrics=ConversionMetrics(extra={"plugin_stage": "validation"}),
        )

    manager = _build_manager(tmp_path, convert)
    request = _request(tmp_path, "reported-failure")
    events: list[TaskEvent] = []

    result = manager.execute_single(request, on_event=events.append)

    assert result.success is False
    assert result.error == error
    assert result.metrics.input_bytes == request.input_refs[0].size_bytes
    assert result.metrics.extra == {"plugin_stage": "validation"}
    assert [event.event_type for event in _terminal_events(events)] == ["task_failed"]
    assert not any(
        event.event_type == "task_progress" and event.payload.get("message") == "Finalizing output" for event in events
    )


def test_finalizer_partial_failure_returns_typed_error_and_failed_terminal(tmp_path: Path) -> None:
    def convert(context: Any) -> ConversionResult:
        good = _successful_artifact(context, content="good")
        missing = ArtifactManifest(
            artifact_id="missing",
            kind=ARTIFACT_KIND_PRIMARY,
            staging_path=context.workspace.create_artifact_path("primary", ".missing"),
            suggested_name="missing.docx",
            is_primary=True,
        )
        return ConversionResult(
            task_id=context.request.request_id,
            success=True,
            artifacts=[good, missing],
            diagnostics=[ConversionDiagnostic(level="info", message="plugin done", code="PLUGIN-DONE")],
            metrics=ConversionMetrics(extra={"plugin_metric": "kept"}),
        )

    manager = _build_manager(tmp_path, convert)
    request = _request(tmp_path, "partial-finalizer")
    events: list[TaskEvent] = []

    result = manager.execute_single(request, on_event=events.append)

    assert result.success is False
    assert result.error is not None
    assert result.error.error_type == "output_finalization_failed"
    assert result.error.diagnostic_code == "FINALIZER_PARTIAL"
    assert [artifact.suggested_name for artifact in result.artifacts] == ["result.docx"]
    assert Path(result.artifacts[0].staging_path).read_text(encoding="utf-8") == "good"
    assert result.metrics.output_bytes == 4
    assert result.metrics.extra["plugin_metric"] == "kept"
    assert [event.event_type for event in _terminal_events(events)] == ["task_failed"]
    assert _terminal_events(events)[0].payload["error_type"] == "output_finalization_failed"


def test_intentional_no_output_success_bypasses_finalizer_and_completes(
    tmp_path: Path,
    monkeypatch,
) -> None:
    def convert(context: Any) -> ConversionResult:
        diagnostic = ConversionDiagnostic(
            level="info",
            message="All checks are disabled",
            code="PROOFREAD-SKIPPED",
        )
        context.progress.report_diagnostic(diagnostic.level, diagnostic.message, diagnostic.code)
        return ConversionResult(
            task_id=context.request.request_id,
            success=True,
            diagnostics=[diagnostic],
            metrics=ConversionMetrics(output_bytes=999, extra={"checks": 0}),
        )

    finalizer = OutputFinalizer()

    def reject_finalize(**_kwargs: Any) -> ConversionResult:
        raise AssertionError("artifact-free success must bypass OutputFinalizer")

    monkeypatch.setattr(finalizer, "finalize", reject_finalize)
    manager = _build_manager(tmp_path, convert, finalizer=finalizer)
    request = _request(
        tmp_path,
        "no-output",
        target_format="markdown",
        action_name="validate",
    )
    events: list[TaskEvent] = []

    result = manager.execute_single(request, on_event=events.append)

    assert result.success is True
    assert result.error is None
    assert result.artifacts == []
    assert [diagnostic.code for diagnostic in result.diagnostics] == ["PROOFREAD-SKIPPED"]
    assert result.metrics.input_bytes == request.input_refs[0].size_bytes
    assert result.metrics.output_bytes == 0
    assert result.metrics.extra == {"checks": 0}
    assert [event.event_type for event in _terminal_events(events)] == ["task_completed"]
    assert not any(
        event.event_type == "task_progress" and event.payload.get("message") == "Finalizing output" for event in events
    )


def test_real_proofread_disabled_result_is_an_intentional_empty_report_success(tmp_path: Path) -> None:
    import json

    from docwen_plugin_proofread import ProofreadPlugin

    registry = PluginRegistry()
    registry.register(ProofreadPlugin())
    manager = TaskManager(
        registry,
        RouteResolver(registry),
        WorkspaceManager(root_dir=str(tmp_path / "proofread-runtime")),
        OutputFinalizer(),
    )
    input_path = tmp_path / "proofread.md"
    input_path.write_text("No checks should run.", encoding="utf-8")
    request = ConversionRequest(
        request_id="proofread-no-output",
        input_refs=[
            FileRef(
                path=str(input_path),
                format="markdown",
                category="markdown",
                size_bytes=input_path.stat().st_size,
            )
        ],
        target_format="markdown",
        action_name="validate",
        options={
            "enable_symbol_pairing": False,
            "enable_symbol_correction": False,
            "enable_typos_rule": False,
            "enable_sensitive_word": False,
        },
        output_policy=OutputPolicy(output_dir=str(tmp_path / "proofread-output")),
    )
    events: list[TaskEvent] = []

    result = manager.execute_single(request, on_event=events.append)

    assert result.success is True
    assert result.error is None
    assert len(result.artifacts) == 1
    report_path = Path(result.artifacts[0].staging_path)
    report = json.loads(report_path.read_text(encoding="utf-8"))
    assert report["schema"] == "docwen.proofread_report.v2"
    assert report["issues"] == []
    assert report["summary"] == {}
    assert not any(report["checks_enabled"].values())
    assert [diagnostic.code for diagnostic in result.diagnostics] == ["PROOFREAD-SKIPPED", "FINALIZER_DONE"]
    assert result.metrics.input_bytes == input_path.stat().st_size
    assert result.metrics.output_bytes == report_path.stat().st_size
    assert result.metrics.output_bytes > 0
    assert [event.event_type for event in _terminal_events(events)] == ["task_completed"]
    assert any(
        event.event_type == "task_progress" and event.payload.get("message") == "Finalizing output" for event in events
    )


def test_ordinary_success_without_artifacts_is_a_typed_finalizer_failure(tmp_path: Path) -> None:
    def convert(context: Any) -> ConversionResult:
        return ConversionResult(task_id=context.request.request_id, success=True)

    manager = _build_manager(tmp_path, convert)
    request = _request(tmp_path, "ordinary-empty")
    events: list[TaskEvent] = []

    result = manager.execute_single(request, on_event=events.append)

    assert result.success is False
    assert result.error is not None
    assert result.error.error_type == "output_finalization_failed"
    assert result.error.diagnostic_code == "FINALIZER_NO_ARTIFACTS"
    assert [event.event_type for event in _terminal_events(events)] == ["task_failed"]
    assert any(
        event.event_type == "task_progress" and event.payload.get("message") == "Finalizing output" for event in events
    )


def test_plugin_success_with_non_cancel_error_is_normalized_to_failed(tmp_path: Path) -> None:
    error = ConversionErrorInfo(
        error_type="conversion_failed",
        message="plugin returned contradictory state",
        diagnostic_code="PLUGIN-CONTRADICTION",
    )

    def convert(context: Any) -> ConversionResult:
        return ConversionResult(
            task_id=context.request.request_id,
            success=True,
            artifacts=[_successful_artifact(context)],
            error=error,
        )

    manager = _build_manager(tmp_path, convert)
    request = _request(tmp_path, "plugin-contradiction")
    events: list[TaskEvent] = []

    result = manager.execute_single(request, on_event=events.append)

    assert result.success is False
    assert result.error == error
    assert result.artifacts == []
    assert [event.event_type for event in _terminal_events(events)] == ["task_failed"]
    assert not any(
        event.event_type == "task_progress" and event.payload.get("message") == "Finalizing output" for event in events
    )


def test_finalizer_exception_preserves_runtime_metrics_and_plugin_extras(
    tmp_path: Path,
    monkeypatch,
) -> None:
    def convert(context: Any) -> ConversionResult:
        return ConversionResult(
            task_id=context.request.request_id,
            success=True,
            artifacts=[_successful_artifact(context)],
            diagnostics=[ConversionDiagnostic(level="warning", message="semantic warning", code="PLUGIN-WARN")],
            metrics=ConversionMetrics(extra={"plugin_metric": "kept"}),
        )

    finalizer = OutputFinalizer()

    def explode(**_kwargs: Any) -> ConversionResult:
        raise OSError("destination unavailable")

    ticks = iter([100.0, 100.25])
    monkeypatch.setattr(finalizer, "finalize", explode)
    monkeypatch.setattr("docwen_runtime.engine.task_manager.time.perf_counter", lambda: next(ticks))
    manager = _build_manager(tmp_path, convert, finalizer=finalizer)
    request = _request(tmp_path, "finalizer-exception")
    events: list[TaskEvent] = []

    result = manager.execute_single(request, on_event=events.append)

    assert result.success is False
    assert result.error is not None
    assert result.error.error_type == "conversion_failed"
    assert result.error.message == "destination unavailable"
    assert [diagnostic.code for diagnostic in result.diagnostics] == ["PLUGIN-WARN"]
    assert result.metrics.duration_ms == 250.0
    assert result.metrics.input_bytes == request.input_refs[0].size_bytes
    assert result.metrics.output_bytes == 0
    assert result.metrics.extra == {"plugin_metric": "kept"}
    assert [event.event_type for event in _terminal_events(events)] == ["task_failed"]
    assert _terminal_events(events)[0].payload["error_type"] == result.error.error_type


def test_terminal_listener_rejection_cannot_change_success_or_duplicate_terminal(tmp_path: Path) -> None:
    def convert(context: Any) -> ConversionResult:
        return ConversionResult(
            task_id=context.request.request_id,
            success=True,
            artifacts=[_successful_artifact(context)],
            metrics=ConversionMetrics(extra={"plugin_metric": "kept"}),
        )

    manager = _build_manager(tmp_path, convert)
    request = _request(tmp_path, "listener-rejection")
    events: list[TaskEvent] = []

    def rejecting_listener(event: TaskEvent) -> None:
        events.append(event)
        if event.event_type == "task_completed":
            raise RuntimeError("presentation callback failed")

    result = manager.execute_single(request, on_event=rejecting_listener)

    assert result.success is True
    assert result.error is None
    assert len(result.artifacts) == 1
    assert Path(result.artifacts[0].staging_path).is_file()
    assert result.metrics.extra["plugin_metric"] == "kept"
    assert "TASK_EVENT_LISTENER_ERROR" in [diagnostic.code for diagnostic in result.diagnostics]
    assert [event.event_type for event in _terminal_events(events)] == ["task_completed"]


def test_dependency_egress_failure_keeps_typed_security_semantics(tmp_path: Path) -> None:
    from docwen_runtime.security import NetworkAccessBlockedError

    def convert(_context: Any) -> ConversionResult:
        raise NetworkAccessBlockedError("socket.getaddrinfo")

    manager = _build_manager(tmp_path, convert)
    request = _request(tmp_path, "network-blocked")
    events: list[TaskEvent] = []

    result = manager.execute_single(request, on_event=events.append)

    assert result.success is False
    assert result.error is not None
    assert result.error.error_type == "network_access_blocked"
    assert result.error.diagnostic_code == "NETWORK_ACCESS_BLOCKED"
    assert [diagnostic.code for diagnostic in result.diagnostics] == ["NETWORK_ACCESS_BLOCKED"]
    assert _terminal_events(events)[0].payload["error_type"] == "network_access_blocked"


def test_diagnostic_deduplication_keeps_distinct_artifact_bindings() -> None:
    first = ConversionDiagnostic(
        level="warning",
        message="page unknown",
        code="resource_page_unresolved",
        artifact_id="resource.one",
    )
    second = ConversionDiagnostic(
        level="warning",
        message="page unknown",
        code="resource_page_unresolved",
        artifact_id="resource.two",
    )

    merged = TaskManager._merge_diagnostics([first], [second], [first])

    assert [diagnostic.artifact_id for diagnostic in merged] == ["resource.one", "resource.two"]

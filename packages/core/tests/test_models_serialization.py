"""Contract tests: all core models must round-trip through to_dict/from_dict."""

from __future__ import annotations

from typing import Any

import pytest

from docwen_core.models.artifact import (
    ArtifactManifest,
)
from docwen_core.models.conversion_manifest import (
    ConversionManifestContext,
    OutputManifestPolicy,
    PreconversionStep,
)
from docwen_core.models.file_ref import FileRef
from docwen_core.models.manifest import OptimizationResourceSpec, PluginManifest, RouteCapabilityRule, RouteSpec
from docwen_core.models.request import ConversionRequest, OutputPolicy
from docwen_core.models.result import (
    ConversionDiagnostic,
    ConversionErrorInfo,
    ConversionMetrics,
    ConversionResult,
)
from docwen_core.models.task import TaskEvent
from docwen_core.models.worker import WorkerRequest, WorkerResult

pytestmark = pytest.mark.contract


# ── FileRef ──────────────────────────────────────────────────────────


class TestFileRefSerialization:
    def test_round_trip_minimal(self) -> None:
        ref = FileRef(path="/tmp/test.md", format="markdown", category="markdown")
        data = ref.to_dict()
        ref2 = FileRef.from_dict(data)
        assert ref2.path == "/tmp/test.md"
        assert ref2.format == "markdown"
        assert ref2.category == "markdown"

    def test_round_trip_full(self) -> None:
        ref = FileRef(
            path="/tmp/test.md",
            format="markdown",
            category="markdown",
            encoding="utf-16",
            size_bytes=2048,
            media_type="text/markdown",
            metadata={"title": "Hello"},
        )
        data = ref.to_dict()
        ref2 = FileRef.from_dict(data)
        assert ref2.encoding == "utf-16"
        assert ref2.size_bytes == 2048
        assert ref2.media_type == "text/markdown"
        assert ref2.metadata == {"title": "Hello"}

    def test_old_payload_defaults_media_type(self) -> None:
        ref = FileRef.from_dict({"path": "/tmp/test.md", "format": "markdown", "category": "markdown"})

        assert ref.media_type == ""

    def test_defaults(self) -> None:
        ref = FileRef(path="/a/b.txt", format="txt", category="document")
        assert ref.encoding == "utf-8"
        assert ref.size_bytes == 0
        assert ref.metadata == {}


# ── OutputPolicy ─────────────────────────────────────────────────────


class TestOutputPolicySerialization:
    def test_round_trip_defaults(self) -> None:
        pol = OutputPolicy()
        data = pol.to_dict()
        pol2 = OutputPolicy.from_dict(data)
        assert pol2.output_dir is None
        assert pol2.output_path is None
        assert pol2.date_subfolder == ""
        assert pol2.overwrite_mode == "rename"
        assert pol2.write_artifacts is True
        assert pol2.open_after_done is False

    def test_round_trip_custom(self) -> None:
        pol = OutputPolicy(
            output_dir="/tmp/out",
            date_subfolder="iso",
            overwrite_mode="skip",
            write_artifacts=False,
            open_after_done=True,
        )
        data = pol.to_dict()
        pol2 = OutputPolicy.from_dict(data)
        assert pol2.output_dir == "/tmp/out"
        assert pol2.date_subfolder == "iso"
        assert pol2.overwrite_mode == "skip"
        assert pol2.write_artifacts is False
        assert pol2.open_after_done is True

    def test_round_trip_exact_output_path(self) -> None:
        pol = OutputPolicy(output_path="/tmp/report.docx", overwrite_mode="error")

        pol2 = OutputPolicy.from_dict(pol.to_dict())

        assert pol2.output_path == "/tmp/report.docx"
        assert pol2.output_dir is None
        assert pol2.overwrite_mode == "error"


class TestOutputManifestContextSerialization:
    @pytest.mark.parametrize(
        ("raw", "expected"),
        [
            ({}, OutputManifestPolicy(False, True)),
            (
                {"output": {"manifest": {"save_to_output": True, "mask_input_path": False}}},
                OutputManifestPolicy(True, False),
            ),
            (
                {"output": {"manifest": {"save_to_output": 1, "mask_input_path": "false"}}},
                OutputManifestPolicy(False, True),
            ),
        ],
    )
    def test_policy_projection_is_strict_and_privacy_fails_closed(
        self, raw: dict[str, Any], expected: OutputManifestPolicy
    ) -> None:
        assert OutputManifestPolicy.from_config_snapshot(raw) == expected

    def test_context_round_trip_and_batch_projection(self) -> None:
        refs = [
            FileRef(path="/private/a.rtf", format="rtf", category="document"),
            FileRef(path="/private/b.csv", format="csv", category="spreadsheet"),
        ]
        context = ConversionManifestContext.from_request_inputs(
            refs,
            {"output": {"manifest": {"save_to_output": True}}},
        ).with_step(
            PreconversionStep(
                input_index=0,
                source_format="rtf",
                target_format="docx",
                status="completed",
                backend="LibreOffice",
            )
        )

        restored = ConversionManifestContext.from_dict(context.to_dict())
        child = restored.for_input(0)

        assert restored == context
        assert child.batch_child is True
        assert [item.path for item in child.inputs] == ["/private/a.rtf"]
        assert child.preconversion_steps[0].input_index == 0

    def test_context_projects_neutral_document_not_numbering_plan(self) -> None:
        refs = [
            FileRef(
                path="/private/resolved-document.json",
                format="markdown",
                category="markdown",
                input_kind="document",
                input_role="neutral_document",
            ),
            FileRef(
                path="/private/numbering-export-plan.json",
                format="resource",
                category="other",
                input_kind="resource",
                input_role="numbering_export_plan",
            ),
        ]

        context = ConversionManifestContext.from_request_inputs(refs, {})

        assert [item.path for item in context.inputs] == ["/private/resolved-document.json"]
        assert context.inputs[0].format == "markdown"


# ── ConversionRequest ────────────────────────────────────────────────


class TestConversionRequestSerialization:
    def test_round_trip_single_file(self) -> None:
        req = ConversionRequest(
            request_id="req-001",
            input_refs=[FileRef(path="/tmp/a.md", format="markdown", category="markdown")],
            target_format="docx",
            options={"remove_numbering": True},
        )
        data = req.to_dict()
        req2 = ConversionRequest.from_dict(data)
        assert req2.request_id == "req-001"
        assert req2.target_format == "docx"
        assert len(req2.input_refs) == 1
        assert req2.input_refs[0].path == "/tmp/a.md"
        assert req2.options["remove_numbering"] is True

    def test_round_trip_with_policy(self) -> None:
        req = ConversionRequest(
            request_id="req-002",
            input_refs=[],
            target_format="pdf",
            output_policy=OutputPolicy(output_dir="/out"),
        )
        data = req.to_dict()
        req2 = ConversionRequest.from_dict(data)
        assert req2.output_policy.output_dir == "/out"

    def test_round_trip_with_manifest_context(self) -> None:
        context = ConversionManifestContext.from_request_inputs(
            [FileRef(path="/private/a.md", format="markdown", category="markdown")],
            {"output": {"manifest": {"save_to_output": True}}},
        )
        req = ConversionRequest(
            request_id="req-manifest",
            input_refs=[FileRef(path="/private/a.md", format="markdown", category="markdown")],
            target_format="docx",
            manifest_context=context,
        )

        restored = ConversionRequest.from_dict(req.to_dict())

        assert restored.manifest_context == context

    def test_defaults(self) -> None:
        req = ConversionRequest(request_id="r", input_refs=[], target_format="md")
        assert req.action_name == ""
        assert req.options == {}
        assert req.output_policy.output_dir is None
        assert req.manifest_context is None


# ── ArtifactManifest ─────────────────────────────────────────────────


class TestArtifactManifestSerialization:
    def test_round_trip_primary(self) -> None:
        art = ArtifactManifest(
            artifact_id="art-1",
            kind="primary",
            staging_path="/tmp/staging/output.docx",
            suggested_name="output.docx",
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            is_primary=True,
        )
        data = art.to_dict()
        art2 = ArtifactManifest.from_dict(data)
        assert art2.artifact_id == "art-1"
        assert art2.kind == "primary"
        assert art2.suggested_name == "output.docx"
        assert art2.is_primary is True

    def test_defaults(self) -> None:
        art = ArtifactManifest(artifact_id="a", kind="log", staging_path="/tmp/log.txt", suggested_name="log.txt")
        assert art.media_type == "application/octet-stream"
        assert art.is_primary is False
        assert art.metadata == {}


# ── ConversionResult ─────────────────────────────────────────────────


class TestConversionResultSerialization:
    def test_round_trip_success(self) -> None:
        result = ConversionResult(
            task_id="task-1",
            success=True,
            artifacts=[
                ArtifactManifest(artifact_id="a1", kind="primary", staging_path="/s/o.docx", suggested_name="o.docx")
            ],
            diagnostics=[
                ConversionDiagnostic(level="info", message="OK"),
            ],
            metrics=ConversionMetrics(duration_ms=150.0, input_bytes=1024, output_bytes=512),
        )
        data = result.to_dict()
        result2 = ConversionResult.from_dict(data)
        assert result2.success is True
        assert len(result2.artifacts) == 1
        assert result2.metrics.duration_ms == 150.0
        assert result2.error is None

    def test_diagnostic_artifact_binding_round_trips_and_none_is_omitted(self) -> None:
        bound = ConversionDiagnostic(
            level="warning",
            message="page unknown",
            code="resource_page_unresolved",
            artifact_id="resource.image.5",
        )
        assert ConversionDiagnostic.from_dict(bound.to_dict()).artifact_id == "resource.image.5"
        assert "artifact_id" not in ConversionDiagnostic(level="info", message="ok").to_dict()

    def test_round_trip_failure(self) -> None:
        err = ConversionErrorInfo(
            error_type="conversion_failed",
            message="Something broke",
            traceback_text="Traceback...",
            recoverable=True,
            diagnostic_code="ERR-001",
        )
        result = ConversionResult(task_id="t2", success=False, error=err)
        data = result.to_dict()
        result2 = ConversionResult.from_dict(data)
        assert result2.success is False
        assert result2.error is not None
        assert result2.error.error_type == "conversion_failed"
        assert result2.error.recoverable is True
        assert result2.error.diagnostic_code == "ERR-001"

    def test_defaults(self) -> None:
        result = ConversionResult(task_id="t", success=True)
        assert result.artifacts == []
        assert result.diagnostics == []
        assert result.error is None
        assert result.metrics.duration_ms == 0.0


# ── TaskEvent ────────────────────────────────────────────────────────


class TestTaskEventSerialization:
    def test_round_trip(self) -> None:
        evt = TaskEvent(
            task_id="task-1",
            event_type="task_progress",
            sequence=3,
            timestamp="2026-06-05T12:00:00+00:00",
            payload={"percent": 50.0},
        )
        data = evt.to_dict()
        evt2 = TaskEvent.from_dict(data)
        assert evt2.task_id == "task-1"
        assert evt2.event_type == "task_progress"
        assert evt2.sequence == 3
        assert evt2.payload["percent"] == 50.0

    def test_auto_timestamp(self) -> None:
        evt = TaskEvent(task_id="t", event_type="task_started", sequence=0)
        assert evt.timestamp  # auto-filled in __post_init__


# ── PluginManifest & RouteSpec ───────────────────────────────────────


class TestPluginManifestSerialization:
    def test_round_trip(self) -> None:
        manifest = PluginManifest(
            plugin_id="docwen_plugin_markdown",
            name="Markdown Plugin",
            version="0.1.0",
            description="Converts markdown to/from other formats",
            author="DocWen",
            routes=[
                RouteSpec(
                    source_format="markdown",
                    target_format="docx",
                    action_name="",
                    label="Markdown → DOCX",
                    options_schema={"type": "object"},
                ),
                RouteSpec(
                    source_format="markdown",
                    target_format="xlsx",
                    action_name="",
                    label="Markdown → XLSX",
                ),
            ],
            requires=[],
            platforms=("windows",),
            capability_rules=[
                RouteCapabilityRule(
                    target_formats=("docx",),
                    required_capabilities=("python.docx",),
                    limitations=("probe",),
                )
            ],
            optimization_resources=[
                OptimizationResourceSpec(
                    id="official_document",
                    name="Official document",
                    action_name="gongwen",
                )
            ],
            extra={"homepage": "https://example.com"},
        )
        data = manifest.to_dict()
        m2 = PluginManifest.from_dict(data)
        assert m2.plugin_id == "docwen_plugin_markdown"
        assert len(m2.routes) == 2
        assert m2.routes[0].source_format == "markdown"
        assert m2.routes[0].target_format == "docx"
        assert m2.routes[1].label == "Markdown → XLSX"
        assert m2.platforms == ("windows",)
        assert m2.capability_rules[0].matches(m2.routes[0]) is True
        assert m2.capability_rules[0].matches(m2.routes[1]) is False
        assert m2.capability_rules[0].required_capabilities == ("python.docx",)
        assert m2.optimization_resources == [
            OptimizationResourceSpec(
                id="official_document",
                name="Official document",
                action_name="gongwen",
            )
        ]

    def test_defaults(self) -> None:
        manifest = PluginManifest(plugin_id="p", name="n", version="1")
        assert manifest.description == ""
        assert manifest.routes == []
        assert manifest.requires == []
        assert manifest.platforms == ("windows", "linux")
        assert manifest.capability_rules == []
        assert manifest.optimization_resources == []

    @pytest.mark.parametrize(
        "overrides",
        [
            {"id": ""},
            {"name": ""},
            {"action_name": "invalid action"},
        ],
    )
    def test_optimization_resource_rejects_invalid_declarations(self, overrides: dict[str, Any]) -> None:
        values: dict[str, Any] = {
            "id": "official_document",
            "name": "Official document",
            "action_name": "gongwen",
        }
        values.update(overrides)

        with pytest.raises(ValueError):
            OptimizationResourceSpec(**values)

    def test_optimization_resource_deserialization_rejects_legacy_scope_tables(self) -> None:
        with pytest.raises(ValueError, match=r"unsupported field.*scopes"):
            OptimizationResourceSpec.from_dict(
                {
                    "id": "official_document",
                    "name": "Official document",
                    "action_name": "gongwen",
                    "scopes": ["document_to_md"],
                }
            )


# ── WorkerRequest / WorkerResult ─────────────────────────────────────


class TestWorkerContractSerialization:
    def test_worker_request_round_trip(self) -> None:
        wr = WorkerRequest(
            task_id="task-1",
            route=RouteSpec(source_format="markdown", target_format="docx"),
            input_ref=FileRef(
                path="/tmp/a.md",
                format="markdown",
                category="markdown",
                media_type="text/markdown",
            ),
            output_policy=OutputPolicy(output_dir="/out"),
            typed_options={"remove_numbering": True},
            workspace_ref="/tmp/ws/task-1",
        )
        data = wr.to_dict()
        wr2 = WorkerRequest.from_dict(data)
        assert wr2.task_id == "task-1"
        assert wr2.route.source_format == "markdown"
        assert wr2.route.target_format == "docx"
        assert wr2.input_ref.media_type == "text/markdown"
        assert wr2.typed_options["remove_numbering"] is True
        assert wr2.workspace_ref == "/tmp/ws/task-1"

    def test_worker_result_round_trip_success(self) -> None:
        wr = WorkerResult(
            task_id="task-1",
            success=True,
            artifacts=[
                ArtifactManifest(artifact_id="a1", kind="primary", staging_path="/s/o.docx", suggested_name="o.docx")
            ],
            metrics=ConversionMetrics(duration_ms=200.0),
        )
        data = wr.to_dict()
        wr2 = WorkerResult.from_dict(data)
        assert wr2.success is True
        assert len(wr2.artifacts) == 1
        assert wr2.metrics.duration_ms == 200.0

    def test_worker_result_round_trip_failure(self) -> None:
        wr = WorkerResult(
            task_id="t2",
            success=False,
            error=ConversionErrorInfo(error_type="timeout", message="timed out"),
        )
        data = wr.to_dict()
        wr2 = WorkerResult.from_dict(data)
        assert wr2.success is False
        assert wr2.error is not None
        assert wr2.error.error_type == "timeout"

"""Fake implementation test — ROUTE-MD-DOCX-001 (markdown → docx).

This test proves that the core contracts can express a real conversion
route from route_matrix.csv.  It uses a fake plugin implementation
(stub) that satisfies ``ConverterPlugin``, a fake execution context,
and exercises the full flow from request to result.

Route: ROUTE-MD-DOCX-001
  source_format = "markdown"
  target_format = "docx"
  plugin        = "docwen_plugin_markdown"
  golden        = GOLDEN-001
"""

from __future__ import annotations

import uuid
from typing import Any

import pytest
from tests.support.config import FakeConfigView
from tests.support.execution import FakeExecutionContext
from tests.support.logging import FakePluginLogger
from tests.support.progress import FakeProgressSink
from tests.support.workspace import FakeWorkspaceHandle

from docwen_core.cancellation import CancellationToken
from docwen_core.errors import CancellationRequested
from docwen_core.models.artifact import ArtifactManifest
from docwen_core.models.file_ref import FileRef
from docwen_core.models.manifest import PluginManifest, RouteSpec
from docwen_core.models.request import ConversionRequest, OutputPolicy
from docwen_core.models.result import (
    ConversionDiagnostic,
    ConversionMetrics,
    ConversionResult,
)

pytestmark = pytest.mark.contract

# ═══════════════════════════════════════════════════════════════════════
# Fake Markdown → DOCX Plugin
# ═══════════════════════════════════════════════════════════════════════


class FakeMarkdownToDocxPlugin:
    """Fake plugin that simulates markdown → docx conversion.

    This plugin satisfies the ``ConverterPlugin`` protocol and registers
    two routes from route_matrix.csv:

    - ROUTE-MD-DOCX-001: markdown → docx
    - ROUTE-MD-XLSX-001: markdown → xlsx
    """

    @property
    def manifest(self) -> PluginManifest:
        return PluginManifest(
            plugin_id="docwen_plugin_markdown",
            name="Markdown Plugin",
            version="0.1.0",
            description="Converts markdown to docx, xlsx, and other formats",
            author="DocWen",
            routes=[
                RouteSpec(
                    source_format="markdown",
                    target_format="docx",
                    action_name="",
                    label="Markdown → DOCX",
                    options_schema={
                        "type": "object",
                        "properties": {
                            "remove_numbering": {"type": "boolean", "default": True},
                            "add_numbering": {"type": "boolean", "default": False},
                            "numbering_scheme": {"type": "string", "default": "hierarchical_standard"},
                            "formatting_mode": {"type": "string", "enum": ["apply", "ignore"], "default": "apply"},
                            "heading_merge_mode": {
                                "type": "string",
                                "enum": ["punct_required", "always"],
                                "default": "punct_required",
                            },
                            "list_separator": {"type": "string", "default": "、"},
                        },
                    },
                ),
                RouteSpec(
                    source_format="markdown",
                    target_format="xlsx",
                    action_name="",
                    label="Markdown → XLSX",
                ),
            ],
        )

    def can_handle(self, source_format: str, target_format: str, action_name: str = "") -> bool:
        for route in self.manifest.routes:
            if (
                route.source_format == source_format
                and route.target_format == target_format
                and (route.action_name == action_name or not action_name)
            ):
                return True
        return False

    def convert(self, context: FakeExecutionContext) -> ConversionResult:
        """Simulate a markdown → docx conversion.

        The fake does not actually parse markdown or produce a real docx.
        It performs the protocol-level steps a real plugin would:
        1. Check cancellation
        2. Report progress
        3. Read options from config/request
        4. Create staging artifacts
        5. Return result with metrics
        """
        task_id = context.request.request_id

        # 1. Check cancellation before starting
        context.cancellation.check()

        # 2. Report start
        context.progress.report_progress(0.0, "Starting markdown → docx conversion")
        context.logger.info(f"Starting conversion: {context.workspace.input_path}")

        # 3. Read options
        remove_numbering = context.request.options.get("remove_numbering", True)
        scheme = context.request.options.get("numbering_scheme", "hierarchical_standard")

        # 4. Simulate parsing (check cancellation again)
        context.cancellation.check()
        context.progress.report_progress(30.0, "Parsing markdown...")

        # 5. Simulate rendering
        context.cancellation.check()
        context.progress.report_progress(
            60.0, f"Rendering DOCX (scheme={scheme}, remove_numbering={remove_numbering})..."
        )

        # 6. Log a diagnostic
        context.progress.report_diagnostic(
            level="info",
            message="Markdown processed: 3 headings, 5 paragraphs, 1 table detected",
            code="MD-INFO-001",
        )

        # 7. Create staging artifact
        output_path = context.workspace.create_artifact_path("primary", ".docx")
        artifact = ArtifactManifest(
            artifact_id=str(uuid.uuid4()),
            kind="primary",
            staging_path=output_path,
            suggested_name=f"{context.workspace.input_path.split('/')[-1].rsplit('.', 1)[0]}.docx",
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            metadata={
                "page_count": 3,
                "heading_count": 3,
                "paragraph_count": 5,
                "table_count": 1,
            },
            is_primary=True,
        )
        context.workspace.add_artifact(artifact)

        context.progress.report_artifact_ready(artifact.artifact_id, artifact.suggested_name)

        # 8. Report completion
        context.progress.report_progress(100.0, "Conversion complete")
        context.logger.info(f"Conversion complete: {artifact.suggested_name}")

        return ConversionResult(
            task_id=task_id,
            success=True,
            artifacts=[artifact],
            diagnostics=[
                ConversionDiagnostic(
                    level="info",
                    message="Conversion completed successfully",
                    code="OK",
                ),
            ],
            error=None,
            metrics=ConversionMetrics(
                duration_ms=42.0,
                input_bytes=2048,
                output_bytes=15360,
                extra={"headings": 3, "paragraphs": 5, "tables": 1},
            ),
        )


# ═══════════════════════════════════════════════════════════════════════
# Tests
# ═══════════════════════════════════════════════════════════════════════


class TestFakeMarkdownToDocxConversion:
    """Test the fake markdown→docx conversion flow end-to-end."""

    @pytest.fixture
    def plugin(self) -> FakeMarkdownToDocxPlugin:
        return FakeMarkdownToDocxPlugin()

    @pytest.fixture
    def context_factory(self):
        """Return a factory so each test gets a fresh cancellation token."""

        def _make(
            *,
            options: dict[str, Any] | None = None,
            config: dict[str, Any] | None = None,
        ) -> tuple[FakeExecutionContext, FakeProgressSink, FakeWorkspaceHandle, CancellationToken]:
            request = ConversionRequest(
                request_id="req-md-docx-001",
                input_refs=[
                    FileRef(
                        path="/tmp/input/report.md",
                        format="markdown",
                        category="markdown",
                        size_bytes=2048,
                    )
                ],
                target_format="docx",
                options=options or {},
                output_policy=OutputPolicy(output_dir="/tmp/output"),
            )
            workspace = FakeWorkspaceHandle("/tmp/input/report.md", "/tmp/staging/req-md-docx-001")
            progress = FakeProgressSink()
            cancellation = CancellationToken()
            logger = FakePluginLogger()
            config_view = FakeConfigView(config or {})
            ctx = FakeExecutionContext(
                request=request,
                workspace=workspace,
                config=config_view,
                progress=progress,
                cancellation=cancellation,
                logger=logger,
            )
            return ctx, progress, workspace, cancellation

        return _make

    def test_manifest_declares_routes(self, plugin: FakeMarkdownToDocxPlugin) -> None:
        """Plugin manifest must declare ROUTE-MD-DOCX-001 and ROUTE-MD-XLSX-001."""
        m = plugin.manifest
        assert m.plugin_id == "docwen_plugin_markdown"
        assert len(m.routes) == 2

        docx_route = next(r for r in m.routes if r.target_format == "docx")
        assert docx_route.source_format == "markdown"
        assert docx_route.label == "Markdown → DOCX"

    def test_can_handle_matches_registered_routes(self, plugin: FakeMarkdownToDocxPlugin) -> None:
        assert plugin.can_handle("markdown", "docx") is True
        assert plugin.can_handle("markdown", "xlsx") is True
        assert plugin.can_handle("markdown", "pdf") is False
        assert plugin.can_handle("docx", "markdown") is False

    def test_conversion_success_flow(self, plugin: FakeMarkdownToDocxPlugin, context_factory) -> None:
        """A successful conversion produces a result with artifacts."""
        ctx, progress, workspace, _ = context_factory(
            options={"remove_numbering": True, "numbering_scheme": "hierarchical_standard"}
        )

        result = plugin.convert(ctx)

        assert result.success is True
        assert result.error is None
        assert len(result.artifacts) == 1
        assert result.artifacts[0].is_primary is True
        assert result.artifacts[0].kind == "primary"
        assert result.artifacts[0].suggested_name.endswith(".docx")
        assert (
            result.artifacts[0].media_type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        assert result.metrics.duration_ms > 0

        # Progress was reported
        assert len(progress.events) >= 3  # 0%, 30%, 60%, 100%
        # First and last percentages
        percentages = [e[0] for e in progress.events]
        assert percentages[0] == 0.0
        assert percentages[-1] == 100.0

        # Artifact was registered in workspace
        assert len(workspace.registered_artifacts) == 1
        registered = workspace.registered_artifacts[0]
        assert registered.is_primary is True

    def test_conversion_uses_options(self, plugin: FakeMarkdownToDocxPlugin, context_factory) -> None:
        """Options from the request should be available to the plugin."""
        ctx, progress, _, _ = context_factory(
            options={
                "remove_numbering": False,
                "add_numbering": True,
                "numbering_scheme": "legal_standard",
            }
        )

        result = plugin.convert(ctx)
        assert result.success is True

        # The progress messages should reflect the options
        progress_msgs = [e[1] for e in progress.events]
        combined = " ".join(progress_msgs)
        assert "remove_numbering=False" in combined

    def test_cancellation_during_conversion(self, plugin: FakeMarkdownToDocxPlugin, context_factory) -> None:
        """A cancelled token should cause CancellationRequested to be raised."""
        ctx, _, _, cancellation = context_factory()
        # Cancel before calling convert
        cancellation.cancel()

        with pytest.raises(CancellationRequested):
            plugin.convert(ctx)

    def test_result_serializable(self, plugin: FakeMarkdownToDocxPlugin, context_factory) -> None:
        """The ConversionResult must round-trip through to_dict/from_dict."""
        ctx, _, _, _ = context_factory()
        result = plugin.convert(ctx)

        data = result.to_dict()
        result2 = ConversionResult.from_dict(data)

        assert result2.success is True
        assert len(result2.artifacts) == 1
        assert result2.artifacts[0].is_primary is True
        assert result2.metrics.duration_ms == result.metrics.duration_ms

    def test_worker_request_construction(self) -> None:
        """A WorkerRequest can be constructed for ROUTE-MD-DOCX-001."""
        from docwen_core.models.worker import WorkerRequest

        wr = WorkerRequest(
            task_id="task-md-docx-001",
            route=RouteSpec(
                source_format="markdown",
                target_format="docx",
                label="Markdown → DOCX",
            ),
            input_ref=FileRef(
                path="/tmp/report.md",
                format="markdown",
                category="markdown",
            ),
            output_policy=OutputPolicy(output_dir="/tmp/out"),
            typed_options={
                "remove_numbering": True,
                "numbering_scheme": "hierarchical_standard",
            },
            workspace_ref="/tmp/ws/task-md-docx-001",
        )

        data = wr.to_dict()
        wr2 = WorkerRequest.from_dict(data)
        assert wr2.task_id == "task-md-docx-001"
        assert wr2.route.source_format == "markdown"
        assert wr2.route.target_format == "docx"

    def test_plugin_logger_receives_messages(self, plugin: FakeMarkdownToDocxPlugin, context_factory) -> None:
        """The logger should capture informational messages."""
        ctx, _, _, _ = context_factory()
        plugin.convert(ctx)
        assert len(ctx.logger.messages) >= 2  # start and complete
        assert any("start" in m["message"].lower() for m in ctx.logger.messages)
        assert any("complete" in m["message"].lower() for m in ctx.logger.messages)

    def test_route_can_be_registered_in_route_registry(self) -> None:
        """ROUTE-MD-DOCX-001 can be registered in a RouteRegistry."""
        from docwen_core.formats.routes import RouteRegistry

        reg = RouteRegistry()
        reg.register(
            RouteSpec(source_format="markdown", target_format="docx", label="Markdown → DOCX"),
            "docwen_plugin_markdown",
        )
        entry = reg.find("markdown", "docx")
        assert entry is not None
        assert entry.plugin_id == "docwen_plugin_markdown"

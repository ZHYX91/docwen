"""Integration test: full closed loop through the runtime pipeline.

Tests the complete path:
  Application Command → Workflow → RuntimePortAdapter → Runtime
    → RouteResolver → PluginRegistry → Workspace → Plugin
    → OutputFinalizer → ConversionResult

Uses a fake plugin (not a real conversion) to validate the contract
enforcement, event ordering, cancellation semantics, and error handling.
"""

from __future__ import annotations

import os
import tempfile
import threading
import uuid
import zipfile
from pathlib import Path
from typing import Any

import pytest

from docwen_core.models.artifact import ArtifactManifest
from docwen_core.models.file_ref import FileRef
from docwen_core.models.manifest import PluginManifest, RouteSpec
from docwen_core.models.request import PRECONVERSION_INTERMEDIATES_OPTION, ConversionRequest, OutputPolicy
from docwen_core.models.result import (
    ConversionDiagnostic,
    ConversionMetrics,
    ConversionResult,
)
from docwen_core.models.task import TaskEvent

pytestmark = pytest.mark.contract

_THREAD_COORDINATION_TIMEOUT_SECONDS = 60.0


def _write_template_package(path: Path, target: str, *, payload: str | None = None) -> None:
    main_part = "word/document.xml" if target == "docx" else "xl/workbook.xml"
    content_type = (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
        if target == "docx"
        else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
    )
    root = "document" if target == "docx" else "workbook"
    namespace = (
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        if target == "docx"
        else "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    )
    with zipfile.ZipFile(path, "w") as package:
        package.writestr(
            "[Content_Types].xml",
            (
                '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                f'<Override PartName="/{main_part}" ContentType="{content_type}"/>'
                "</Types>"
            ),
        )
        package.writestr(
            "_rels/.rels",
            (
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" '
                'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
                f'Target="{main_part}"/>'
                "</Relationships>"
            ),
        )
        package.writestr(main_part, f'<{root} xmlns="{namespace}"/>')
        if payload is not None:
            package.writestr("docwen-test-payload.txt", payload)


def _read_template_payload(path: Path) -> str:
    with zipfile.ZipFile(path) as package:
        return package.read("docwen-test-payload.txt").decode("utf-8")


class FakeClosedLoopPlugin:
    """A fake plugin that writes to staging and returns artifacts.

    This plugin:
    - Checks cancellation
    - Reports progress
    - Creates a staging artifact (writes a real file to staging)
    - Returns a ConversionResult with staging artifacts
    - NEVER writes to the final output directory
    """

    def __init__(
        self,
        plugin_id: str = "fake_plugin",
        source_format: str = "markdown",
        target_format: str = "docx",
        *,
        should_fail: bool = False,
        fail_message: str = "Simulated failure",
    ) -> None:
        self._should_fail = should_fail
        self._fail_message = fail_message
        self._manifest = PluginManifest(
            plugin_id=plugin_id,
            name=f"Fake {plugin_id}",
            version="0.1.0",
            description="Fake plugin for integration testing",
            routes=[
                RouteSpec(
                    source_format=source_format,
                    target_format=target_format,
                    label=f"{source_format} → {target_format}",
                ),
            ],
        )

    @property
    def manifest(self) -> PluginManifest:
        return self._manifest

    def can_handle(self, source_format: str, target_format: str, action_name: str = "") -> bool:
        for r in self._manifest.routes:
            if (
                r.source_format == source_format
                and r.target_format == target_format
                and (r.action_name == action_name or not action_name)
            ):
                return True
        return False

    def convert(self, context: Any) -> ConversionResult:
        task_id = context.request.request_id

        # 1. Check cancellation
        context.cancellation.check()

        # 2. Report progress
        context.progress.report_progress(0.0, "Starting fake conversion")
        context.logger.info(f"Converting: {context.workspace.input_path}")

        # 3. Simulate failure if configured
        if self._should_fail:
            from docwen_core.models.result import ConversionErrorInfo

            context.progress.report_diagnostic("error", self._fail_message, code="FAKE-ERR")
            return ConversionResult(
                task_id=task_id,
                success=False,
                error=ConversionErrorInfo(
                    error_type="conversion_failed",
                    message=self._fail_message,
                    diagnostic_code="FAKE-ERR",
                ),
                diagnostics=[
                    ConversionDiagnostic(level="error", message=self._fail_message, code="FAKE-ERR"),
                ],
                metrics=ConversionMetrics(extra={"failure_stage": "plugin"}),
            )

        # 4. Create staging artifact (write a real file)
        output_path = context.workspace.create_artifact_path("primary", ".docx")
        with open(output_path, "w") as f:
            f.write(f"Fake conversion output for task {task_id}\n")
            f.write(f"Source: {context.workspace.input_path}\n")

        # 5. Register artifact
        suggested = os.path.basename(context.workspace.input_path).rsplit(".", 1)[0] + ".docx"
        artifact = ArtifactManifest(
            artifact_id=str(uuid.uuid4()),
            kind="primary",
            staging_path=output_path,
            suggested_name=suggested,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            is_primary=True,
        )
        context.workspace.add_artifact(artifact)

        # 6. Report completion
        context.progress.report_progress(100.0, "Done")
        context.progress.report_artifact_ready(artifact.artifact_id, suggested)

        return ConversionResult(
            task_id=task_id,
            success=True,
            artifacts=[artifact],
            diagnostics=[
                ConversionDiagnostic(level="info", message="Fake conversion complete", code="OK"),
            ],
            metrics=ConversionMetrics(
                duration_ms=12.5,
                input_bytes=100,
                output_bytes=200,
                extra={"plugin_metric": "preserved"},
            ),
        )


class CapturingClosedLoopPlugin(FakeClosedLoopPlugin):
    """Fake plugin variant that records the options visible to converters."""

    def __init__(self, *args: Any, **kwargs: Any) -> None:
        super().__init__(*args, **kwargs)
        self.seen_options: list[dict[str, Any]] = []

    def convert(self, context: Any) -> ConversionResult:
        self.seen_options.append(dict(context.request.options))
        return super().convert(context)


class PreconversionIdentityPlugin(FakeClosedLoopPlugin):
    """Fake DOCX→Markdown plugin that exposes the admitted physical input."""

    def __init__(self) -> None:
        super().__init__("fake_docx2md", "docx", "md")
        self.seen: list[tuple[str, str]] = []

    def convert(self, context: Any) -> ConversionResult:
        input_path = Path(context.request.input_refs[0].path)
        content = _read_template_payload(input_path)
        self.seen.append((str(input_path), content))
        output_path = context.workspace.create_artifact_path("primary", ".md")
        Path(output_path).write_text(content, encoding="utf-8")
        artifact = ArtifactManifest(
            artifact_id=str(uuid.uuid4()),
            kind="primary",
            staging_path=output_path,
            suggested_name=f"{input_path.stem}.md",
            media_type="text/markdown",
            is_primary=True,
        )
        context.workspace.add_artifact(artifact)
        return ConversionResult(task_id=context.request.request_id, success=True, artifacts=[artifact])


class StreamedWarningPlugin(FakeClosedLoopPlugin):
    """Plugin that both streams and returns the same warning diagnostic."""

    def convert(self, context: Any) -> ConversionResult:
        warning = ConversionDiagnostic(
            level="warning",
            message="Optional OCR unavailable",
            code="OCR-BEST-EFFORT",
            location="sample.png",
        )
        context.progress.report_diagnostic(
            warning.level,
            warning.message,
            warning.code,
            warning.location,
        )
        result = super().convert(context)
        result.diagnostics.append(warning)
        return result


class StreamedWarningThenRaisePlugin(FakeClosedLoopPlugin):
    """Plugin that streams a warning before aborting conversion."""

    def __init__(self, failure: Exception) -> None:
        super().__init__("streamed_warning_then_raise", "markdown", "docx")
        self._failure = failure

    def convert(self, context: Any) -> ConversionResult:
        context.progress.report_diagnostic(
            "warning",
            "Optional OCR unavailable",
            code="OCR-BEST-EFFORT",
            location="sample.png",
        )
        raise self._failure


def build_runtime(plugin: FakeClosedLoopPlugin) -> Any:
    """Build the full runtime pipeline for testing."""
    from docwen_runtime.engine.route_resolver import RouteResolver
    from docwen_runtime.engine.task_manager import TaskManager
    from docwen_runtime.output.finalizer import OutputFinalizer
    from docwen_runtime.plugin_registry.registry import PluginRegistry
    from docwen_runtime.workspace.manager import WorkspaceManager

    registry = PluginRegistry()
    registry.register(plugin)

    resolver = RouteResolver(registry)

    with tempfile.TemporaryDirectory() as ws_root:
        ws_mgr = WorkspaceManager(root_dir=ws_root)
        finalizer = OutputFinalizer()
        task_mgr = TaskManager(registry, resolver, ws_mgr, finalizer)
        yield task_mgr


def build_app(task_manager: Any) -> Any:
    """Build the application layer wired to the runtime."""
    from docwen_application.controller import ApplicationController
    from docwen_runtime.adapters import RuntimePortAdapter

    adapter = RuntimePortAdapter(task_manager)
    controller = ApplicationController(runtime_port=adapter)
    return controller


from docwen_application.controller import ApplicationController

__all__ = (
    "PRECONVERSION_INTERMEDIATES_OPTION",
    "_THREAD_COORDINATION_TIMEOUT_SECONDS",
    "Any",
    "ApplicationController",
    "CapturingClosedLoopPlugin",
    "ConversionDiagnostic",
    "ConversionMetrics",
    "ConversionRequest",
    "ConversionResult",
    "FakeClosedLoopPlugin",
    "FileRef",
    "OutputPolicy",
    "Path",
    "PreconversionIdentityPlugin",
    "StreamedWarningPlugin",
    "StreamedWarningThenRaisePlugin",
    "TaskEvent",
    "_read_template_payload",
    "_write_template_package",
    "os",
    "pytest",
    "pytestmark",
    "tempfile",
    "threading",
)

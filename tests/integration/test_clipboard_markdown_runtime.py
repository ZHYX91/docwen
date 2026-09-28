"""Managed clipboard Markdown through Application, Runtime, plugin, and final publication."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest
from docx import Document

from docwen_application.controller import ApplicationController
from docwen_core.detection import inspect_utf8_markdown_snapshot
from docwen_core.models import FILE_INSPECTION_METADATA_KEY
from docwen_core.models.file_ref import FileRef
from docwen_core.models.request import ConversionRequest, OutputPolicy
from docwen_gui.clipboard_inputs import ClipboardInputStore
from docwen_plugin_markdown.plugin import MarkdownPlugin
from docwen_runtime.adapters import RuntimePortAdapter
from docwen_runtime.engine.route_resolver import RouteResolver
from docwen_runtime.engine.task_manager import TaskManager
from docwen_runtime.output.finalizer import OutputFinalizer
from docwen_runtime.plugin_registry.registry import PluginRegistry
from docwen_runtime.workspace.manager import WorkspaceManager

pytestmark = [pytest.mark.integration, pytest.mark.pr_gate, pytest.mark.release_gate]


def test_clipboard_markdown_snapshot_runs_through_existing_runtime_pipeline(tmp_path: Path) -> None:
    text = "| Name | Value |\n| --- | --- |\n| A | 00123 |\n"
    store = ClipboardInputStore(tmp_path / "managed")
    snapshot = store.create(text, display_name_template="Clipboard Markdown {index}.md")
    source = Path(snapshot.path)
    inspection = inspect_utf8_markdown_snapshot(source)
    assert inspection.workflow_category == "markdown"

    registry = PluginRegistry()
    registry.register(MarkdownPlugin())
    workspaces = WorkspaceManager(root_dir=str(tmp_path / "workspaces"))
    runtime = RuntimePortAdapter(
        TaskManager(
            registry,
            RouteResolver(registry),
            workspaces,
            OutputFinalizer(),
        )
    )
    controller = ApplicationController(runtime_port=runtime)
    request = ConversionRequest(
        request_id="clipboard-markdown-runtime",
        input_refs=[
            FileRef(
                path=str(source),
                format=inspection.detected_format,
                category=inspection.workflow_category,
                size_bytes=source.stat().st_size,
                metadata={FILE_INSPECTION_METADATA_KEY: inspection.to_dict()},
            )
        ],
        target_format="csv",
        output_policy=OutputPolicy(output_dir=str(tmp_path / "published")),
    )

    result = controller.execute_single(request)

    assert result.success is True, result.error
    assert source.read_bytes() == text.encode("utf-8")
    assert result.artifacts
    output = Path(result.artifacts[0].staging_path)
    assert output.is_file()
    delivered = output.read_text(encoding="utf-8-sig")
    assert "Name" in delivered
    assert "00123" in delivered
    assert len(workspaces) == 0

    store.close()


@pytest.mark.parametrize("text", ["Ordinary unmarked clipboard text 00123", "https://example.invalid/no-request"])
def test_plain_clipboard_text_exports_real_docx_without_reclassification(
    tmp_path: Path, round_trip_runtime: Any, text: str
) -> None:
    store = ClipboardInputStore(tmp_path / "managed")
    snapshot = store.create(text, display_name_template="Clipboard Markdown {index}.md")
    source = Path(snapshot.path)
    inspection = inspect_utf8_markdown_snapshot(source)
    request = ConversionRequest(
        request_id="clipboard-plain-docx",
        input_refs=[
            FileRef(
                path=str(source),
                format="markdown",
                category="markdown",
                metadata={FILE_INSPECTION_METADATA_KEY: inspection.to_dict()},
            )
        ],
        target_format="docx",
        output_policy=OutputPolicy(output_dir=str(tmp_path / "published")),
    )
    result = ApplicationController(runtime_port=round_trip_runtime).execute_single(request)
    assert result.success, result.error
    primary = next(artifact for artifact in result.artifacts if artifact.kind == "primary")
    document = Document(primary.staging_path)
    assert text in "\n".join(paragraph.text for paragraph in document.paragraphs)
    assert source.read_bytes() == text.encode("utf-8")
    store.close()

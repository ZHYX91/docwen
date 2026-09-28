"""Managed clipboard Markdown through Application, Runtime, plugin, and final publication."""

from __future__ import annotations

from pathlib import Path

import pytest

from docwen_application.controller import ApplicationController
from docwen_core.detection import inspect_file
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
    inspection = inspect_file(source)
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

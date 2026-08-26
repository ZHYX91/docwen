"""Request snapshot ownership for generic Markdown export policy."""

from __future__ import annotations

from collections.abc import Callable
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from threading import Barrier, Lock
from typing import Any

import pytest

from docwen_core.models.manifest import PluginManifest, RouteSpec
from docwen_core.models.result import ConversionResult

pytestmark = pytest.mark.contract


def _snapshot(marker: str) -> dict[str, Any]:
    return {
        "gui": {"language": {"locale": "zh_CN"}},
        "link": {
            "format": {
                "image_link_style": "markdown_embed" if marker == "A" else "wiki_embed",
                "md_file_link_style": "markdown_link" if marker == "A" else "wiki_link",
            }
        },
        "conversion": {
            "ocr_output": {
                "show_blockquote_title": True,
                "blockquote_title_override_by_locale": {"zh_CN": f"TITLE-{marker}"},
            },
            "export": {
                "base64_compress_enabled": marker == "A",
                "base64_compress_threshold_kb": 11 if marker == "A" else 22,
            },
        },
        "export": {
            "to_md_image_extraction_mode": "base64" if marker == "A" else "file",
            "to_md_ocr_placement_mode": "main_md" if marker == "A" else "separate_md",
        },
    }


class _CaptureMarkdownExportPlugin:
    def __init__(
        self,
        *,
        barrier: Barrier | None = None,
        before_capture: Callable[[], None] | None = None,
    ) -> None:
        self.seen: dict[str, tuple[Any, ...]] = {}
        self._barrier = barrier
        self._before_capture = before_capture
        self._lock = Lock()
        self._manifest = PluginManifest(
            plugin_id="request_markdown_export_probe",
            name="Request markdown export probe",
            version="1.0",
            description="captures request-scoped Markdown export policy",
            routes=[
                RouteSpec(
                    source_format="markdown",
                    target_format="md",
                    action_name="request_markdown_export_probe",
                    label="request Markdown export probe",
                )
            ],
        )

    @property
    def manifest(self) -> PluginManifest:
        return self._manifest

    def can_handle(self, source_format: str, target_format: str, action_name: str = "") -> bool:
        return source_format == "markdown" and target_format == "md" and action_name == "request_markdown_export_probe"

    def convert(self, context: Any) -> ConversionResult:
        if self._barrier is not None:
            self._barrier.wait(timeout=5)
        if self._before_capture is not None:
            self._before_capture()
        policy = context.markdown_export_semantics
        captured = (
            policy.image_link_style,
            policy.md_file_link_style,
            policy.image_extraction_mode,
            policy.ocr_placement_mode,
            policy.export_base64_compress_enabled,
            policy.export_base64_compress_threshold_kb,
            context.ocr_blockquote_title,
        )
        with self._lock:
            self.seen[context.request.request_id] = captured
        return ConversionResult(task_id=context.request.request_id, success=False)


def _task_manager(tmp_path: Path, plugin: _CaptureMarkdownExportPlugin) -> Any:
    from docwen_runtime.engine.route_resolver import RouteResolver
    from docwen_runtime.engine.task_manager import TaskManager
    from docwen_runtime.output.finalizer import OutputFinalizer
    from docwen_runtime.plugin_registry.registry import PluginRegistry
    from docwen_runtime.workspace.manager import WorkspaceManager

    plugins = PluginRegistry()
    plugins.register(plugin)
    return TaskManager(
        plugins,
        RouteResolver(plugins),
        WorkspaceManager(root_dir=str(tmp_path / "workspace")),
        OutputFinalizer(),
    )


def test_concurrent_requests_keep_markdown_export_snapshots_isolated(tmp_path: Path) -> None:
    from docwen_core.models.file_ref import FileRef
    from docwen_core.models.request import ConversionRequest

    source = tmp_path / "concurrent.md"
    source.write_text("# concurrent\n", encoding="utf-8")
    plugin = _CaptureMarkdownExportPlugin(barrier=Barrier(2))
    manager = _task_manager(tmp_path, plugin)
    requests = [
        ConversionRequest(
            request_id=f"concurrent-{marker}",
            input_refs=[FileRef(path=str(source), format="markdown", category="markdown")],
            target_format="md",
            action_name="request_markdown_export_probe",
            config_snapshot=_snapshot(marker),
        )
        for marker in ("A", "B")
    ]

    with ThreadPoolExecutor(max_workers=2) as executor:
        results = list(executor.map(manager.execute_single, requests))

    assert all(not result.success for result in results)
    assert plugin.seen == {
        "concurrent-A": ("markdown_embed", "markdown_link", "base64", "main_md", True, 11, "TITLE-A"),
        "concurrent-B": ("wiki_embed", "wiki_link", "file", "separate_md", False, 22, "TITLE-B"),
    }


def test_empty_snapshot_projects_deterministic_request_defaults(tmp_path: Path) -> None:
    from docwen_core.models.file_ref import FileRef
    from docwen_core.models.request import ConversionRequest

    source = tmp_path / "empty-snapshot.md"
    source.write_text("# empty snapshot\n", encoding="utf-8")
    plugin = _CaptureMarkdownExportPlugin()
    manager = _task_manager(tmp_path, plugin)

    result = manager.execute_single(
        ConversionRequest(
            request_id="empty-snapshot",
            input_refs=[FileRef(path=str(source), format="markdown", category="markdown")],
            target_format="md",
            action_name="request_markdown_export_probe",
        )
    )

    assert result.success is False
    assert plugin.seen == {
        "empty-snapshot": (
            "wiki_embed",
            "wiki_embed",
            "file",
            "main_md",
            True,
            100,
            "🖼️ **图片文字识别**：",
        )
    }

"""Direct TaskManager admission tests for Markdown OCR request options."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest

from docwen_core.models.file_ref import FileRef
from docwen_core.models.manifest import PluginManifest, RouteSpec
from docwen_core.models.request import ConversionRequest
from docwen_core.models.result import ConversionResult
from docwen_runtime.engine.route_resolver import RouteResolver
from docwen_runtime.engine.task_manager import TaskManager
from docwen_runtime.output.finalizer import OutputFinalizer
from docwen_runtime.plugin_registry.registry import PluginRegistry
from docwen_runtime.workspace.manager import WorkspaceManager

pytestmark = pytest.mark.unit


class _CaptureOptionsPlugin:
    def __init__(self) -> None:
        self.seen: dict[str, dict[str, object]] = {}
        self._manifest = PluginManifest(
            plugin_id="direct_ocr_admission_probe",
            name="Direct OCR admission probe",
            version="1.0",
            description="captures TaskManager-admitted request options",
            routes=[
                RouteSpec(
                    source_format="probe",
                    target_format=target_format,
                    action_name="capture_ocr_options",
                    label=f"capture OCR options for {target_format}",
                )
                for target_format in ("md", "docx")
            ],
        )

    @property
    def manifest(self) -> PluginManifest:
        return self._manifest

    def can_handle(self, source_format: str, target_format: str, action_name: str = "") -> bool:
        return source_format == "probe" and target_format in {"md", "docx"} and action_name == "capture_ocr_options"

    def convert(self, context: Any) -> ConversionResult:
        self.seen[context.request.request_id] = dict(context.request.options)
        return ConversionResult(task_id=context.request.request_id, success=False)


def _manager(tmp_path: Path, plugin: _CaptureOptionsPlugin) -> TaskManager:
    registry = PluginRegistry()
    registry.register(plugin)
    return TaskManager(
        registry,
        RouteResolver(registry),
        WorkspaceManager(root_dir=str(tmp_path / "workspace")),
        OutputFinalizer(),
    )


def _request(
    source: Path,
    *,
    request_id: str,
    target_format: str = "md",
    options: dict[str, object] | None = None,
    config_snapshot: dict[str, object] | None = None,
) -> ConversionRequest:
    return ConversionRequest(
        request_id=request_id,
        input_refs=[FileRef(path=str(source), format="probe", category="probe")],
        target_format=target_format,
        action_name="capture_ocr_options",
        options=dict(options or {}),
        config_snapshot=dict(config_snapshot or {}),
    )


def test_direct_single_projects_snapshot_values_without_mutating_caller(tmp_path: Path) -> None:
    source = tmp_path / "single.probe"
    source.write_text("probe", encoding="utf-8")
    plugin = _CaptureOptionsPlugin()
    manager = _manager(tmp_path, plugin)
    request = _request(
        source,
        request_id="single",
        config_snapshot={
            "image": {"ocr_language": "japanese"},
            "gui": {"language": {"locale": "ja_JP"}},
        },
    )

    manager.execute_single(request)

    assert plugin.seen["single"] == {"ocr_language": "japanese", "locale": "ja_JP"}
    assert request.options == {}


def test_direct_single_preserves_present_falsey_keys(tmp_path: Path) -> None:
    source = tmp_path / "falsey.probe"
    source.write_text("probe", encoding="utf-8")
    plugin = _CaptureOptionsPlugin()
    manager = _manager(tmp_path, plugin)
    request = _request(
        source,
        request_id="falsey",
        options={"ocr_language": "", "locale": None},
        config_snapshot={
            "image": {"ocr_language": "japanese"},
            "gui": {"language": {"locale": "ja_JP"}},
        },
    )

    manager.execute_single(request)

    assert plugin.seen["falsey"] == {"ocr_language": "", "locale": None}
    assert request.options == {"ocr_language": "", "locale": None}


@pytest.mark.parametrize(
    ("request_id", "target_format", "snapshot", "expected"),
    [
        ("partial", "md", {"image": {"ocr_language": "english"}}, {"ocr_language": "english", "locale": "zh_CN"}),
        ("empty", "md", {}, {}),
        (
            "non-markdown",
            "docx",
            {"image": {"ocr_language": "japanese"}, "gui": {"language": {"locale": "ja_JP"}}},
            {},
        ),
    ],
)
def test_direct_single_keeps_projection_scope(
    tmp_path: Path,
    request_id: str,
    target_format: str,
    snapshot: dict[str, object],
    expected: dict[str, object],
) -> None:
    source = tmp_path / f"{request_id}.probe"
    source.write_text("probe", encoding="utf-8")
    plugin = _CaptureOptionsPlugin()
    manager = _manager(tmp_path, plugin)

    manager.execute_single(
        _request(
            source,
            request_id=request_id,
            target_format=target_format,
            config_snapshot=snapshot,
        )
    )

    assert plugin.seen[request_id] == expected


def test_direct_batch_projects_each_derived_request(tmp_path: Path) -> None:
    sources = [tmp_path / "first.probe", tmp_path / "second.probe"]
    for source in sources:
        source.write_text("probe", encoding="utf-8")
    plugin = _CaptureOptionsPlugin()
    manager = _manager(tmp_path, plugin)
    request = ConversionRequest(
        request_id="batch",
        input_refs=[FileRef(path=str(source), format="probe", category="probe") for source in sources],
        target_format="md",
        action_name="capture_ocr_options",
        options={},
        config_snapshot={
            "image": {"ocr_language": "english"},
            "gui": {"language": {"locale": "en_US"}},
        },
    )

    manager.execute_batch(request)

    assert plugin.seen == {
        "batch-0": {"ocr_language": "english", "locale": "en_US"},
        "batch-1": {"ocr_language": "english", "locale": "en_US"},
    }
    assert request.options == {}

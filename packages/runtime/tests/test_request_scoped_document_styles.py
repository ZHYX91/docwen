"""TaskManager delivery of request-owned document style catalogs."""

from __future__ import annotations

from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from threading import Barrier, Lock
from typing import Any

import pytest

from docwen_core.models.manifest import PluginManifest, RouteSpec
from docwen_core.models.result import ConversionResult

pytestmark = pytest.mark.contract


class _CaptureDocumentStylesPlugin:
    def __init__(self, *, barrier: Barrier | None = None, target_format: str = "docx") -> None:
        self.seen: dict[str, tuple[str, str]] = {}
        self._target_format = target_format
        self._barrier = barrier
        self._lock = Lock()
        self._manifest = PluginManifest(
            plugin_id="request_document_style_probe",
            name="Request document style probe",
            version="1.0",
            description="captures request-scoped DOCX style catalogs",
            routes=[
                RouteSpec(
                    source_format="markdown",
                    target_format=target_format,
                    label="request DOCX style probe",
                )
            ],
        )

    @property
    def manifest(self) -> PluginManifest:
        return self._manifest

    def can_handle(self, source_format: str, target_format: str, action_name: str = "") -> bool:
        return source_format == "markdown" and target_format == self._target_format and not action_name

    def convert(self, context: Any) -> ConversionResult:
        if self._barrier is not None:
            self._barrier.wait(timeout=5)
        catalog = context.document_style_catalog
        captured = ("", "") if catalog is None else (catalog.locale, catalog.name_for("figure_caption"))
        with self._lock:
            self.seen[context.request.request_id] = captured
        return ConversionResult(task_id=context.request.request_id, success=False)


def _task_manager(tmp_path: Path, plugin: _CaptureDocumentStylesPlugin) -> Any:
    from docwen_runtime.engine.route_resolver import RouteResolver
    from docwen_runtime.engine.task_manager import TaskManager
    from docwen_runtime.output.finalizer import OutputFinalizer
    from docwen_runtime.plugin_registry.registry import PluginRegistry
    from docwen_runtime.workspace.manager import WorkspaceManager

    registry = PluginRegistry()
    registry.register(plugin)
    return TaskManager(
        registry,
        RouteResolver(registry),
        WorkspaceManager(root_dir=str(tmp_path / "workspace")),
        OutputFinalizer(),
    )


def _request(tmp_path: Path, request_id: str, locale: object) -> Any:
    from docwen_core.models.file_ref import FileRef
    from docwen_core.models.request import ConversionRequest

    source = tmp_path / f"{request_id}.md"
    source.write_text("# style catalog\n", encoding="utf-8")
    return ConversionRequest(
        request_id=request_id,
        input_refs=[FileRef(path=str(source), format="markdown", category="markdown")],
        target_format="docx",
        options={"locale": locale},
    )


def test_concurrent_requests_keep_style_catalogs_isolated(tmp_path: Path) -> None:
    plugin = _CaptureDocumentStylesPlugin(barrier=Barrier(2))
    manager = _task_manager(tmp_path, plugin)
    requests = (
        _request(tmp_path, "style-en", "en_US"),
        _request(tmp_path, "style-zh", "zh_CN"),
    )
    with ThreadPoolExecutor(max_workers=2) as executor:
        results = list(executor.map(manager.execute_single, requests))
    assert all(not result.success for result in results)
    assert plugin.seen == {
        "style-en": ("en_US", "Figure Caption"),
        "style-zh": ("zh_CN", "图题"),
    }


@pytest.mark.parametrize(
    ("locale", "code"),
    [
        ("", "DOCX_STYLE_LOCALE_INVALID"),
        ("en_GB", "DOCX_STYLE_LOCALE_UNSUPPORTED"),
    ],
)
def test_invalid_locale_is_a_stable_invalid_input(
    tmp_path: Path,
    locale: object,
    code: str,
) -> None:
    manager = _task_manager(tmp_path, _CaptureDocumentStylesPlugin())
    result = manager.execute_single(_request(tmp_path, f"style-invalid-{code}", locale))
    assert result.success is False
    assert result.error is not None
    assert result.error.error_type == "invalid_input"
    assert result.error.diagnostic_code == code
    assert [item.code for item in result.diagnostics if item.level == "error"] == [code]


def test_non_docx_route_never_requires_style_resources(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    import docwen_runtime.config as runtime_config

    plugin = _CaptureDocumentStylesPlugin(target_format="md")
    manager = _task_manager(tmp_path, plugin)
    request = _request(tmp_path, "style-non-docx", None)
    request.target_format = "md"
    monkeypatch.setattr(
        runtime_config,
        "build_document_style_catalog",
        lambda *_args, **_kwargs: pytest.fail("non-DOCX route loaded style resources"),
    )

    result = manager.execute_single(request)

    assert result.success is False
    assert plugin.seen["style-non-docx"] == ("", "")


def test_broken_packaged_style_resource_is_not_reported_as_user_input(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import docwen_runtime.config as runtime_config
    from docwen_runtime.config import DocumentStyleCatalogError

    def _broken_resource(*_args: object, **_kwargs: object) -> object:
        raise DocumentStyleCatalogError(
            "DOCX_STYLE_LOCALE_MISSING",
            "packaged locale missing",
            error_type="conversion_failed",
        )

    monkeypatch.setattr(runtime_config, "build_document_style_catalog", _broken_resource)
    manager = _task_manager(tmp_path, _CaptureDocumentStylesPlugin())
    result = manager.execute_single(_request(tmp_path, "style-resource-broken", "en_US"))

    assert result.success is False
    assert result.error is not None
    assert result.error.error_type == "conversion_failed"
    assert result.error.diagnostic_code == "DOCX_STYLE_LOCALE_MISSING"

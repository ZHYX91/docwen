"""Builder → Application → Runtime coverage for structured resource groups."""

from __future__ import annotations

import hashlib
import json
import logging
import threading
import time
from pathlib import Path
from types import SimpleNamespace
from typing import TYPE_CHECKING, cast

import pytest

from docwen_application.controller import ApplicationController
from docwen_core.detection import inspect_structured_clipboard_snapshot
from docwen_core.models import FILE_INSPECTION_METADATA_KEY
from docwen_core.models.artifact import ARTIFACT_KIND_PRIMARY, ArtifactManifest
from docwen_core.models.clipboard_document import CLIPBOARD_DOCUMENT_SCHEMA
from docwen_core.models.file_ref import FileRef
from docwen_core.models.manifest import PluginManifest, RouteSpec
from docwen_core.models.request import OutputPolicy
from docwen_core.models.result import ConversionResult
from docwen_gui.clipboard_inputs import ClipboardInputStore
from docwen_gui.execution_requests import ExecutionRequestBuilder
from docwen_gui.path_identity import normalize_path
from docwen_plugin_markdown.plugin import MarkdownPlugin
from docwen_runtime._execution_context import _RuntimePluginLogger
from docwen_runtime.adapters import RuntimePortAdapter
from docwen_runtime.engine.route_resolver import RouteResolver
from docwen_runtime.engine.task_manager import TaskManager
from docwen_runtime.output.finalizer import OutputFinalizer
from docwen_runtime.plugin_registry.registry import PluginRegistry
from docwen_runtime.templates import TemplateRegistry
from docwen_runtime.workspace.manager import WorkspaceManager

if TYPE_CHECKING:
    from docwen_core.models.request import ConversionRequest
    from docwen_gui.view_models.batch_list_vm import BatchListViewModel
    from docwen_gui.view_models.main_window_vm import MainWindowViewModel

pytestmark = [pytest.mark.integration, pytest.mark.pr_gate, pytest.mark.release_gate]


def _payload(resource_id: str, resource_bytes: bytes, *, sentinel: str = "") -> bytes:
    logical_path = f"resources/{resource_id}.bin"
    return json.dumps(
        {
            "schema": CLIPBOARD_DOCUMENT_SCHEMA,
            "blocks": [
                {
                    "type": "paragraph",
                    "inlines": [
                        {
                            "type": "text",
                            "value": f"document-{resource_id}:{sentinel}",
                        },
                        {
                            "type": "image",
                            "resourceId": resource_id,
                            "alt": f"{resource_id}{sentinel}",
                            "missingReason": "",
                        },
                    ],
                }
            ],
            "resources": [
                {
                    "resourceId": resource_id,
                    "logicalPath": logical_path,
                    "mediaType": "application/octet-stream",
                    "sizeBytes": len(resource_bytes),
                    "sha256": hashlib.sha256(resource_bytes).hexdigest(),
                }
            ],
        },
        ensure_ascii=False,
        separators=(",", ":"),
        sort_keys=True,
    ).encode("utf-8")


def _bundle(store: ClipboardInputStore, resource_id: str, resource_bytes: bytes, *, sentinel: str = ""):
    logical_path = f"resources/{resource_id}.bin"
    return store.create_bundle(
        _payload(resource_id, resource_bytes, sentinel=sentinel),
        display_name_template="Clipboard Document {index}.dwclip",
        preview=resource_id,
        resources=((resource_id, logical_path, "application/octet-stream", resource_bytes),),
    )


def _source_ref(bundle) -> FileRef:
    inspection = inspect_structured_clipboard_snapshot(bundle.main.path)
    return FileRef(
        path=bundle.main.path,
        format="clipboard_document",
        category="markdown",
        size_bytes=inspection.size_bytes,
        metadata={FILE_INSPECTION_METADATA_KEY: inspection.to_dict()},
    )


def _builder(store: ClipboardInputStore, refs: list[FileRef]) -> ExecutionRequestBuilder:
    template = next(
        item for item in TemplateRegistry.default().list_templates("docx") if item.name == "English General Template"
    )
    contexts = {normalize_path(ref.path): ("clipboard_document", "markdown") for ref in refs}
    return ExecutionRequestBuilder(
        cast("MainWindowViewModel", SimpleNamespace(files=refs, controller=None)),
        cast("BatchListViewModel", SimpleNamespace(get_file_entry=lambda _path: None)),
        file_contexts=lambda: contexts,
        selected_template=lambda: ("docx", template.id),
        source_label=lambda path: store.descriptor(path).display_name if store.descriptor(path) else None,
        synthetic_input=store.is_snapshot,
        snapshot_bundle=store.bundle,
    )


def _controller(tmp_path: Path, plugin) -> ApplicationController:
    registry = PluginRegistry()
    registry.register(plugin)
    return ApplicationController(
        runtime_port=RuntimePortAdapter(
            TaskManager(
                registry,
                RouteResolver(registry),
                WorkspaceManager(root_dir=str(tmp_path / "workspaces")),
                OutputFinalizer(),
            )
        )
    )


def _build_grouped_request(tmp_path: Path, *, sentinel: str = ""):
    store = ClipboardInputStore(tmp_path / "managed")
    first = _bundle(store, "asset-one", b"first-resource", sentinel=sentinel)
    second = _bundle(store, "asset-two", b"second-resource", sentinel=sentinel)
    refs = [_source_ref(first), _source_ref(second)]
    builder = _builder(store, refs)
    parent, context = builder.batch(
        file_paths=[first.main.path, second.main.path],
        target_format="md",
        action_name="",
        options={"markdown_extensions": {"output": {"structural_tables": False}}},
        route_options=("markdown_extensions",),
        output_policy=OutputPolicy(output_dir=str(tmp_path / "out")),
    )
    groups = cast("tuple[ConversionRequest, ...]", context.pop("_document_group_requests"))
    store.sync_visible([first.main.path, second.main.path])
    return store, (first, second), parent, groups


def test_builder_application_runtime_executes_two_resource_groups_in_order(tmp_path: Path) -> None:
    store, bundles, parent, groups = _build_grouped_request(tmp_path)
    controller = _controller(tmp_path, MarkdownPlugin())
    try:
        results = controller.execute_document_group_batch(parent, groups)
        assert [result.task_id for result in results] == [
            f"{parent.request_id}-0",
            f"{parent.request_id}-1",
        ]
        assert [result.success for result in results] == [True, True]
        for result, resource_id in zip(results, ("asset-one", "asset-two"), strict=True):
            codes = {item.code for item in result.diagnostics}
            assert "CLIPBOARD-IMAGE-RESOURCE-UNAVAILABLE" not in codes
            assert "CLIPBOARD-IMAGE-RESOURCE-NOT-RENDERED" in codes
            assert resource_id not in "\n".join(item.message for item in result.diagnostics)
            primary = next(item for item in result.artifacts if item.is_primary)
            text = Path(primary.staging_path).read_text(encoding="utf-8")
            assert f"document-{resource_id}" in text
            assert resource_id in text
        assert bundles[0].resources[0].path != bundles[1].resources[0].path
    finally:
        store.close()


@pytest.mark.parametrize("failure", ["none", "tampered", "raw-exception"])
def test_resource_group_content_and_exceptions_stay_out_of_shared_messages(tmp_path, monkeypatch, caplog, failure):
    sentinel = "RESOURCE_PRIVATE_83b2"
    private_text = f"https://example.invalid/{sentinel}/private?alt={sentinel}"
    store, bundles, parent, groups = _build_grouped_request(tmp_path, sentinel=private_text)
    controller = _controller(tmp_path, MarkdownPlugin())
    loggers: list[_RuntimePluginLogger] = []
    original_init = _RuntimePluginLogger.__init__
    attempted: list[str] = []

    def capture_logger(self, task_id: str) -> None:
        original_init(self, task_id)
        loggers.append(self)

    def fail_preconversion(request, **_kwargs):
        attempted.append(request.request_id)
        raise ValueError(f"raw authored exception: {sentinel}")

    monkeypatch.setattr(_RuntimePluginLogger, "__init__", capture_logger)
    caplog.set_level(logging.DEBUG)
    if failure == "tampered":
        Path(bundles[0].resources[0].path).write_bytes(private_text.encode())
    elif failure == "raw-exception":
        monkeypatch.setattr(controller, "_maybe_preconvert", fail_preconversion)
    try:
        results = controller.execute_document_group_batch(parent, groups)
        assert len(results) == 2
        if failure == "none":
            assert all(result.success for result in results)
            assert all(
                "CLIPBOARD-IMAGE-RESOURCE-NOT-RENDERED" in {item.code for item in result.diagnostics}
                for result in results
            )
            for result in results:
                primary = next(item for item in result.artifacts if item.is_primary)
                assert sentinel in Path(primary.staging_path).read_text(encoding="utf-8")
        elif failure == "tampered":
            assert not results[0].success and results[1].success
            assert results[0].artifacts == []
        else:
            assert attempted == [f"{parent.request_id}-0", f"{parent.request_id}-1"]
            assert all(not result.success and not result.artifacts for result in results)
            assert all(
                result.error is not None and result.error.message == "Document group execution failed."
                for result in results
            )
        if failure != "raw-exception":
            assert loggers, "Observe the actual Runtime logging channel"
        assert sentinel not in caplog.text
        assert sentinel not in repr([logger.messages for logger in loggers])
        assert sentinel not in repr([result.diagnostics for result in results])
        assert sentinel not in repr([result.error for result in results])
    finally:
        store.close()


def test_resource_tamper_after_build_fails_only_its_group_and_keeps_order(tmp_path: Path) -> None:
    store, bundles, parent, groups = _build_grouped_request(tmp_path)
    controller = _controller(tmp_path, MarkdownPlugin())
    try:
        Path(bundles[0].resources[0].path).write_bytes(b"tampered-after-build")
        results = controller.execute_document_group_batch(parent, groups)

        assert [result.task_id for result in results] == [
            f"{parent.request_id}-0",
            f"{parent.request_id}-1",
        ]
        assert results[0].success is False
        assert results[0].error is not None
        assert results[0].error.message == "typed input copy failed integrity verification"
        assert results[0].artifacts == []
        assert results[1].success is True
    finally:
        store.close()


class _SlowStructuredPlugin:
    def __init__(self, entered: threading.Event) -> None:
        self._entered = entered
        self.started: list[str] = []

    @property
    def manifest(self) -> PluginManifest:
        return PluginManifest(
            plugin_id="slow_structured_groups",
            name="Slow Structured Groups",
            version="0.1.0",
            routes=[RouteSpec(source_format="clipboard_document", target_format="md")],
        )

    def can_handle(self, source_format: str, target_format: str, action_name: str = "") -> bool:
        return source_format == "clipboard_document" and target_format == "md" and not action_name

    def convert(self, context) -> ConversionResult:
        self.started.append(context.request.request_id)
        self._entered.set()
        deadline = time.monotonic() + 5.0
        while time.monotonic() < deadline:
            context.cancellation.check()
            time.sleep(0.005)
        output = context.workspace.create_artifact_path("primary", ".md")
        Path(output).write_text("unexpected completion", encoding="utf-8")
        return ConversionResult(
            task_id=context.request.request_id,
            success=True,
            artifacts=[
                ArtifactManifest(
                    artifact_id=f"{context.request.request_id}-primary",
                    kind=ARTIFACT_KIND_PRIMARY,
                    staging_path=output,
                    suggested_name="slow.md",
                    media_type="text/markdown",
                    is_primary=True,
                )
            ],
        )


def test_parent_cancel_during_first_group_cancels_current_and_future_groups(tmp_path: Path) -> None:
    store, _bundles, parent, groups = _build_grouped_request(tmp_path)
    entered = threading.Event()
    plugin = _SlowStructuredPlugin(entered)
    controller = _controller(tmp_path, plugin)
    holder: list[list[ConversionResult]] = []

    def run() -> None:
        holder.append(controller.execute_document_group_batch(parent, groups))

    worker = threading.Thread(target=run, daemon=True)
    worker.start()
    try:
        assert entered.wait(2.0)
        controller.cancel(parent.request_id)
        worker.join(5.0)
        assert not worker.is_alive()
        assert len(holder) == 1
        results = holder[0]
        assert [result.task_id for result in results] == [
            f"{parent.request_id}-0",
            f"{parent.request_id}-1",
        ]
        assert all(result.error is not None and result.error.error_type == "cancelled" for result in results)
        assert plugin.started == [f"{parent.request_id}-0"]
    finally:
        controller.cancel(parent.request_id)
        worker.join(1.0)
        store.close()


def test_parent_cancel_before_grouped_worker_prevents_all_runtime_starts(tmp_path: Path) -> None:
    store, _bundles, parent, groups = _build_grouped_request(tmp_path)
    entered = threading.Event()
    plugin = _SlowStructuredPlugin(entered)
    controller = _controller(tmp_path, plugin)
    reservation = controller.prepare_execution_cancellation(parent, batch=True)
    try:
        controller.cancel(parent.request_id)
        results = controller.execute_document_group_batch(parent, groups)

        assert not entered.is_set()
        assert plugin.started == []
        assert [result.task_id for result in results] == [
            f"{parent.request_id}-0",
            f"{parent.request_id}-1",
        ]
        assert all(result.error is not None and result.error.error_type == "cancelled" for result in results)
    finally:
        controller.release_execution_cancellation(parent.request_id, reservation)
        store.close()


def test_execution_thread_keeps_frozen_validation_failure_scoped_to_one_group(tmp_path: Path) -> None:
    from docwen_gui.qt_bridge.execution import ExecutionThread

    store, bundles, parent, groups = _build_grouped_request(tmp_path)
    controller = _controller(tmp_path, MarkdownPlugin())
    Path(bundles[0].resources[0].path).write_bytes(b"tampered-before-worker")
    results: list[list[ConversionResult]] = []
    errors: list[str] = []
    thread = ExecutionThread(
        controller=controller,
        request=parent,
        context={"request_id": parent.request_id},
        batch_execution=True,
        document_group_requests=groups,
        clipboard_bundles=bundles,
    )
    thread.result_signal.connect(lambda result, _context: results.append(result))
    thread.error_signal.connect(lambda message, _context: errors.append(message))
    try:
        thread.run()
        assert errors == []
        assert len(results) == 1
        batch = results[0]
        assert [item.task_id for item in batch] == [
            f"{parent.request_id}-0",
            f"{parent.request_id}-1",
        ]
        assert batch[0].success is False
        assert batch[0].error is not None
        assert batch[0].error.error_type == "invalid_input"
        assert batch[0].error.message == "Document group validation failed."
        assert "tampered" not in batch[0].error.message
        assert batch[1].success is True
    finally:
        store.close()

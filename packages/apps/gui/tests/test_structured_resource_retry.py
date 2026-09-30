"""Actual failed-history retry must retain and revalidate the complete bundle."""

from __future__ import annotations

import hashlib
import json
from dataclasses import replace
from pathlib import Path

import pytest

from docwen_application.controller import ApplicationController
from docwen_core.detection import inspect_structured_clipboard_snapshot
from docwen_core.models.artifact import ARTIFACT_KIND_PRIMARY, ArtifactManifest
from docwen_core.models.clipboard_document import CLIPBOARD_DOCUMENT_SCHEMA
from docwen_core.models.manifest import PluginManifest, RouteSpec
from docwen_core.models.result import ConversionErrorInfo, ConversionResult
from docwen_gui.app import create_main_window
from docwen_gui.path_identity import normalize_path
from docwen_runtime.adapters import RuntimePortAdapter
from docwen_runtime.capabilities import build_runtime_capability_projection
from docwen_runtime.engine.route_resolver import RouteResolver
from docwen_runtime.engine.task_manager import TaskManager
from docwen_runtime.output.finalizer import OutputFinalizer
from docwen_runtime.plugin_registry.registry import PluginRegistry
from docwen_runtime.templates import TemplateRegistry
from docwen_runtime.workspace.manager import WorkspaceManager

pytestmark = [pytest.mark.gui, pytest.mark.pr_gate, pytest.mark.release_gate]


class _FailOnceResourcePlugin:
    def __init__(self) -> None:
        self.observed: list[bytes] = []

    @property
    def manifest(self) -> PluginManifest:
        return PluginManifest(
            plugin_id="resource_retry_probe",
            name="Resource Retry Probe",
            version="0.1.0",
            routes=[RouteSpec(source_format="clipboard_document", target_format="md")],
        )

    def can_handle(self, source_format: str, target_format: str, action_name: str = "") -> bool:
        return source_format == "clipboard_document" and target_format == "md" and not action_name

    def convert(self, context) -> ConversionResult:
        resources = context.workspace.input_resources("linked_resource")
        assert len(resources) == 1
        content = Path(resources[0].path).read_bytes()
        self.observed.append(content)
        if len(self.observed) == 1:
            return ConversionResult(
                task_id=context.request.request_id,
                success=False,
                error=ConversionErrorInfo(error_type="conversion_failed", message="Controlled first failure."),
            )
        output = context.workspace.create_artifact_path("primary", ".md")
        Path(output).write_bytes(content)
        return ConversionResult(
            task_id=context.request.request_id,
            success=True,
            artifacts=[
                ArtifactManifest(
                    artifact_id=f"{context.request.request_id}-primary",
                    kind=ARTIFACT_KIND_PRIMARY,
                    staging_path=output,
                    suggested_name="resource-retry.md",
                    media_type="text/markdown",
                    is_primary=True,
                )
            ],
        )


@pytest.mark.parametrize("resource_state", ["unchanged", "missing", "tampered"])
def test_real_failed_history_retry_revalidates_resource_group(
    qapp,
    qtbot,
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    resource_state: str,
) -> None:
    plugin = _FailOnceResourcePlugin()
    registry = PluginRegistry()
    registry.register(plugin)
    runtime = RuntimePortAdapter(
        TaskManager(
            registry,
            RouteResolver(registry),
            WorkspaceManager(root_dir=str(tmp_path / "runtime-workspaces")),
            OutputFinalizer(),
        ),
        capability_provider=lambda: build_runtime_capability_projection(
            registry.list_manifests(), platform_id="windows", egress_guard_status={}
        ),
    )
    controller = ApplicationController(runtime_port=runtime)
    controller.start()
    window = create_main_window(controller=controller, clipboard_input_root=str(tmp_path / "clipboard-inputs"))
    qtbot.addWidget(window)
    window.show()
    output_dir = tmp_path / "outputs"
    output_dir.mkdir()
    monkeypatch.setattr(
        window._workflow,
        "_prepare_output_policy",
        lambda _paths, _mode, policy: replace(policy, output_dir=str(output_dir), per_input_output_dirs={}),
    )
    original = b"original frozen resource bytes"
    logical_path = "resources/asset.bin"
    payload = json.dumps(
        {
            "schema": CLIPBOARD_DOCUMENT_SCHEMA,
            "blocks": [
                {
                    "type": "paragraph",
                    "inlines": [{"type": "image", "resourceId": "asset", "alt": "asset", "missingReason": ""}],
                }
            ],
            "resources": [
                {
                    "resourceId": "asset",
                    "logicalPath": logical_path,
                    "mediaType": "application/octet-stream",
                    "sizeBytes": len(original),
                    "sha256": hashlib.sha256(original).hexdigest(),
                }
            ],
        }
    ).encode("utf-8")
    try:
        store = window._clipboard_store_for_paste()
        bundle = store.create_bundle(
            payload,
            display_name_template="Clipboard Document {index}.dwclip",
            preview="controlled resource retry",
            resources=(("asset", logical_path, "application/octet-stream", original),),
        )
        window._input_area_vm.add_files([bundle.main.path], file_inspector=inspect_structured_clipboard_snapshot)
        qtbot.waitUntil(lambda: not window.view_model.inspection_busy)
        qtbot.waitUntil(lambda: len(window.view_model.files) == 1)
        template = next(
            item
            for item in TemplateRegistry.default().list_templates("docx")
            if item.name == "English General Template"
        )
        assert window._template_selector is not None
        selector = window._template_selector.get_selector("docx")
        assert selector is not None
        selector.select_template(template.id, selection_source="user")
        window._workflow.single(file_path=bundle.main.path, target_format="md", action_name="", options={})

        def failed() -> bool:
            record = window._task_history.get(window._info_area_vm.task_summary.operation_id)
            return record is not None and bool(record.failed_paths) and not window._execution.busy

        qtbot.waitUntil(failed, timeout=5000)
        original_operation = window._info_area_vm.task_summary.operation_id
        assert plugin.observed == [original]
        window.view_model.remove_file(bundle.main.path)
        qapp.processEvents()
        resource = Path(bundle.resources[0].path)
        assert Path(bundle.main.path).is_file() and resource.is_file()
        if resource_state == "missing":
            resource.unlink()
        elif resource_state == "tampered":
            resource.write_bytes(b"changed resource bytes")
        qapp.clipboard().setText("New clipboard contents must never replace the saved retry input.")
        window._retry_failed_request()

        if resource_state == "unchanged":
            qtbot.waitUntil(lambda: len(plugin.observed) == 2 and not window._execution.busy, timeout=5000)
            record = window._task_history.get(window._info_area_vm.task_summary.operation_id)
            assert record is not None and not record.failed_paths
            outcome = record.outcomes[normalize_path(bundle.main.path)]
            assert outcome.output_path
            assert Path(outcome.output_path).read_bytes() == original
            assert plugin.observed == [original, original]
        else:
            qapp.processEvents()
            assert plugin.observed == [original]
            assert window._info_area_vm.task_summary.operation_id == original_operation
            assert not window._execution.busy
            assert not list(output_dir.rglob("*.md"))
    finally:
        window.close()
        qapp.clipboard().clear()
        qapp.processEvents()
        controller.stop()

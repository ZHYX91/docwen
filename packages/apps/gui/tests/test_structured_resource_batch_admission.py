"""Real GUI batch admission keeps damaged clipboard groups isolated."""

from __future__ import annotations

import json
from dataclasses import replace
from pathlib import Path

import pytest
from packages.apps.gui.tests.test_structured_resource_main_window import _create_bundle, _select_docx_template

from docwen_application.controller import ApplicationController
from docwen_core.detection import inspect_structured_clipboard_snapshot
from docwen_core.models.artifact import ARTIFACT_KIND_PRIMARY, ArtifactManifest
from docwen_core.models.clipboard_document import CLIPBOARD_DOCUMENT_SCHEMA
from docwen_core.models.manifest import PluginManifest, RouteSpec
from docwen_core.models.result import ConversionResult
from docwen_gui.app import create_main_window
from docwen_gui.path_identity import normalize_path
from docwen_runtime.adapters import RuntimePortAdapter
from docwen_runtime.capabilities import build_runtime_capability_projection
from docwen_runtime.engine.route_resolver import RouteResolver
from docwen_runtime.engine.task_manager import TaskManager
from docwen_runtime.output.finalizer import OutputFinalizer
from docwen_runtime.plugin_registry.registry import PluginRegistry
from docwen_runtime.workspace.manager import WorkspaceManager

pytestmark = [pytest.mark.gui, pytest.mark.pr_gate, pytest.mark.release_gate]


class _ResourceEcho:
    def __init__(self):
        self.observed = []

    @property
    def manifest(self):
        return PluginManifest(
            plugin_id="resource_preflight_probe",
            name="Resource Preflight Probe",
            version="0.1.0",
            routes=[RouteSpec(source_format="clipboard_document", target_format="md")],
        )

    def can_handle(self, source_format, target_format, action_name=""):
        return source_format == "clipboard_document" and target_format == "md" and not action_name

    def convert(self, context):
        resources = context.workspace.input_resources("linked_resource")
        if resources:
            data = Path(resources[0].path).read_bytes()
        else:
            document = json.loads(Path(context.workspace.input_path).read_bytes())
            data = document["blocks"][0]["inlines"][0]["value"].encode()
        self.observed.append((context.request.request_id, data))
        output = context.workspace.create_artifact_path("primary", ".md")
        Path(output).write_bytes(data)
        return ConversionResult(
            task_id=context.request.request_id,
            success=True,
            artifacts=[
                ArtifactManifest(
                    artifact_id=context.request.request_id + "-primary",
                    kind=ARTIFACT_KIND_PRIMARY,
                    staging_path=output,
                    suggested_name="resource.md",
                    media_type="text/markdown",
                    is_primary=True,
                )
            ],
        )


@pytest.mark.parametrize(
    "fault",
    [
        "intact",
        "pre-build-missing",
        "pre-build-tampered",
        "pending-missing",
        "pending-tampered",
        "pre-build-main-missing",
        "pre-build-main-tampered",
        "pending-main-missing",
        "pending-main-tampered",
        "pre-build-main-changed",
        "pending-main-changed",
        "pre-build-marker-missing",
        "pending-marker-tampered",
        "pre-build-partial",
        "pending-partial",
        "pre-build-main-missing-empty",
        "pending-main-tampered-empty",
    ],
)
def test_real_main_window_batch_keeps_valid_sibling_before_runtime(qapp, qtbot, tmp_path, monkeypatch, fault):
    plugin = _ResourceEcho()
    registry = PluginRegistry()
    registry.register(plugin)
    runtime = RuntimePortAdapter(
        TaskManager(
            registry, RouteResolver(registry), WorkspaceManager(root_dir=str(tmp_path / "runtime")), OutputFinalizer()
        ),
        capability_provider=lambda: build_runtime_capability_projection(
            registry.list_manifests(), platform_id="windows", egress_guard_status={}
        ),
    )
    controller = ApplicationController(runtime_port=runtime)
    controller.start()
    window = create_main_window(controller=controller, clipboard_input_root=str(tmp_path / "clipboard"))
    qtbot.addWidget(window)
    window.show()
    monkeypatch.setattr(
        window._workflow,
        "_prepare_output_policy",
        lambda _paths, _mode, policy: replace(policy, output_dir=str(tmp_path / "out"), per_input_output_dirs={}),
    )
    try:
        window.view_model.set_mode("batch")

        def create_bundle(resource_id, data):
            if not fault.endswith("-empty"):
                return _create_bundle(window, resource_id, data)
            payload = {
                "schema": CLIPBOARD_DOCUMENT_SCHEMA,
                "blocks": [{"type": "paragraph", "inlines": [{"type": "text", "value": data.decode()}]}],
                "resources": [],
            }
            return window._clipboard_store_for_paste().create_bundle(
                json.dumps(payload).encode(), display_name_template="Empty {index}.dwclip", preview=resource_id
            )

        first = create_bundle("asset-one", b"first-resource")
        second = create_bundle("asset-two", b"second-resource")
        window._input_area_vm.add_files(
            [first.main.path, second.main.path], file_inspector=inspect_structured_clipboard_snapshot
        )
        qtbot.waitUntil(lambda: not window.view_model.inspection_busy and len(window.view_model.files) == 2)
        _select_docx_template(window)

        def corrupt():
            case = fault.removesuffix("-empty")
            if case.endswith("partial"):
                Path(first.root_path, ".bundle.partial").write_bytes(b"incomplete")
                return
            if "marker" in case:
                resource = Path(first.marker_path)
            else:
                resource = Path(first.main.path if "main" in case else first.resources[0].path)
            if case.endswith("missing"):
                resource.unlink()
            elif case.endswith("changed"):
                resource.write_bytes(resource.read_bytes().replace(b"opaque resource", b"private content"))
            else:
                resource.write_bytes(b"tampered-before-runtime")

        if fault.startswith("pre-build"):
            corrupt()
        elif fault.startswith("pending"):
            original_build = window._requests.batch

            def after_build(**kwargs):
                built = original_build(**kwargs)
                corrupt()
                return built

            monkeypatch.setattr(window._requests, "batch", after_build)
        window._workflow.batch(
            file_paths=[first.main.path, second.main.path], target_format="md", action_name="", options={}
        )
        second_key = normalize_path(second.main.path)

        def valid_sibling_finished():
            entry = window._batch_list_vm.get_file_entry(second_key)
            return entry is not None and entry.status == "completed" and not window._execution.busy

        qtbot.waitUntil(valid_sibling_finished, timeout=3000)
        second_entry = window._batch_list_vm.get_file_entry(second_key)
        assert second_entry is not None and second_entry.output_path is not None
        assert Path(second_entry.output_path).read_bytes() == b"second-resource"
        if fault == "intact":
            assert [data for _task, data in plugin.observed] == [b"first-resource", b"second-resource"]
        else:
            first_entry = window._batch_list_vm.get_file_entry(normalize_path(first.main.path))
            assert first_entry is not None and first_entry.status == "failed"
            assert [data for _task, data in plugin.observed] == [b"second-resource"]
            assert plugin.observed[0][0].endswith("-1")
    finally:
        window.close()
        qapp.clipboard().clear()
        qapp.processEvents()
        controller.stop()

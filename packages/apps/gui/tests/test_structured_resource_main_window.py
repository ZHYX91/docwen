"""Real MainWindow request-chain coverage for structured resource groups."""

from __future__ import annotations

import hashlib
import json
from pathlib import Path

import pytest

from docwen_core.detection import inspect_structured_clipboard_snapshot
from docwen_core.models import FILE_INSPECTION_METADATA_KEY
from docwen_core.models.clipboard_document import CLIPBOARD_DOCUMENT_SCHEMA
from docwen_core.models.file_ref import MANAGED_RESOURCE_ID_METADATA_KEY
from docwen_gui.execution_admission import (
    ExecutionAdmission,
    ExecutionAdmissionError,
    check_frozen_request,
)
from docwen_gui.main_window import MainWindow
from docwen_gui.view_models.main_window_vm import MainWindowViewModel
from docwen_runtime.templates import TemplateRegistry

pytestmark = [pytest.mark.gui, pytest.mark.pr_gate, pytest.mark.release_gate]


def _payload(resource_id: str, logical_path: str, resource_bytes: bytes) -> bytes:
    return json.dumps(
        {
            "schema": CLIPBOARD_DOCUMENT_SCHEMA,
            "blocks": [
                {
                    "type": "paragraph",
                    "inlines": [
                        {
                            "type": "image",
                            "resourceId": resource_id,
                            "alt": "opaque resource",
                            "missingReason": "",
                        }
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


@pytest.fixture
def resource_window(qapp, qtbot, tmp_path: Path):
    window = MainWindow(
        view_model=MainWindowViewModel(controller=None),
        clipboard_input_root=tmp_path / "clipboard-inputs",
    )
    qtbot.addWidget(window)
    window.show()
    qapp.processEvents()
    yield window
    window.close()
    qapp.clipboard().clear()
    qapp.processEvents()


def _create_bundle(window: MainWindow, resource_id: str, resource_bytes: bytes):
    store = window._clipboard_store_for_paste()
    logical_path = f"resources/{resource_id}.bin"
    return store.create_bundle(
        _payload(resource_id, logical_path, resource_bytes),
        display_name_template="Clipboard Document {index}.dwclip",
        preview=resource_id,
        resources=((resource_id, logical_path, "application/octet-stream", resource_bytes),),
    )


def _select_docx_template(window: MainWindow) -> str:
    template = next(
        item for item in TemplateRegistry.default().list_templates("docx") if item.name == "English General Template"
    )
    selector = window._template_selector.get_selector("docx")
    assert selector is not None
    selector.select_template(template.id, selection_source="user")
    return template.id


def test_main_window_single_request_carries_resource_without_fake_source_inspection(
    resource_window: MainWindow,
    qapp,
    qtbot,
) -> None:
    bundle = _create_bundle(resource_window, "asset-a", b"first-resource")
    resource_window._input_area_vm.add_files(
        [bundle.main.path],
        file_inspector=inspect_structured_clipboard_snapshot,
    )
    qtbot.waitUntil(lambda: not resource_window.view_model.inspection_busy)
    qtbot.waitUntil(lambda: len(resource_window.view_model.files) == 1)
    template_id = _select_docx_template(resource_window)

    request, _context = resource_window._requests.single(
        file_path=bundle.main.path,
        target_format="docx",
        action_name="",
        options={},
        route_options=("template_name",),
    )

    assert request.options["template_name"] == template_id
    assert len(request.input_refs) == 2
    source, resource = request.input_refs
    assert source.input_role == "source"
    assert resource.input_role == "linked_resource"
    assert resource.metadata[MANAGED_RESOURCE_ID_METADATA_KEY] == "asset-a"
    assert FILE_INSPECTION_METADATA_KEY not in resource.metadata
    assert ExecutionAdmission.pending(request) == []
    check_frozen_request(request)

    Path(resource.path).write_bytes(b"tampered")
    with pytest.raises(ExecutionAdmissionError):
        ExecutionAdmission.pending(request)
    with pytest.raises(ExecutionAdmissionError):
        check_frozen_request(request)


def test_main_window_batch_freezes_two_isolated_document_groups_in_source_order(
    resource_window: MainWindow,
    qapp,
    qtbot,
) -> None:
    resource_window.view_model.set_mode("batch")
    first = _create_bundle(resource_window, "asset-one", b"one-bytes")
    second = _create_bundle(resource_window, "asset-two", b"two-bytes")
    resource_window._input_area_vm.add_files(
        [first.main.path, second.main.path],
        file_inspector=inspect_structured_clipboard_snapshot,
    )
    qtbot.waitUntil(lambda: not resource_window.view_model.inspection_busy)
    qtbot.waitUntil(lambda: len(resource_window.view_model.files) == 2)
    _select_docx_template(resource_window)

    request, context = resource_window._requests.batch(
        file_paths=[first.main.path, second.main.path],
        target_format="docx",
        action_name="",
        options={},
        route_options=("template_name",),
    )

    assert [ref.path for ref in request.input_refs] == [first.main.path, second.main.path]
    assert all(ref.input_role == "source" for ref in request.input_refs)
    groups = context["_document_group_requests"]
    assert [group.request_id for group in groups] == [f"{request.request_id}-0", f"{request.request_id}-1"]
    assert [[ref.input_role for ref in group.input_refs] for group in groups] == [
        ["source", "linked_resource"],
        ["source", "linked_resource"],
    ]
    assert [
        group.input_refs[1].metadata[MANAGED_RESOURCE_ID_METADATA_KEY]
        for group in groups
    ] == ["asset-one", "asset-two"]
    assert groups[0].input_refs[1].path != groups[1].input_refs[1].path

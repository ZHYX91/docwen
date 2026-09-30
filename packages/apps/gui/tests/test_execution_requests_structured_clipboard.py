"""Request identity for managed structured clipboard input."""

from __future__ import annotations

import hashlib
import json
from types import SimpleNamespace
from typing import TYPE_CHECKING, cast

import pytest

from docwen_core.detection import inspect_structured_clipboard_snapshot
from docwen_core.models import FILE_INSPECTION_METADATA_KEY
from docwen_core.models.clipboard_document import (
    CLIPBOARD_DOCUMENT_SCHEMA,
    load_clipboard_document_bytes,
)
from docwen_core.models.file_ref import (
    MANAGED_INPUT_SHA256_METADATA_KEY,
    MANAGED_INPUT_SIZE_BYTES_METADATA_KEY,
    MANAGED_RESOURCE_ID_METADATA_KEY,
    SOURCE_PRESENTATION_NAME_METADATA_KEY,
    FileRef,
)
from docwen_gui.clipboard_inputs import ClipboardInputStore
from docwen_gui.execution_requests import ExecutionRequestBuilder
from docwen_gui.path_identity import normalize_path
from docwen_runtime.templates import TemplateRegistry

if TYPE_CHECKING:
    from docwen_gui.view_models.batch_list_vm import BatchListViewModel
    from docwen_gui.view_models.main_window_vm import MainWindowViewModel


pytestmark = [pytest.mark.integration, pytest.mark.pr_gate, pytest.mark.release_gate]


def test_builder_separates_public_name_virtual_path_and_managed_integrity(tmp_path) -> None:
    source = tmp_path / "opaque.dwclip"
    payload = ('{"blocks":[],"resources":[],"schema":"' + CLIPBOARD_DOCUMENT_SCHEMA + '"}').encode()
    load_clipboard_document_bytes(payload)
    source.write_bytes(payload)
    inspection = inspect_structured_clipboard_snapshot(str(source))
    ref = FileRef(
        path=str(source),
        format="clipboard_document",
        category="markdown",
        metadata={FILE_INSPECTION_METADATA_KEY: inspection.to_dict()},
    )
    template = next(
        item for item in TemplateRegistry.default().list_templates("docx") if item.name == "English General Template"
    )
    builder = ExecutionRequestBuilder(
        cast("MainWindowViewModel", SimpleNamespace(files=[ref], controller=None)),
        cast("BatchListViewModel", SimpleNamespace(get_file_entry=lambda _path: None)),
        file_contexts=lambda: {normalize_path(str(source)): ("clipboard_document", "markdown")},
        selected_template=lambda: ("docx", template.id),
        source_label=lambda _path: "剪贴板文档 1.dwclip",
        synthetic_input=lambda _path: True,
    )

    request, context = builder.single(
        file_path=str(source),
        target_format="docx",
        action_name="",
        options={},
        route_options=("template_name",),
    )

    frozen = request.input_refs[0]
    assert frozen.logical_path == "document.dwclip"
    assert frozen.metadata[SOURCE_PRESENTATION_NAME_METADATA_KEY] == "剪贴板文档 1.dwclip"
    assert frozen.metadata[MANAGED_INPUT_SHA256_METADATA_KEY] == inspection.content_sha256
    assert frozen.metadata[MANAGED_INPUT_SIZE_BYTES_METADATA_KEY] == inspection.size_bytes
    assert request.source_stem == "剪贴板文档 1"
    assert context["display_name"] == "剪贴板文档 1.dwclip"
    assert request.options["template_name"] == template.id


def test_builder_freezes_complete_structured_resource_group_before_request(tmp_path) -> None:
    resource_bytes = b"opaque-resource-bytes"
    resource_sha = hashlib.sha256(resource_bytes).hexdigest()
    payload = json.dumps(
        {
            "schema": CLIPBOARD_DOCUMENT_SCHEMA,
            "blocks": [],
            "resources": [
                {
                    "resourceId": "asset-a",
                    "logicalPath": "resources/asset-a.bin",
                    "mediaType": "application/octet-stream",
                    "sizeBytes": len(resource_bytes),
                    "sha256": resource_sha,
                }
            ],
        },
        separators=(",", ":"),
        sort_keys=True,
    ).encode("utf-8")
    store = ClipboardInputStore(tmp_path / "managed")
    bundle = store.create_bundle(
        payload,
        display_name_template="Clipboard Document {index}.dwclip",
        preview="resource group",
        resources=(
            (
                "asset-a",
                "resources/asset-a.bin",
                "application/octet-stream",
                resource_bytes,
            ),
        ),
    )
    inspection = inspect_structured_clipboard_snapshot(bundle.main.path)
    source_metadata = {
        FILE_INSPECTION_METADATA_KEY: inspection.to_dict(),
        "nested": {"values": ["before"]},
    }
    source_ref = FileRef(
        path=bundle.main.path,
        format="clipboard_document",
        category="markdown",
        metadata=source_metadata,
    )
    view_model = cast(
        "MainWindowViewModel",
        SimpleNamespace(files=[source_ref], controller=None),
    )
    template = next(
        item for item in TemplateRegistry.default().list_templates("docx") if item.name == "English General Template"
    )
    builder = ExecutionRequestBuilder(
        view_model,
        cast("BatchListViewModel", SimpleNamespace(get_file_entry=lambda _path: None)),
        file_contexts=lambda: {normalize_path(bundle.main.path): ("clipboard_document", "markdown")},
        selected_template=lambda: ("docx", template.id),
        source_label=lambda _path: bundle.main.display_name,
        synthetic_input=lambda _path: True,
        snapshot_bundle=store.bundle,
    )

    request, _context = builder.single(
        file_path=bundle.main.path,
        target_format="docx",
        action_name="",
        options={},
        route_options=("template_name",),
    )
    assert len(request.input_refs) == 2
    frozen_source, frozen_resource = request.input_refs
    assert frozen_source.input_role == "source"
    assert frozen_source.metadata["nested"] == {"values": ["before"]}
    assert frozen_resource.input_role == "linked_resource"
    assert frozen_resource.input_kind == "resource"
    assert frozen_resource.logical_path == "resources/asset-a.bin"
    assert frozen_resource.media_type == "application/octet-stream"
    assert frozen_resource.metadata[MANAGED_RESOURCE_ID_METADATA_KEY] == "asset-a"
    assert frozen_resource.metadata[MANAGED_INPUT_SIZE_BYTES_METADATA_KEY] == len(resource_bytes)
    assert frozen_resource.metadata[MANAGED_INPUT_SHA256_METADATA_KEY] == resource_sha
    assert FILE_INSPECTION_METADATA_KEY not in frozen_resource.metadata

    source_metadata["nested"]["values"].append("after")
    view_model.files.clear()
    assert frozen_source.metadata["nested"] == {"values": ["before"]}
    assert len(request.input_refs) == 2
    store.sync_visible([bundle.main.path])
    store.close()

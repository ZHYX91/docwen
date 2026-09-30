"""Admission contracts for managed structured clipboard resources."""

from __future__ import annotations

import hashlib
import json
from pathlib import Path

import pytest

from docwen_core.detection import inspect_structured_clipboard_snapshot
from docwen_core.models import FILE_INSPECTION_METADATA_KEY
from docwen_core.models.clipboard_document import (
    CLIPBOARD_DOCUMENT_MEDIA_TYPE,
    CLIPBOARD_DOCUMENT_SCHEMA,
)
from docwen_core.models.file_ref import (
    MANAGED_INPUT_SHA256_METADATA_KEY,
    MANAGED_INPUT_SIZE_BYTES_METADATA_KEY,
    MANAGED_RESOURCE_ID_METADATA_KEY,
    FileRef,
)
from docwen_core.models.request import ConversionRequest
from docwen_gui.clipboard_inputs import ClipboardInputStore
from docwen_gui.execution_admission import (
    ExecutionAdmission,
    ExecutionAdmissionError,
    check_frozen_request,
)

pytestmark = [pytest.mark.integration, pytest.mark.pr_gate, pytest.mark.release_gate]


def _payload(resource_bytes: bytes) -> bytes:
    return json.dumps(
        {
            "schema": CLIPBOARD_DOCUMENT_SCHEMA,
            "blocks": [],
            "resources": [
                {
                    "resourceId": "asset-a",
                    "logicalPath": "resources/asset-a.bin",
                    "mediaType": "application/octet-stream",
                    "sizeBytes": len(resource_bytes),
                    "sha256": hashlib.sha256(resource_bytes).hexdigest(),
                }
            ],
        },
        separators=(",", ":"),
        sort_keys=True,
    ).encode("utf-8")


def _group(tmp_path: Path) -> tuple[ClipboardInputStore, ConversionRequest]:
    store = ClipboardInputStore(tmp_path / "managed")
    resource_bytes = b"resource-bytes"
    bundle = store.create_bundle(
        _payload(resource_bytes),
        display_name_template="Clipboard Document {index}.dwclip",
        preview="group",
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
    source = FileRef(
        path=bundle.main.path,
        format="clipboard_document",
        category="markdown",
        size_bytes=inspection.size_bytes,
        input_kind="document",
        input_role="source",
        logical_path="document.dwclip",
        media_type=CLIPBOARD_DOCUMENT_MEDIA_TYPE,
        metadata={
            FILE_INSPECTION_METADATA_KEY: inspection.to_dict(),
            MANAGED_INPUT_SIZE_BYTES_METADATA_KEY: inspection.size_bytes,
            MANAGED_INPUT_SHA256_METADATA_KEY: inspection.content_sha256,
        },
    )
    descriptor = bundle.resources[0]
    resource = FileRef(
        path=descriptor.path,
        format="resource",
        category="other",
        size_bytes=descriptor.size_bytes,
        input_kind="resource",
        input_role="linked_resource",
        logical_path=descriptor.logical_path,
        media_type=descriptor.media_type,
        metadata={
            MANAGED_RESOURCE_ID_METADATA_KEY: descriptor.resource_id,
            MANAGED_INPUT_SIZE_BYTES_METADATA_KEY: descriptor.size_bytes,
            MANAGED_INPUT_SHA256_METADATA_KEY: descriptor.sha256,
        },
    )
    request = ConversionRequest(
        request_id="structured-admission",
        input_refs=[source, resource],
        target_format="docx",
    )
    store.sync_visible([bundle.main.path])
    return store, request


def test_linked_resource_uses_own_frozen_integrity_without_source_inspection(tmp_path: Path) -> None:
    store, request = _group(tmp_path)
    try:
        resource = request.input_refs[1]
        assert FILE_INSPECTION_METADATA_KEY not in resource.metadata
        assert ExecutionAdmission.pending(request) == []
        check_frozen_request(request)
    finally:
        store.close()


@pytest.mark.parametrize("mutation", ["delete", "tamper"])
def test_linked_resource_missing_or_tampered_rejects_before_execution(tmp_path: Path, mutation: str) -> None:
    store, request = _group(tmp_path)
    resource = Path(request.input_refs[1].path)
    try:
        if mutation == "delete":
            resource.unlink()
        else:
            resource.write_bytes(b"tampered-bytes")

        with pytest.raises(ExecutionAdmissionError):
            ExecutionAdmission.pending(request)
        with pytest.raises(ExecutionAdmissionError):
            check_frozen_request(request)
    finally:
        store.close()


def test_linked_resource_wrong_size_or_hash_metadata_rejects(tmp_path: Path) -> None:
    store, request = _group(tmp_path)
    try:
        resource = request.input_refs[1]
        resource.metadata[MANAGED_INPUT_SIZE_BYTES_METADATA_KEY] = resource.size_bytes + 1
        resource.metadata[MANAGED_INPUT_SHA256_METADATA_KEY] = "0" * 64

        with pytest.raises(ExecutionAdmissionError):
            ExecutionAdmission.pending(request)
        with pytest.raises(ExecutionAdmissionError):
            check_frozen_request(request)
    finally:
        store.close()


@pytest.mark.parametrize("shape", ["missing", "extra"])
def test_document_resource_set_mismatch_rejects_whole_group(tmp_path: Path, shape: str) -> None:
    store, request = _group(tmp_path)
    try:
        if shape == "missing":
            request.input_refs = request.input_refs[:1]
        else:
            original = request.input_refs[1]
            extra = FileRef(
                path=original.path,
                format="resource",
                category="other",
                size_bytes=original.size_bytes,
                input_kind="resource",
                input_role="linked_resource",
                logical_path="resources/extra.bin",
                media_type=original.media_type,
                metadata={
                    MANAGED_RESOURCE_ID_METADATA_KEY: "asset-extra",
                    MANAGED_INPUT_SIZE_BYTES_METADATA_KEY: original.size_bytes,
                    MANAGED_INPUT_SHA256_METADATA_KEY: original.metadata[MANAGED_INPUT_SHA256_METADATA_KEY],
                },
            )
            request.input_refs.append(extra)

        with pytest.raises(ExecutionAdmissionError):
            ExecutionAdmission.pending(request)
        with pytest.raises(ExecutionAdmissionError):
            check_frozen_request(request)
    finally:
        store.close()

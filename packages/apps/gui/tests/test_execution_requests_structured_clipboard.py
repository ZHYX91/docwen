"""Request identity for managed structured clipboard input."""

from __future__ import annotations

from types import SimpleNamespace

from docwen_core.detection import inspect_structured_clipboard_snapshot
from docwen_core.models import FILE_INSPECTION_METADATA_KEY
from docwen_core.models.clipboard_document import (
    CLIPBOARD_DOCUMENT_SCHEMA,
    clipboard_document_to_bytes,
    load_clipboard_document_bytes,
)
from docwen_core.models.file_ref import (
    MANAGED_INPUT_SHA256_METADATA_KEY,
    MANAGED_INPUT_SIZE_BYTES_METADATA_KEY,
    SOURCE_PRESENTATION_NAME_METADATA_KEY,
    FileRef,
)
from docwen_gui.execution_requests import ExecutionRequestBuilder
from docwen_gui.path_identity import normalize_path


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
    builder = ExecutionRequestBuilder(
        SimpleNamespace(files=[ref], controller=None),
        SimpleNamespace(get_file_entry=lambda _path: None),
        file_contexts=lambda: {normalize_path(str(source)): ("clipboard_document", "markdown")},
        selected_template=lambda: None,
        source_label=lambda _path: "剪贴板文档 1.dwclip",
        synthetic_input=lambda _path: True,
    )

    request, context = builder.single(
        file_path=str(source),
        target_format="md",
        action_name="",
        options={},
        route_options=("markdown_extensions",),
    )

    frozen = request.input_refs[0]
    assert frozen.logical_path == "document.dwclip"
    assert frozen.metadata[SOURCE_PRESENTATION_NAME_METADATA_KEY] == "剪贴板文档 1.dwclip"
    assert frozen.metadata[MANAGED_INPUT_SHA256_METADATA_KEY] == inspection.content_sha256
    assert frozen.metadata[MANAGED_INPUT_SIZE_BYTES_METADATA_KEY] == inspection.size_bytes
    assert request.source_stem == "剪贴板文档 1"
    assert context["display_name"] == "剪贴板文档 1.dwclip"

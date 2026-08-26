"""RuntimePortAdapter is a non-bypassable Core file-admission boundary."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest

from docwen_core.detection import FileAdmissionError, inspect_file
from docwen_core.models.file_inspection import (
    FILE_ADMISSION_ACCEPTANCE_METADATA_KEY,
    FILE_INSPECTION_METADATA_KEY,
    make_admission_acceptance,
)
from docwen_core.models.file_ref import FileRef
from docwen_core.models.request import ConversionRequest
from docwen_core.models.resolved_numbering import (
    NUMBERING_EXPORT_PLAN_MEDIA_TYPE,
    RESOLVED_DOCUMENT_MEDIA_TYPE,
)
from docwen_runtime.adapters import RuntimePortAdapter

pytestmark = pytest.mark.contract


class _RecordingTaskManager:
    def __init__(self) -> None:
        self.single_requests: list[ConversionRequest] = []
        self.batch_requests: list[ConversionRequest] = []

    def execute_single(self, request: ConversionRequest, *, on_event: Any = None) -> ConversionRequest:
        del on_event
        self.single_requests.append(request)
        return request

    def execute_batch(self, request: ConversionRequest, *, on_event: Any = None) -> ConversionRequest:
        del on_event
        self.batch_requests.append(request)
        return request

    def cancel(self, task_id: str) -> None:
        del task_id

    def reserve_cancellation(self, task_id: str) -> None:
        del task_id

    def release_cancellation(self, task_id: str) -> None:
        del task_id

    def cancel_all(self) -> None:
        return None


class _ConfigLoaderThatRejectsProjection:
    @property
    def config(self) -> Any:
        raise AssertionError("configuration must not be projected before file admission")


def _mismatched_pdf_request(path: Path, *, metadata: dict[str, Any] | None = None) -> ConversionRequest:
    return ConversionRequest(
        request_id="direct-runtime-admission",
        input_refs=[
            FileRef(
                path=str(path),
                format="docx",
                category="document",
                metadata=metadata or {},
            )
        ],
        target_format="md",
    )


def _accepted_metadata(path: Path) -> dict[str, Any]:
    inspection = inspect_file(str(path))
    return {
        FILE_INSPECTION_METADATA_KEY: inspection.to_dict(),
        FILE_ADMISSION_ACCEPTANCE_METADATA_KEY: make_admission_acceptance(inspection),
    }


def test_direct_adapter_rejects_wrong_suffix_before_config_or_task_manager(tmp_path: Path) -> None:
    source = tmp_path / "disguised.docx"
    source.write_bytes(b"%PDF-1.4\ncontent\n")
    manager = _RecordingTaskManager()
    config_loader = _ConfigLoaderThatRejectsProjection()
    adapter = RuntimePortAdapter(manager, config_loader=config_loader)  # type: ignore[arg-type]

    with pytest.raises(FileAdmissionError) as exc_info:
        adapter.execute(_mismatched_pdf_request(source))

    assert exc_info.value.error_type == "file_format_confirmation_required"
    assert manager.single_requests == []
    assert manager.batch_requests == []


def test_direct_adapter_forwards_explicitly_accepted_canonical_content(tmp_path: Path) -> None:
    source = tmp_path / "disguised.docx"
    source.write_bytes(b"%PDF-1.4\naccepted\n")
    manager = _RecordingTaskManager()
    adapter = RuntimePortAdapter(manager)  # type: ignore[arg-type]

    result = adapter.execute(_mismatched_pdf_request(source, metadata=_accepted_metadata(source)))

    assert result is manager.single_requests[0]
    admitted_ref = result.input_refs[0]
    assert admitted_ref.format == "pdf"
    assert admitted_ref.category == "layout"
    assert admitted_ref.metadata[FILE_ADMISSION_ACCEPTANCE_METADATA_KEY]["accepted"] is True


@pytest.mark.parametrize(
    "payload",
    [
        "city;value\n北京;1\n上海;2\n".encode(),
        "city,value\n北京,1\n上海,2\n".encode("utf-16"),
    ],
)
def test_direct_adapter_admits_metadata_less_delimited_content(tmp_path: Path, payload: bytes) -> None:
    source = tmp_path / "table.csv"
    source.write_bytes(payload)
    request = ConversionRequest(
        request_id="direct-delimited-ingress",
        input_refs=[FileRef(path=str(source), format="csv", category="spreadsheet")],
        target_format="md",
    )
    manager = _RecordingTaskManager()
    adapter = RuntimePortAdapter(manager)  # type: ignore[arg-type]

    result = adapter.execute(request)

    assert result is manager.single_requests[0]
    admitted_ref = result.input_refs[0]
    assert admitted_ref.format == "csv"
    assert admitted_ref.category == "spreadsheet"
    assert admitted_ref.metadata[FILE_INSPECTION_METADATA_KEY]["decision"] == "allow"


def test_direct_adapter_invalidates_stale_acceptance_after_replacement(tmp_path: Path) -> None:
    source = tmp_path / "disguised.docx"
    source.write_bytes(b"%PDF-1.4\nfirst-content\n")
    request = _mismatched_pdf_request(source, metadata=_accepted_metadata(source))
    source.write_bytes(b"%PDF-1.4\nreplacement-content\n")
    manager = _RecordingTaskManager()
    config_loader = _ConfigLoaderThatRejectsProjection()
    adapter = RuntimePortAdapter(manager, config_loader=config_loader)  # type: ignore[arg-type]

    with pytest.raises(FileAdmissionError) as exc_info:
        adapter.execute(request)

    assert exc_info.value.error_type == "file_format_confirmation_required"
    assert manager.single_requests == []
    assert manager.batch_requests == []


def test_resolved_numbering_pair_is_one_conversion_not_a_batch(tmp_path: Path) -> None:
    neutral = tmp_path / "resolved-document.json"
    plan = tmp_path / "numbering-export-plan.json"
    neutral.write_text("{}", encoding="utf-8")
    plan.write_text("{}", encoding="utf-8")
    request = ConversionRequest(
        request_id="resolved-numbering-pair",
        input_refs=[
            FileRef(
                path=str(neutral),
                format="markdown",
                category="markdown",
                input_kind="document",
                input_role="neutral_document",
                logical_path="request/resolved-document.json",
                media_type=RESOLVED_DOCUMENT_MEDIA_TYPE,
            ),
            FileRef(
                path=str(plan),
                format="resource",
                category="other",
                input_kind="resource",
                input_role="numbering_export_plan",
                logical_path="request/numbering-export-plan.json",
                media_type=NUMBERING_EXPORT_PLAN_MEDIA_TYPE,
            ),
        ],
        target_format="docx",
    )
    manager = _RecordingTaskManager()
    adapter = RuntimePortAdapter(manager)  # type: ignore[arg-type]

    result = adapter.execute(request)

    assert result is manager.single_requests[0]
    assert [item.input_role for item in result.input_refs] == ["neutral_document", "numbering_export_plan"]
    assert manager.batch_requests == []

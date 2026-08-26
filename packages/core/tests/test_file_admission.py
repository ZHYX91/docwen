"""Contract tests for content-first file admission."""

from __future__ import annotations

import os
import sys
import time
import zipfile
from pathlib import Path
from types import SimpleNamespace

import pytest

from docwen_core.detection import FileAdmissionError, FileAdmissionPathError, enforce_file_admission, inspect_file
from docwen_core.models import (
    FILE_ADMISSION_ACCEPTANCE_METADATA_KEY,
    FILE_INSPECTION_METADATA_KEY,
    AdmissionDecision,
    ConversionRequest,
    FileRef,
    FormatRelation,
    StructureStatus,
    admission_is_satisfied,
    make_admission_acceptance,
)
from docwen_core.paths import filesystem_path, windows_utf16_units

pytestmark = pytest.mark.contract


def test_equivalent_alias_is_silently_allowed(tmp_path: Path) -> None:
    source = tmp_path / "photo.jpg"
    source.write_bytes(b"\xff\xd8\xff\xe0\x00\x10JFIF\x00")

    inspection = inspect_file(str(source))

    assert inspection.relation is FormatRelation.EQUIVALENT_ALIAS
    assert inspection.decision is AdmissionDecision.ALLOW
    assert inspection.warning_code == ""


def test_compatible_txt_and_markdown_share_one_workflow(tmp_path: Path) -> None:
    markdown_in_txt = tmp_path / "notes.txt"
    markdown_in_txt.write_text("# Heading\n\nBody.\n", encoding="utf-8")
    text_in_markdown = tmp_path / "notes.md"
    text_in_markdown.write_text("Plain text without Markdown markers.\n", encoding="utf-8")

    inspections = [inspect_file(str(markdown_in_txt)), inspect_file(str(text_in_markdown))]

    assert {item.detected_format for item in inspections} == {"markdown", "txt"}
    assert {item.workflow_category for item in inspections} == {"markdown"}
    assert {item.relation for item in inspections} == {FormatRelation.COMPATIBLE_TEXT}
    assert {item.decision for item in inspections} == {AdmissionDecision.ALLOW_WITH_WARNING}
    assert {item.warning_code for item in inspections} == {"FILE_FORMAT_COMPATIBLE_TEXT"}


def test_same_family_image_mismatch_is_allowed_with_warning(tmp_path: Path) -> None:
    source = tmp_path / "photo.jpg"
    source.write_bytes(b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR")

    inspection = inspect_file(str(source))

    assert inspection.relation is FormatRelation.SAME_FAMILY_MISMATCH
    assert inspection.decision is AdmissionDecision.ALLOW_WITH_WARNING
    assert inspection.warning_code == "FILE_FORMAT_SAME_FAMILY_MISMATCH"


@pytest.mark.parametrize(
    ("name", "content", "detected_format"),
    [
        ("plain.docx", b"ordinary text content\n", "txt"),
        ("rich.txt", b"{\\rtf1\\ansi rich text}", "rtf"),
        ("layout.docx", b"%PDF-1.4\n", "pdf"),
    ],
)
def test_cross_workflow_mismatch_requires_explicit_acceptance(
    tmp_path: Path,
    name: str,
    content: bytes,
    detected_format: str,
) -> None:
    source = tmp_path / name
    source.write_bytes(content)

    inspection = inspect_file(str(source))

    assert inspection.detected_format == detected_format
    assert inspection.relation is FormatRelation.CROSS_FAMILY_MISMATCH
    assert inspection.decision is AdmissionDecision.REQUIRE_EXPLICIT_ACCEPTANCE
    assert inspection.warning_code == "FILE_FORMAT_CROSS_FAMILY_MISMATCH"
    assert inspection.reason_code == "FILE_FORMAT_CONFIRMATION_REQUIRED"


def test_supported_content_with_unknown_extension_requires_acceptance(tmp_path: Path) -> None:
    source = tmp_path / "layout.bin"
    source.write_bytes(b"%PDF-1.4\n")

    inspection = inspect_file(str(source))

    assert inspection.detected_format == "pdf"
    assert inspection.relation is FormatRelation.UNRECOGNIZED_EXTENSION
    assert inspection.decision is AdmissionDecision.REQUIRE_EXPLICIT_ACCEPTANCE
    assert inspection.warning_code == "FILE_EXTENSION_UNSUPPORTED"


def test_ordinary_zip_cannot_pose_as_docx(tmp_path: Path) -> None:
    source = tmp_path / "ordinary.docx"
    with zipfile.ZipFile(source, "w") as package:
        package.writestr("hello.txt", "hello")

    inspection = inspect_file(str(source))

    assert inspection.detected_format == "zip"
    assert inspection.decision is AdmissionDecision.BLOCK
    assert inspection.reason_code == "FILE_CONTAINER_INVALID"


def test_unknown_ole_is_blocked_instead_of_suffix_routed(tmp_path: Path, monkeypatch) -> None:
    """An unrecognized compound document must not inherit a trusted .doc identity."""
    source = tmp_path / "unknown.doc"
    source.write_bytes(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"x" * 64)

    class FakeOleFile:
        def __init__(self, _path: str) -> None:
            pass

        def listdir(self) -> list[list[str]]:
            return [["UnrelatedStream"]]

    monkeypatch.setitem(
        sys.modules,
        "olefile",
        SimpleNamespace(isOleFile=lambda _path: True, OleFileIO=FakeOleFile),
    )

    inspection = inspect_file(str(source))

    assert inspection.detected_format == "ole"
    assert inspection.decision is AdmissionDecision.BLOCK
    assert inspection.reason_code == "FILE_CONTAINER_UNRECOGNIZED"


def test_empty_zip_cannot_pose_as_docx(tmp_path: Path) -> None:
    source = tmp_path / "empty-package.docx"
    with zipfile.ZipFile(source, "w"):
        pass

    inspection = inspect_file(str(source))

    assert inspection.detected_format == "zip"
    assert inspection.decision is AdmissionDecision.BLOCK
    assert inspection.reason_code == "FILE_CONTAINER_INVALID"


@pytest.mark.parametrize(
    ("file_format", "marker"),
    [
        ("docx", "word/document.xml"),
        ("xlsx", "xl/workbook.xml"),
        ("pptx", "ppt/presentation.xml"),
        ("ofd", "OFD.xml"),
        ("epub", "META-INF/container.xml"),
        ("xps", "FixedDocumentSequence.fdseq"),
    ],
)
def test_marker_only_zip_document_package_is_blocked(
    tmp_path: Path,
    file_format: str,
    marker: str,
) -> None:
    source = tmp_path / f"fake.{file_format}"
    with zipfile.ZipFile(source, "w") as package:
        package.writestr(marker, "<marker/>")

    inspection = inspect_file(str(source))

    assert inspection.detected_format == file_format
    assert inspection.structure_status is StructureStatus.INVALID
    assert inspection.decision is AdmissionDecision.BLOCK
    assert inspection.reason_code == "FILE_CONTAINER_INVALID"


def test_truncated_zip_header_is_blocked_without_stalling(tmp_path: Path) -> None:
    source = tmp_path / "truncated.docx"
    source.write_bytes(b"PK\x03\x04")

    started = time.monotonic()
    inspection = inspect_file(str(source))
    elapsed = time.monotonic() - started

    assert elapsed < 1.0
    assert inspection.decision is AdmissionDecision.BLOCK
    assert inspection.reason_code == "FILE_CONTAINER_INVALID"


@pytest.mark.parametrize(
    ("name", "content", "reason_code"),
    [
        ("corrupt.docx", b"PK\x03\x04not-a-valid-central-directory", "FILE_CONTAINER_INVALID"),
        ("unknown.docx", b"\x00\x01\x02\x03\x04\x05", "FILE_CONTENT_UNRECOGNIZED"),
        ("empty.docx", b"", "FILE_EMPTY"),
    ],
)
def test_corrupt_unknown_and_empty_inputs_are_blocked(
    tmp_path: Path,
    name: str,
    content: bytes,
    reason_code: str,
) -> None:
    source = tmp_path / name
    source.write_bytes(content)

    inspection = inspect_file(str(source))

    assert inspection.decision is AdmissionDecision.BLOCK
    assert inspection.reason_code == reason_code


def test_acceptance_record_is_bound_to_the_exact_inspection(tmp_path: Path) -> None:
    source = tmp_path / "layout.docx"
    source.write_bytes(b"%PDF-1.4\n")
    inspection = inspect_file(str(source))
    metadata = {
        FILE_INSPECTION_METADATA_KEY: inspection.to_dict(),
        FILE_ADMISSION_ACCEPTANCE_METADATA_KEY: make_admission_acceptance(inspection),
    }

    assert admission_is_satisfied(inspection, metadata)
    metadata[FILE_ADMISSION_ACCEPTANCE_METADATA_KEY]["detected_format"] = "docx"
    assert not admission_is_satisfied(inspection, metadata)


def test_application_guard_rejects_missing_acceptance_and_normalizes_text(tmp_path: Path) -> None:
    mismatched = tmp_path / "layout.docx"
    mismatched.write_bytes(b"%PDF-1.4\n")
    mismatch_inspection = inspect_file(str(mismatched))
    blocked_request = ConversionRequest(
        request_id="blocked",
        input_refs=[
            FileRef(
                path=str(mismatched),
                format="docx",
                category="document",
                metadata={FILE_INSPECTION_METADATA_KEY: mismatch_inspection.to_dict()},
            )
        ],
        target_format="md",
    )

    with pytest.raises(FileAdmissionError) as exc_info:
        enforce_file_admission(blocked_request)
    assert exc_info.value.error_type == "file_format_confirmation_required"

    text = tmp_path / "plain.md"
    text.write_text("plain content\n", encoding="utf-8")
    text_inspection = inspect_file(str(text))
    allowed_request = ConversionRequest(
        request_id="text",
        input_refs=[
            FileRef(
                path=str(text),
                format="txt",
                category="document",
                metadata={FILE_INSPECTION_METADATA_KEY: text_inspection.to_dict()},
            )
        ],
        target_format="docx",
    )

    admitted = enforce_file_admission(allowed_request)

    assert admitted.input_refs[0].format == "txt"
    assert admitted.input_refs[0].category == "markdown"


def test_application_guard_inspects_metadata_less_public_request(tmp_path: Path) -> None:
    source = tmp_path / "plain.md"
    source.write_text("plain content\n", encoding="utf-8")
    request = ConversionRequest(
        request_id="missing-inspection",
        input_refs=[FileRef(path=str(source), format="md", category="markdown")],
        target_format="docx",
    )

    admitted = enforce_file_admission(request)

    raw = admitted.input_refs[0].metadata[FILE_INSPECTION_METADATA_KEY]
    assert raw["file_path"] == str(source.resolve())
    assert raw["mtime_ns"] == source.stat().st_mtime_ns
    assert admitted.input_refs[0].format == "txt"
    assert admitted.input_refs[0].category == "markdown"


@pytest.mark.skipif(sys.platform != "win32", reason="Windows long-path contract")
def test_application_guard_admits_existing_docx_beyond_max_path(tmp_path: Path) -> None:
    parent = tmp_path
    source = parent / "semantic-output.docx"
    while windows_utf16_units(source) < 280:
        parent /= "long-path-segment"
        source = parent / "semantic-output.docx"

    filesystem_path(parent, force_extended=True).mkdir(parents=True)
    io_source = filesystem_path(source, force_extended=True)
    with zipfile.ZipFile(io_source, "w") as package:
        package.writestr(
            "[Content_Types].xml",
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Override PartName="/word/document.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
            "</Types>",
        )
        package.writestr(
            "_rels/.rels",
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
            'Target="word/document.xml"/>'
            "</Relationships>",
        )
        package.writestr("word/document.xml", "<document/>")

    request = ConversionRequest(
        request_id="long-path-admission",
        input_refs=[FileRef(path=str(source), format="docx", category="document")],
        target_format="md",
    )

    admitted = enforce_file_admission(request)

    inspection = admitted.input_refs[0].metadata[FILE_INSPECTION_METADATA_KEY]
    assert inspection["file_path"] == str(source.resolve(strict=False))
    assert inspection["detected_format"] == "docx"
    assert inspection["decision"] == "allow"


def test_application_guard_preserves_typed_resource_without_sniffing(tmp_path: Path) -> None:
    resource = tmp_path / "bibliography.json"
    resource.write_text('{"schema":"docwen.semantic_bibliography.v1","entries":[]}', encoding="utf-8")
    ref = FileRef(
        path=str(resource),
        format="resource",
        category="other",
        input_kind="resource",
        input_role="bibliography",
        logical_path="bibliography.json",
        media_type="application/vnd.docwen.semantic-bibliography+json",
    )
    request = ConversionRequest(
        request_id="typed-resource",
        input_refs=[ref],
        target_format="docx",
    )

    admitted = enforce_file_admission(request)

    assert admitted is request
    assert admitted.input_refs[0] is ref
    assert admitted.input_refs[0].media_type == "application/vnd.docwen.semantic-bibliography+json"
    assert admitted.input_refs[0].metadata == {}


def test_application_guard_rejects_link_before_content_inspection(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    source = tmp_path / "source.md"
    source.write_text("sentinel", encoding="utf-8")
    request = ConversionRequest(
        request_id="link-before-inspection",
        input_refs=[FileRef(path=str(source), format="md", category="markdown")],
        target_format="docx",
    )
    original_is_symlink = Path.is_symlink

    monkeypatch.setattr(
        Path,
        "is_symlink",
        lambda path: path == source or original_is_symlink(path),
    )

    with pytest.raises(FileAdmissionPathError) as rejected:
        enforce_file_admission(request)

    assert rejected.value.error_type == "input_is_link"
    assert rejected.value.details == {"file": str(source)}


@pytest.mark.parametrize(
    "payload",
    [
        "city;value\n北京;1\n上海;2\n".encode(),
        "city,value\n北京,1\n上海,2\n".encode("utf-16"),
    ],
)
def test_application_guard_admits_metadata_less_delimited_content(tmp_path: Path, payload: bytes) -> None:
    source = tmp_path / "table.csv"
    source.write_bytes(payload)
    inspection = inspect_file(str(source))
    request = ConversionRequest(
        request_id="delimited-ingress",
        input_refs=[FileRef(path=str(source), format="csv", category="spreadsheet")],
        target_format="md",
    )

    admitted = enforce_file_admission(request)

    assert inspection.relation is FormatRelation.EXACT_MATCH
    assert inspection.decision is AdmissionDecision.ALLOW
    assert admitted.input_refs[0].format == "csv"
    assert admitted.input_refs[0].category == "spreadsheet"
    assert admitted.input_refs[0].metadata[FILE_INSPECTION_METADATA_KEY]["decision"] == "allow"


def test_application_guard_invalidates_acceptance_when_file_changes(tmp_path: Path) -> None:
    source = tmp_path / "layout.docx"
    source.write_bytes(b"%PDF-1.4\nfirst\n")
    inspection = inspect_file(str(source))
    metadata = {
        FILE_INSPECTION_METADATA_KEY: inspection.to_dict(),
        FILE_ADMISSION_ACCEPTANCE_METADATA_KEY: make_admission_acceptance(inspection),
    }
    request = ConversionRequest(
        request_id="changed-after-confirmation",
        input_refs=[FileRef(path=str(source), format="pdf", category="layout", metadata=metadata)],
        target_format="md",
    )
    source.write_bytes(b"%PDF-1.4\nreplacement content\n")

    with pytest.raises(FileAdmissionError) as exc_info:
        enforce_file_admission(request)

    assert exc_info.value.error_type == "file_format_confirmation_required"


def test_application_guard_invalidates_same_size_replacement_with_restored_mtime(tmp_path: Path) -> None:
    source = tmp_path / "layout.docx"
    source.write_bytes(b"%PDF-1.4\nfirst-content\n")
    inspection = inspect_file(str(source))
    metadata = {
        FILE_INSPECTION_METADATA_KEY: inspection.to_dict(),
        FILE_ADMISSION_ACCEPTANCE_METADATA_KEY: make_admission_acceptance(inspection),
    }
    request = ConversionRequest(
        request_id="same-size-replacement",
        input_refs=[FileRef(path=str(source), format="pdf", category="layout", metadata=metadata)],
        target_format="md",
    )
    original_mtime_ns = source.stat().st_mtime_ns
    source.write_bytes(b"%PDF-1.4\nother-content\n")
    os.utime(source, ns=(original_mtime_ns, original_mtime_ns))

    replacement = inspect_file(str(source))
    assert replacement.size_bytes == inspection.size_bytes
    assert replacement.mtime_ns == inspection.mtime_ns
    assert replacement.content_sha256 != inspection.content_sha256
    with pytest.raises(FileAdmissionError) as exc_info:
        enforce_file_admission(request)

    assert exc_info.value.error_type == "file_format_confirmation_required"


def test_application_guard_does_not_reuse_acceptance_for_another_path(tmp_path: Path) -> None:
    first = tmp_path / "first.docx"
    second = tmp_path / "second.docx"
    first.write_bytes(b"%PDF-1.4\nfirst\n")
    second.write_bytes(b"%PDF-1.4\nsecond\n")
    inspection = inspect_file(str(first))
    metadata = {
        FILE_INSPECTION_METADATA_KEY: inspection.to_dict(),
        FILE_ADMISSION_ACCEPTANCE_METADATA_KEY: make_admission_acceptance(inspection),
    }
    request = ConversionRequest(
        request_id="different-path",
        input_refs=[FileRef(path=str(second), format="pdf", category="layout", metadata=metadata)],
        target_format="md",
    )

    with pytest.raises(FileAdmissionError) as exc_info:
        enforce_file_admission(request)

    assert exc_info.value.error_type == "file_format_confirmation_required"

"""Presence-only OOXML signature graph contracts."""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest

from docwen_core.detection import (
    OOXML_SIGNATURE_INFO_METADATA_KEY,
    OOXML_SIGNATURE_VALIDATION_UNAVAILABLE,
    freeze_ooxml_signature_info,
    inspect_file,
    inspect_ooxml_signature_graph,
    signature_derived_output_diagnostic,
    signature_validation_diagnostic,
)
from docwen_core.models.file_ref import FileRef
from docwen_core.models.request import ConversionRequest

pytestmark = pytest.mark.contract

_OWNER_PARTS = {
    "docx": "word/document.xml",
    "xlsx": "xl/workbook.xml",
    "pptx": "ppt/presentation.xml",
}
_OWNER_CONTENT_TYPES = {
    "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
    "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
    "pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
}
_OFFICE_DOCUMENT_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
_ORIGIN_REL_TYPE = "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin"
_SIGNATURE_REL_TYPE = "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature"


def _write_ooxml(
    path: Path,
    *,
    state: str,
) -> None:
    extension = path.suffix.lstrip(".")
    owner_part = _OWNER_PARTS[extension]
    content_types = [
        '<Default Extension="xml" ContentType="application/xml"/>',
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
        f'<Override PartName="/{owner_part}" ContentType="{_OWNER_CONTENT_TYPES[extension]}"/>',
    ]
    root_relationships = [
        f'<Relationship Id="office-document" Type="{_OFFICE_DOCUMENT_REL_TYPE}" Target="{owner_part}"/>'
    ]
    entries: dict[str, str | bytes] = {owner_part: "<root/>"}
    if state in {"complete", "suspicious"}:
        root_relationships.append(
            f'<Relationship Id="sig-origin" Type="{_ORIGIN_REL_TYPE}" Target="_xmlsignatures/origin.sigs"/>'
        )
        content_types.extend(
            [
                '<Default Extension="sigs" '
                'ContentType="application/vnd.openxmlformats-package.digital-signature-origin"/>',
                '<Override PartName="/_xmlsignatures/sig1.xml" '
                'ContentType="application/vnd.openxmlformats-package.'
                'digital-signature-xmlsignature+xml"/>',
            ]
        )
        entries["_xmlsignatures/origin.sigs"] = b""
        entries["_xmlsignatures/sig1.xml"] = "<Signature/>"
        if state == "complete":
            entries["_xmlsignatures/_rels/origin.sigs.rels"] = (
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                f'<Relationship Id="sig1" Type="{_SIGNATURE_REL_TYPE}" Target="sig1.xml"/>'
                "</Relationships>"
            )
    entries["_rels/.rels"] = (
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + "".join(root_relationships)
        + "</Relationships>"
    )
    entries["[Content_Types].xml"] = (
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        + "".join(content_types)
        + "</Types>"
    )
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as package:
        for name, payload in entries.items():
            package.writestr(name, payload)


@pytest.mark.parametrize("extension", ["docx", "xlsx", "pptx"])
@pytest.mark.parametrize(
    ("package_state", "expected_state"),
    [("unsigned", "unsigned"), ("complete", "complete"), ("suspicious", "suspicious")],
)
def test_shared_detector_classifies_all_ooxml_owners(
    tmp_path: Path,
    extension: str,
    package_state: str,
    expected_state: str,
) -> None:
    source = tmp_path / f"source.{extension}"
    _write_ooxml(source, state=package_state)

    info = inspect_ooxml_signature_graph(str(source), actual_format=extension)

    assert info.state == expected_state
    assert info.has_signature_material is (package_state != "unsigned")


def test_inspect_file_exposes_typed_presence_warning_without_trust_claim(
    tmp_path: Path,
) -> None:
    source = tmp_path / "signed.docx"
    _write_ooxml(source, state="complete")

    inspection = inspect_file(str(source))

    assert inspection.ooxml_signature["state"] == "complete"
    assert [warning["code"] for warning in inspection.warnings] == [OOXML_SIGNATURE_VALIDATION_UNAVAILABLE]
    warning = inspection.warnings[0]["message"]
    assert "intact and tampered inputs cannot be distinguished" in warning
    assert "did not validate document integrity" in warning
    assert "valid signature" not in warning.lower()
    assert OOXML_SIGNATURE_VALIDATION_UNAVAILABLE in inspection.warning_message


def test_unsigned_inspection_has_no_signature_warning(tmp_path: Path) -> None:
    source = tmp_path / "unsigned.xlsx"
    _write_ooxml(source, state="unsigned")

    inspection = inspect_file(str(source))

    assert inspection.ooxml_signature["state"] == "unsigned"
    assert inspection.warnings == ()
    assert OOXML_SIGNATURE_VALIDATION_UNAVAILABLE not in inspection.warning_message


def test_duplicate_signature_graph_part_is_suspicious(tmp_path: Path) -> None:
    source = tmp_path / "duplicate.docx"
    _write_ooxml(source, state="complete")
    with (
        pytest.warns(UserWarning, match="Duplicate name"),
        zipfile.ZipFile(source, "a") as package,
    ):
        package.writestr("_xmlsignatures/sig1.xml", "<DifferentSignature/>")

    info = inspect_ooxml_signature_graph(str(source), actual_format="docx")

    assert info.state == "suspicious"
    assert "duplicate_signature_graph_part" in info.reason


def test_unreadable_ooxml_does_not_claim_signature_presence(tmp_path: Path) -> None:
    path = tmp_path / "broken.docx"
    path.write_bytes(b"not an OPC package")

    info = inspect_ooxml_signature_graph(str(path), actual_format="docx")

    assert info.state == "suspicious"
    assert info.marker_count == 0
    assert info.has_signature_material is False
    assert signature_validation_diagnostic(info) is None
    assert signature_derived_output_diagnostic(info) is None


def test_signature_graph_uses_explicit_format_not_filename_suffix(tmp_path: Path) -> None:
    source = tmp_path / "signed.docx"
    _write_ooxml(source, state="complete")
    explicit_non_ooxml = inspect_ooxml_signature_graph(str(source), actual_format="pdf")
    misleading = source.with_suffix(".bin")
    source.rename(misleading)

    explicit_ooxml = inspect_ooxml_signature_graph(str(misleading), actual_format="docx")

    assert explicit_ooxml.state == "complete"
    assert explicit_non_ooxml.state == "not_applicable"
    assert explicit_non_ooxml.format == "pdf"


def test_signature_graph_rejects_missing_concrete_format(tmp_path: Path) -> None:
    source = tmp_path / "signed.docx"
    _write_ooxml(source, state="complete")

    with pytest.raises(ValueError, match="concrete admitted format"):
        inspect_ooxml_signature_graph(str(source), actual_format="")


def test_request_admission_freezes_signature_fact_before_source_changes(
    tmp_path: Path,
) -> None:
    source = tmp_path / "source.pptx"
    _write_ooxml(source, state="complete")
    request = ConversionRequest(
        request_id="freeze-signature",
        input_refs=[
            FileRef(
                path=str(source),
                format="pptx",
                category="presentation",
            )
        ],
        target_format="md",
    )

    admitted = freeze_ooxml_signature_info(request)
    _write_ooxml(source, state="unsigned")
    readmitted = freeze_ooxml_signature_info(admitted)

    assert OOXML_SIGNATURE_INFO_METADATA_KEY not in request.input_refs[0].metadata
    assert readmitted.input_refs[0].metadata[OOXML_SIGNATURE_INFO_METADATA_KEY]["state"] == "complete"


def test_request_signature_freeze_does_not_inspect_typed_resource(tmp_path: Path) -> None:
    resource = tmp_path / "bibliography.docx"
    resource.write_text("not an OOXML package", encoding="utf-8")
    ref = FileRef(
        path=str(resource),
        format="docx",
        category="document",
        input_kind="resource",
        input_role="bibliography",
        media_type="application/vnd.docwen.semantic-bibliography+json",
    )
    request = ConversionRequest(request_id="resource", input_refs=[ref], target_format="docx")

    frozen = freeze_ooxml_signature_info(request)

    assert frozen is request
    assert frozen.input_refs[0].metadata == {}

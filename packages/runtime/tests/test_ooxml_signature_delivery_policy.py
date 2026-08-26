"""Runtime delivery-first projection for presence-only OOXML signatures."""

from __future__ import annotations

import hashlib
import zipfile
from pathlib import Path
from typing import Any

import pytest

from docwen_core.detection import (
    OOXML_SIGNATURE_DERIVED_OUTPUT_UNSIGNED,
    OOXML_SIGNATURE_VALIDATION_UNAVAILABLE,
    freeze_ooxml_signature_info,
)
from docwen_core.models.artifact import ARTIFACT_KIND_PRIMARY, ArtifactManifest
from docwen_core.models.file_ref import FileRef
from docwen_core.models.manifest import PluginManifest, RouteSpec
from docwen_core.models.request import ConversionRequest, OutputPolicy
from docwen_core.models.result import ConversionErrorInfo, ConversionResult
from docwen_runtime.engine.route_resolver import RouteResolver
from docwen_runtime.engine.task_manager import TaskManager
from docwen_runtime.output.finalizer import OutputFinalizer
from docwen_runtime.plugin_registry.registry import PluginRegistry
from docwen_runtime.workspace.manager import WorkspaceManager

pytestmark = pytest.mark.integration

_ORIGIN_REL_TYPE = "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin"
_SIGNATURE_REL_TYPE = "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature"


def _write_docx(path: Path, *, signed: bool) -> None:
    entries: dict[str, str | bytes] = {
        "word/document.xml": "<document><text>Hello</text></document>",
    }
    content_types = [
        '<Default Extension="xml" ContentType="application/xml"/>',
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
    ]
    root_relationships = []
    if signed:
        content_types.extend(
            [
                '<Default Extension="sigs" '
                'ContentType="application/vnd.openxmlformats-package.digital-signature-origin"/>',
                '<Override PartName="/_xmlsignatures/sig1.xml" '
                'ContentType="application/vnd.openxmlformats-package.'
                'digital-signature-xmlsignature+xml"/>',
            ]
        )
        root_relationships.append(
            f'<Relationship Id="origin" Type="{_ORIGIN_REL_TYPE}" Target="_xmlsignatures/origin.sigs"/>'
        )
        entries.update(
            {
                "_xmlsignatures/origin.sigs": b"",
                "_xmlsignatures/sig1.xml": "<Signature/>",
                "_xmlsignatures/_rels/origin.sigs.rels": (
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                    f'<Relationship Id="sig1" Type="{_SIGNATURE_REL_TYPE}" Target="sig1.xml"/>'
                    "</Relationships>"
                ),
            }
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


class _DocxToMarkdownPlugin:
    def __init__(self, *, fail: bool = False) -> None:
        self._fail = fail
        self._manifest = PluginManifest(
            plugin_id="signature_delivery_probe",
            name="Signature delivery probe",
            version="1.0",
            description="Deterministic test plugin",
            routes=[RouteSpec(source_format="docx", target_format="md", label="DOCX to Markdown")],
        )

    @property
    def manifest(self) -> PluginManifest:
        return self._manifest

    def can_handle(
        self,
        source_format: str,
        target_format: str,
        action_name: str = "",
    ) -> bool:
        return source_format == "docx" and target_format == "md" and not action_name

    def convert(self, context: Any) -> ConversionResult:
        if self._fail:
            return ConversionResult(
                task_id=context.request.request_id,
                success=False,
                error=ConversionErrorInfo(
                    error_type="conversion_failed",
                    message="probe failure",
                ),
            )
        staging_path = context.workspace.create_artifact_path("primary", ".md")
        Path(staging_path).write_text("# delivered", encoding="utf-8")
        artifact = ArtifactManifest(
            artifact_id="primary",
            kind=ARTIFACT_KIND_PRIMARY,
            staging_path=staging_path,
            suggested_name=f"{Path(context.workspace.input_path).stem}.md",
            is_primary=True,
        )
        return ConversionResult(
            task_id=context.request.request_id,
            success=True,
            artifacts=[artifact],
        )


def _manager(tmp_path: Path, *, fail: bool = False) -> TaskManager:
    registry = PluginRegistry()
    registry.register(_DocxToMarkdownPlugin(fail=fail))
    return TaskManager(
        registry,
        RouteResolver(registry),
        WorkspaceManager(root_dir=str(tmp_path / "runtime")),
        OutputFinalizer(),
    )


def _request(tmp_path: Path, *sources: Path) -> ConversionRequest:
    return ConversionRequest(
        request_id="signature-policy",
        input_refs=[
            FileRef(
                path=str(source),
                format="docx",
                category="document",
                size_bytes=source.stat().st_size,
            )
            for source in sources
        ],
        target_format="md",
        output_policy=OutputPolicy(output_dir=str(tmp_path / "output")),
    )


def test_batch_delivers_signed_and_unsigned_sources_with_truthful_warning_split(
    tmp_path: Path,
) -> None:
    signed = tmp_path / "signed.docx"
    unsigned = tmp_path / "unsigned.docx"
    _write_docx(signed, signed=True)
    _write_docx(unsigned, signed=False)
    source_hashes = {path.name: hashlib.sha256(path.read_bytes()).hexdigest() for path in (signed, unsigned)}

    results = _manager(tmp_path).execute_batch(_request(tmp_path, signed, unsigned))

    assert [result.success for result in results] == [True, True]
    assert [len(result.artifacts) for result in results] == [1, 1]
    assert [
        diagnostic.code for diagnostic in results[0].diagnostics if diagnostic.code.startswith("OOXML_SIGNATURE_")
    ] == [
        OOXML_SIGNATURE_VALIDATION_UNAVAILABLE,
        OOXML_SIGNATURE_DERIVED_OUTPUT_UNSIGNED,
    ]
    assert not any(diagnostic.code.startswith("OOXML_SIGNATURE_") for diagnostic in results[1].diagnostics)
    messages_by_code = {diagnostic.code: diagnostic.message for diagnostic in results[0].diagnostics}
    assert (
        "intact and tampered inputs cannot be distinguished" in messages_by_code[OOXML_SIGNATURE_VALIDATION_UNAVAILABLE]
    )
    assert "derived and unsigned" in messages_by_code[OOXML_SIGNATURE_DERIVED_OUTPUT_UNSIGNED]
    assert {path.name: hashlib.sha256(path.read_bytes()).hexdigest() for path in (signed, unsigned)} == source_hashes


def test_failed_signed_conversion_does_not_claim_a_delivered_unsigned_artifact(
    tmp_path: Path,
) -> None:
    signed = tmp_path / "signed.docx"
    _write_docx(signed, signed=True)

    result = _manager(tmp_path, fail=True).execute_single(_request(tmp_path, signed))

    assert result.success is False
    assert result.artifacts == []
    assert OOXML_SIGNATURE_DERIVED_OUTPUT_UNSIGNED not in {diagnostic.code for diagnostic in result.diagnostics}


def test_signature_diagnostics_select_source_when_resource_is_first(tmp_path: Path) -> None:
    resource = tmp_path / "bibliography.json"
    source = tmp_path / "signed.docx"
    resource.write_text('{"schema":"docwen.semantic_bibliography.v1","entries":[]}', encoding="utf-8")
    _write_docx(source, signed=True)
    request = ConversionRequest(
        request_id="resource-first",
        input_refs=[
            FileRef(
                path=str(resource),
                format="resource",
                category="other",
                input_kind="resource",
                input_role="bibliography",
                media_type="application/vnd.docwen.semantic-bibliography+json",
            ),
            FileRef(
                path=str(source),
                format="docx",
                category="document",
                input_kind="document",
                input_role="source",
                media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            ),
        ],
        target_format="md",
    )

    diagnostics = TaskManager._ooxml_signature_diagnostics(
        freeze_ooxml_signature_info(request),
        delivered_artifact=True,
    )

    assert [diagnostic.code for diagnostic in diagnostics] == [
        OOXML_SIGNATURE_VALIDATION_UNAVAILABLE,
        OOXML_SIGNATURE_DERIVED_OUTPUT_UNSIGNED,
    ]


def test_signature_diagnostics_include_later_signed_source(tmp_path: Path) -> None:
    unsigned = tmp_path / "unsigned.docx"
    signed = tmp_path / "signed.docx"
    _write_docx(unsigned, signed=False)
    _write_docx(signed, signed=True)
    request = freeze_ooxml_signature_info(_request(tmp_path, unsigned, signed))

    diagnostics = TaskManager._ooxml_signature_diagnostics(request, delivered_artifact=True)

    assert [diagnostic.code for diagnostic in diagnostics] == [
        OOXML_SIGNATURE_VALIDATION_UNAVAILABLE,
        OOXML_SIGNATURE_DERIVED_OUTPUT_UNSIGNED,
    ]


def test_runtime_batch_rejects_typed_resource_fanout(tmp_path: Path) -> None:
    resource = tmp_path / "bibliography.json"
    source = tmp_path / "source.docx"
    resource.write_text("{}", encoding="utf-8")
    _write_docx(source, signed=False)
    request = ConversionRequest(
        request_id="typed-batch",
        input_refs=[
            FileRef(
                path=str(resource),
                format="resource",
                category="other",
                input_role="bibliography",
            ),
            FileRef(path=str(source), format="docx", category="document"),
        ],
        target_format="md",
    )

    with pytest.raises(ValueError, match="only independent source"):
        _manager(tmp_path).execute_batch(request)

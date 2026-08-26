"""Format-identity contracts at the Application/Runtime handoff."""

from __future__ import annotations

from pathlib import Path
from unittest.mock import MagicMock

import pytest

from docwen_application.controller import ApplicationController
from docwen_application.ports.runtime import RuntimePort
from docwen_application.preconversion.pre_converter import PreConversionResult
from docwen_core.detection.ooxml_signature import OOXML_SIGNATURE_INFO_METADATA_KEY
from docwen_core.models import (
    FILE_ADMISSION_ACCEPTANCE_METADATA_KEY,
    FILE_INSPECTION_METADATA_KEY,
    ConversionRequest,
    FileRef,
    OutputPolicy,
)
from docwen_core.models.result import ConversionResult

pytestmark = pytest.mark.unit


@pytest.fixture
def admitted_controller(monkeypatch: pytest.MonkeyPatch) -> tuple[ApplicationController, MagicMock]:
    """Start below Core admission so these tests isolate Application semantics."""

    monkeypatch.setattr("docwen_core.detection.enforce_file_admission", lambda request: request)
    runtime = MagicMock(spec=RuntimePort)
    runtime.execute.side_effect = lambda request: ConversionResult(task_id=request.request_id, success=True)
    return ApplicationController(runtime_port=runtime), runtime


def test_preconversion_consumes_frozen_format_instead_of_misleading_suffix(
    admitted_controller: tuple[ApplicationController, MagicMock],
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    from docwen_application.preconversion import pre_converter

    controller, runtime = admitted_controller
    source = tmp_path / "looks-like-text.txt"
    source.write_text("admitted legacy document", encoding="utf-8")
    seen_source_formats: list[str] = []

    def fake_pre_convert(
        input_path: str,
        source_format: str,
        *,
        staging_dir: str,
        **_kwargs: object,
    ) -> PreConversionResult:
        seen_source_formats.append(source_format)
        output = Path(staging_dir) / "converted.docx"
        output.write_text(Path(input_path).read_text(encoding="utf-8"), encoding="utf-8")
        return PreConversionResult(str(output), source_format, "Fake Office")

    monkeypatch.setattr(pre_converter, "pre_convert", fake_pre_convert)
    request = ConversionRequest(
        request_id="format-over-suffix",
        input_refs=[FileRef(path=str(source), format="doc", category="document")],
        target_format="md",
    )

    assert controller.execute_single(request).success is True

    assert seen_source_formats == ["doc"]
    runtime_ref = runtime.execute.call_args.args[0].input_refs[0]
    assert runtime_ref.format == "docx"
    assert runtime_ref.category == "document"
    assert Path(runtime_ref.path).suffix == ".docx"


@pytest.mark.parametrize("source_format", ["doc", "odt", "rtf", "wps"])
def test_document_hub_action_preconverts_before_runtime_dispatch(
    admitted_controller: tuple[ApplicationController, MagicMock],
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
    source_format: str,
) -> None:
    from docwen_application.preconversion import pre_converter

    controller, runtime = admitted_controller
    source = tmp_path / f"legacy.{source_format}"
    source.write_text("legacy", encoding="utf-8")
    seen_source_formats: list[str] = []

    def fake_pre_convert(
        input_path: str,
        admitted_source_format: str,
        *,
        staging_dir: str,
        **_kwargs: object,
    ) -> PreConversionResult:
        seen_source_formats.append(admitted_source_format)
        output = Path(staging_dir) / "proofread-source.docx"
        output.write_text(Path(input_path).read_text(encoding="utf-8"), encoding="utf-8")
        return PreConversionResult(str(output), admitted_source_format, "Fake Office")

    monkeypatch.setattr(pre_converter, "pre_convert", fake_pre_convert)
    request = ConversionRequest(
        request_id=f"proofread-{source_format}",
        input_refs=[FileRef(path=str(source), format=source_format, category="document")],
        target_format="docx",
        action_name="validate",
        output_policy=OutputPolicy(output_path=str(tmp_path / f"reviewed-{source_format}.docx")),
    )

    assert controller.execute_single(request).success is True

    assert seen_source_formats == [source_format]
    runtime_request = runtime.execute.call_args.args[0]
    runtime_ref = runtime_request.input_refs[0]
    assert runtime_request.action_name == "validate"
    assert runtime_request.target_format == "docx"
    assert runtime_request.output_policy.output_path == str(tmp_path / f"reviewed-{source_format}.docx")
    assert runtime_request.output_policy.output_dir is None
    assert runtime_ref.format == "docx"
    assert runtime_ref.category == "document"
    assert Path(runtime_ref.path).suffix == ".docx"


def test_unknown_format_is_rejected_without_reinspection_or_suffix_guessing(
    admitted_controller: tuple[ApplicationController, MagicMock],
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    from docwen_application.preconversion import pre_converter

    controller, runtime = admitted_controller
    source = tmp_path / "looks-like-legacy.rtf"
    source.write_bytes(b"{\\rtf1\\ansi content that could be detected}")
    pre_convert = MagicMock()
    monkeypatch.setattr(pre_converter, "pre_convert", pre_convert)
    monkeypatch.setattr(
        "docwen_core.detection.inspect_file",
        lambda _path: (_ for _ in ()).throw(AssertionError("Application must not reinspect")),
    )
    request = ConversionRequest(
        request_id="unknown-stays-unknown",
        input_refs=[FileRef(path=str(source), format="unknown", category="other")],
        target_format="md",
    )

    with pytest.raises(ValueError, match="concrete admitted format"):
        controller.execute_single(request)

    pre_convert.assert_not_called()
    runtime.execute.assert_not_called()


def test_preconverted_ref_removes_source_admission_identity_and_records_provenance(
    admitted_controller: tuple[ApplicationController, MagicMock],
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    from docwen_application.preconversion import pre_converter

    controller, runtime = admitted_controller
    source = tmp_path / "legacy.doc"
    source.write_text("legacy", encoding="utf-8")
    source_inspection = {
        "file_path": str(source),
        "detected_format": "doc",
        "decision": "allow",
    }

    def fake_pre_convert(
        input_path: str,
        source_format: str,
        *,
        staging_dir: str,
        **_kwargs: object,
    ) -> PreConversionResult:
        output = Path(staging_dir) / "legacy.docx"
        output.write_text(Path(input_path).read_text(encoding="utf-8"), encoding="utf-8")
        return PreConversionResult(str(output), source_format, "Fake Office")

    monkeypatch.setattr(pre_converter, "pre_convert", fake_pre_convert)
    request = ConversionRequest(
        request_id="derived-provenance",
        input_refs=[
            FileRef(
                path=str(source),
                format="doc",
                category="document",
                warning_message="warning about the original source",
                metadata={
                    "consumer_context": "keep",
                    FILE_INSPECTION_METADATA_KEY: source_inspection,
                    FILE_ADMISSION_ACCEPTANCE_METADATA_KEY: {"accepted": True},
                    OOXML_SIGNATURE_INFO_METADATA_KEY: {"source": True},
                },
            )
        ],
        target_format="md",
    )

    assert controller.execute_single(request).success is True

    derived = runtime.execute.call_args.args[0].input_refs[0]
    assert derived.format == "docx"
    assert derived.warning_message == ""
    assert derived.metadata["consumer_context"] == "keep"
    assert FILE_INSPECTION_METADATA_KEY not in derived.metadata
    assert FILE_ADMISSION_ACCEPTANCE_METADATA_KEY not in derived.metadata
    assert OOXML_SIGNATURE_INFO_METADATA_KEY not in derived.metadata
    assert derived.metadata["_docwen_preconversion_source"] == {
        "path": str(source),
        "format": "doc",
        "category": "document",
        "warning_message": "warning about the original source",
        "inspection": source_inspection,
    }

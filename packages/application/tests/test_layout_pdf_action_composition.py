"""PDF operations compose layout normalization before the PDF-only runtime route."""

from __future__ import annotations

from pathlib import Path
from unittest.mock import MagicMock

import pytest

from docwen_application.controller import ApplicationController
from docwen_application.ports.runtime import RuntimePort
from docwen_application.preconversion.chain_resolver import resolve_chain
from docwen_application.preconversion.pre_converter import PreConversionFailure, PreConversionResult
from docwen_core.models.file_ref import FileRef
from docwen_core.models.request import ConversionRequest, OutputPolicy
from docwen_core.models.result import ConversionResult

pytestmark = pytest.mark.unit


@pytest.mark.parametrize("fmt", ["ofd", "xps"])
@pytest.mark.parametrize("action", ["merge_pdfs", "split_pdf"])
def test_layout_action_resolves_through_pdf_hub(fmt: str, action: str) -> None:
    assert resolve_chain(fmt, "pdf", action_name=action) == ["pdf", "pdf"]
    assert resolve_chain(fmt, "pdf") == ["pdf"]


def test_aggregate_normalization_preserves_order_and_source_output_root(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setattr("docwen_core.detection.enforce_file_admission", lambda request: request)
    original = []
    for name in ("first.pdf", "second.ofd", "third.xps"):
        source = tmp_path / name
        source.write_bytes(b"%PDF-placeholder")
        original.append(FileRef(path=str(source), format=source.suffix[1:], category="layout"))
    executed: list[ConversionRequest] = []
    runtime = MagicMock(spec=RuntimePort)

    def execute(request: ConversionRequest) -> ConversionResult:
        executed.append(request)
        return ConversionResult(task_id=request.request_id, success=True)

    runtime.execute.side_effect = execute

    def prepare(source: str, fmt: str, *, staging_dir: str, cancel=None):
        destination = Path(staging_dir) / (Path(source).stem + ".pdf")
        destination.write_bytes(b"%PDF-mocked")
        return PreConversionResult(
            pre_converted_path=str(destination),
            original_source_format=fmt,
            backend="fixture",
        )

    monkeypatch.setattr("docwen_application.preconversion.layout_actions.pre_convert_layout", prepare)
    controller = ApplicationController(runtime_port=runtime)
    result = controller.execute_aggregate(
        ConversionRequest(
            request_id="aggregate-layout",
            input_refs=original,
            target_format="pdf",
            action_name="merge_pdfs",
            output_policy=OutputPolicy(),
        ),
        "merge_pdfs",
    )
    assert result.success
    assert len(executed) == 1
    request = executed[0]
    assert [ref.format for ref in request.input_refs] == ["pdf", "pdf", "pdf"]
    assert [Path(ref.path).stem for ref in request.input_refs] == ["first", "second", "third"]
    assert request.output_policy.output_dir == str(tmp_path)
    assert request.input_refs[0].path == str(tmp_path / "first.pdf")
    assert not Path(request.input_refs[1].path).exists()
    assert not Path(request.input_refs[2].path).exists()


def test_failed_layout_normalization_does_not_run_aggregate(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr("docwen_core.detection.enforce_file_admission", lambda request: request)
    first = tmp_path / "first.pdf"
    second = tmp_path / "second.ofd"
    first.write_bytes(b"%PDF-")
    second.write_bytes(b"bad")
    runtime = MagicMock(spec=RuntimePort)
    monkeypatch.setattr(
        "docwen_application.preconversion.layout_actions.pre_convert_layout",
        lambda *_args, **_kwargs: PreConversionFailure(
            message="malformed OFD",
            error_type="conversion_failed",
            diagnostic_code="LAYOUT-ACTION-PRECONVERSION-FAILED",
        ),
    )
    controller = ApplicationController(runtime_port=runtime)
    result = controller.execute_aggregate(
        ConversionRequest(
            request_id="aggregate-fail",
            input_refs=[
                FileRef(path=str(first), format="pdf", category="layout"),
                FileRef(path=str(second), format="ofd", category="layout"),
            ],
            target_format="pdf",
            action_name="merge_pdfs",
        ),
        "merge_pdfs",
    )
    assert not result.success
    assert result.error is not None
    runtime.execute.assert_not_called()

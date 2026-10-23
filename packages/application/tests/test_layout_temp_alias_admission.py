"""Application-owned layout intermediates must use physical temp paths."""

from __future__ import annotations

import os
import tempfile
from pathlib import Path
from unittest.mock import MagicMock

import pytest

from docwen_application.controller import ApplicationController
from docwen_application.ports.runtime import RuntimePort
from docwen_application.preconversion.pre_converter import PreConversionResult
from docwen_core.models.file_ref import FileRef
from docwen_core.models.request import ConversionRequest
from docwen_core.models.result import ConversionResult

pytestmark = pytest.mark.unit


@pytest.mark.skipif(os.name == "nt", reason="POSIX symlinked system temp aliases")
def test_layout_pdf_action_uses_physical_owned_temp_and_keeps_user_path_guard(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch,
) -> None:
    import fitz

    import docwen_core.detection as detection
    from docwen_core.detection import FileAdmissionPathError, enforce_file_admission

    physical = tmp_path / "physical"
    physical.mkdir()
    alias = tmp_path / "alias"
    alias.symlink_to(physical, target_is_directory=True)
    monkeypatch.setattr(tempfile, "tempdir", str(alias))
    source = physical / "original.xps"
    source.write_bytes(b"xps source bytes")

    # Preserve Core's real admission on derived PDFs.  Only the original XPS
    # transport is a fixture, rather than a fully encoded XPS document.
    def admit(request):
        if request.input_refs[0].format == "xps":
            return request
        return enforce_file_admission(request)

    monkeypatch.setattr(detection, "enforce_file_admission", admit)
    with pytest.raises(FileAdmissionPathError):
        enforce_file_admission(
            ConversionRequest(
                request_id="external-link",
                input_refs=[FileRef(path=str(alias / "original.xps"), format="xps", category="layout")],
                target_format="pdf",
            )
        )

    observed: list[str] = []

    def normalize(_input: str, fmt: str, *, staging_dir: str, cancel=None):
        assert fmt == "xps"
        dest = Path(staging_dir) / "original.pdf"
        assert str(dest).startswith(str(physical))
        with fitz.open() as document:
            document.new_page()
            dest.write_bytes(document.tobytes())
        observed.append(str(dest))
        return PreConversionResult(
            pre_converted_path=str(dest),
            original_source_format="xps",
            backend="fixture",
        )

    monkeypatch.setattr(
        "docwen_application.preconversion.layout_actions.pre_convert_layout",
        normalize,
    )
    runtime = MagicMock(spec=RuntimePort)

    def check_runtime_input(request: ConversionRequest) -> ConversionResult:
        # Assert real Core admission at the runtime boundary, not only the
        # path returned by the preconverter.
        assert request.input_refs[0].format == "pdf"
        enforce_file_admission(request)
        return ConversionResult(task_id=request.request_id, success=True)

    runtime.execute.side_effect = check_runtime_input

    result = ApplicationController(runtime_port=runtime).execute_single(
        ConversionRequest(
            request_id="owned-layout-alias",
            input_refs=[FileRef(path=str(source), format="xps", category="layout")],
            target_format="pdf",
            action_name="split_pdf",
        )
    )
    assert result.success, result.error
    assert len(observed) == 1
    runtime.execute.assert_called_once()
    assert not Path(observed[0]).exists()

"""Application-owned layout intermediates must use physical temp paths."""

from __future__ import annotations

import os
import tempfile
from pathlib import Path
from unittest.mock import MagicMock

import pytest
from tests.support.xps import create_minimal_xps

from docwen_application.controller import ApplicationController
from docwen_application.ports.runtime import RuntimePort
from docwen_core.models.file_ref import FileRef
from docwen_core.models.request import ConversionRequest
from docwen_core.models.result import ConversionResult

pytestmark = [pytest.mark.integration, pytest.mark.release_gate]


@pytest.mark.skipif(os.name == "nt", reason="POSIX symlinked system temp aliases")
def test_layout_pdf_action_uses_physical_owned_temp_and_keeps_user_path_guard(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import fitz

    from docwen_core.detection import FileAdmissionPathError, enforce_file_admission

    physical = tmp_path / "physical"
    physical.mkdir()
    alias = tmp_path / "alias"
    alias.symlink_to(physical, target_is_directory=True)
    monkeypatch.setattr(tempfile, "tempdir", str(alias))
    source = physical / "original.xps"
    create_minimal_xps(source)
    source_bytes = source.read_bytes()
    with pytest.raises(FileAdmissionPathError):
        enforce_file_admission(
            ConversionRequest(
                request_id="external-link",
                input_refs=[FileRef(path=str(alias / "original.xps"), format="xps", category="layout")],
                target_format="pdf",
            )
        )

    observed: list[str] = []

    runtime = MagicMock(spec=RuntimePort)

    def check_runtime_input(request: ConversionRequest) -> ConversionResult:
        # Assert real Core admission at the runtime boundary, not only the
        # path returned by the preconverter.
        assert request.input_refs[0].format == "pdf"
        enforce_file_admission(request)
        derived = Path(request.input_refs[0].path)
        assert derived.is_relative_to(physical)
        with fitz.open(derived) as document:
            assert document.page_count == 2
        observed.append(str(derived))
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
    assert source.read_bytes() == source_bytes

"""ApplicationController admission guard tests."""

from __future__ import annotations

from pathlib import Path
from unittest.mock import MagicMock

import pytest

from docwen_application.controller import ApplicationController
from docwen_application.ports.runtime import RuntimePort
from docwen_core.detection import FileAdmissionError, inspect_file
from docwen_core.models import FILE_INSPECTION_METADATA_KEY, ConversionRequest, FileRef

pytestmark = pytest.mark.unit


def test_controller_rejects_unaccepted_mismatch_before_runtime(tmp_path: Path) -> None:
    source = tmp_path / "layout.docx"
    source.write_bytes(b"%PDF-1.4\n")
    inspection = inspect_file(str(source))
    request = ConversionRequest(
        request_id="admission-block",
        input_refs=[
            FileRef(
                path=str(source),
                format="docx",
                category="document",
                metadata={FILE_INSPECTION_METADATA_KEY: inspection.to_dict()},
            )
        ],
        target_format="md",
    )
    runtime = MagicMock(spec=RuntimePort)
    controller = ApplicationController(runtime_port=runtime)

    with pytest.raises(FileAdmissionError) as exc_info:
        controller.execute_single(request)

    assert exc_info.value.error_type == "file_format_confirmation_required"
    runtime.execute.assert_not_called()

"""services 单元测试。"""

from __future__ import annotations

from pathlib import Path

import pytest

from docwen.errors import InvalidInputError
from docwen.services.requests import ConversionRequest

pytestmark = pytest.mark.unit


def test_request_validate_requires_file_path() -> None:
    req = ConversionRequest(action_type="validate")
    with pytest.raises(InvalidInputError):
        req.validate()


def test_request_validate_requires_existing_file(tmp_path: Path) -> None:
    missing = tmp_path / "missing.docx"
    req = ConversionRequest(file_path=str(missing), action_type="validate")
    with pytest.raises(InvalidInputError):
        req.validate()


def test_request_validate_aggregate_requires_two_inputs(tmp_path: Path) -> None:
    f1 = tmp_path / "a.pdf"
    f1.write_text("x", encoding="utf-8")
    req = ConversionRequest(file_path=str(f1), action_type="merge_pdfs", file_list=[str(f1)])
    with pytest.raises(InvalidInputError):
        req.validate()


def test_request_validate_convert_requires_target(tmp_path: Path) -> None:
    f = tmp_path / "a.docx"
    f.write_text("x", encoding="utf-8")
    req = ConversionRequest(file_path=str(f), action_type=None)
    with pytest.raises(InvalidInputError):
        req.validate()


def test_request_validate_md_allows_ocr_without_extract_image(tmp_path: Path) -> None:
    f = tmp_path / "a.docx"
    f.write_text("x", encoding="utf-8")
    req = ConversionRequest(
        file_path=str(f),
        action_type=None,
        target_format="md",
        options={"extract_ocr": True, "extract_image": False},
    )
    req.validate()

"""Synthetic UTF-8 Markdown admission stays explicit and content-bound."""

from pathlib import Path

import pytest

from docwen_core.detection import (
    FileAdmissionError,
    enforce_file_admission,
    inspect_file,
    inspect_utf8_markdown_snapshot,
)
from docwen_core.models import FILE_INSPECTION_METADATA_KEY
from docwen_core.models.file_inspection import DetectionMethod
from docwen_core.models.file_ref import FileRef
from docwen_core.models.request import ConversionRequest, OutputPolicy

pytestmark = pytest.mark.contract


def _request(path: Path, inspection) -> ConversionRequest:
    return ConversionRequest(
        request_id="synthetic-markdown",
        input_refs=[
            FileRef(
                path=str(path),
                format=inspection.detected_format,
                category=inspection.workflow_category,
                size_bytes=inspection.size_bytes,
                metadata={FILE_INSPECTION_METADATA_KEY: inspection.to_dict()},
            )
        ],
        target_format="csv",
        output_policy=OutputPolicy(output_dir=str(path.parent / "out")),
    )


def test_html_shaped_plain_text_is_markdown_only_for_explicit_synthetic_inspection(tmp_path: Path) -> None:
    source = tmp_path / "clipboard.md"
    source.write_text("<html><body><p>literal clipboard text</p></body></html>", encoding="utf-8")

    ordinary = inspect_file(str(source))
    synthetic = inspect_utf8_markdown_snapshot(str(source))

    assert ordinary.detected_format != "markdown"
    assert synthetic.detected_format == "markdown"
    assert synthetic.workflow_category == "markdown"
    assert synthetic.detection_method is DetectionMethod.SYNTHETIC_MARKDOWN
    assert synthetic.may_execute and not synthetic.requires_explicit_acceptance

    admitted = enforce_file_admission(_request(source, synthetic))
    assert admitted.input_refs[0].format == "markdown"
    assert admitted.input_refs[0].category == "markdown"


def test_synthetic_snapshot_change_after_admission_fails_closed(tmp_path: Path) -> None:
    source = tmp_path / "clipboard.md"
    source.write_text("# Original\n", encoding="utf-8")
    inspection = inspect_utf8_markdown_snapshot(str(source))
    request = _request(source, inspection)

    source.write_text("# Changed after admission\n", encoding="utf-8")

    with pytest.raises(FileAdmissionError, match="changed after admission"):
        enforce_file_admission(request)


def test_synthetic_markdown_requires_utf8_non_whitespace_md_file(tmp_path: Path) -> None:
    invalid = tmp_path / "invalid.md"
    invalid.write_bytes(b"\xff\xfe")
    with pytest.raises(UnicodeDecodeError):
        inspect_utf8_markdown_snapshot(str(invalid))

    blank = tmp_path / "blank.md"
    blank.write_text(" \t\r\n ", encoding="utf-8")
    with pytest.raises(ValueError, match="non-whitespace"):
        inspect_utf8_markdown_snapshot(str(blank))

    wrong_suffix = tmp_path / "clipboard.txt"
    wrong_suffix.write_text("# Markdown text\n", encoding="utf-8")
    with pytest.raises(ValueError, match=r"\.md declaration"):
        inspect_utf8_markdown_snapshot(str(wrong_suffix))

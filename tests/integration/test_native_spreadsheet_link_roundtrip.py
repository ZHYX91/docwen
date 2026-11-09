"""A native external cell link survives the production MD/XLSX/MD route."""

from pathlib import Path

import openpyxl
import pytest

from docwen_core.models.file_ref import FileRef
from docwen_core.models.request import ConversionRequest, OutputPolicy
from tests.integration._round_trip_helper import _primary_path, _run

pytestmark = pytest.mark.integration


def test_native_external_link_roundtrip_preserves_display_and_target(tmp_path: Path, round_trip_runtime) -> None:
    source = tmp_path / "links.md"
    authored = b"| Link |\n| --- |\n| [Site](https://example.com/page?q=1) |\n"
    source.write_bytes(authored)
    forward = _run(
        round_trip_runtime,
        "native-link-forward",
        source,
        source_format="md",
        target_format="xlsx",
        output_dir=tmp_path / "forward",
    )
    workbook_path = _primary_path(forward)
    workbook = openpyxl.load_workbook(workbook_path)
    sheet = workbook.active
    assert sheet is not None
    assert sheet["A2"].value == "Site"
    assert sheet["A2"].hyperlink.target == "https://example.com/page?q=1"
    workbook.close()
    result = round_trip_runtime.execute(
        ConversionRequest(
            request_id="native-link-reverse",
            input_refs=[FileRef(path=str(workbook_path), format="xlsx", category="spreadsheet")],
            target_format="md",
            options={"to_md_keep_images": False, "to_md_enable_ocr": False},
            output_policy=OutputPolicy(output_dir=str(tmp_path / "reverse")),
        )
    )
    markdown = _primary_path(result).read_text(encoding="utf-8")
    assert "[Site](<https://example.com/page?q=1>)" in markdown
    assert source.read_bytes() == authored

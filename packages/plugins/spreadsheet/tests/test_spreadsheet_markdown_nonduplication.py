"""Keep source spreadsheet values and per-sheet OCR references unambiguous."""

from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace

import openpyxl
import pytest

from ._spreadsheet_image_extraction_support import _make_xlsx_with_images
from ._xlsx_to_md_golden_support import _build_fake_context

pytestmark = pytest.mark.contract


@pytest.mark.parametrize("source_format", ["csv", "xlsx"])
def test_disconnected_island_within_table_rectangle_is_emitted_once(tmp_path: Path, source_format: str) -> None:
    from docwen_plugin_spreadsheet.to_markdown.converter import SpreadsheetToMarkdownConverter

    source = tmp_path / f"l-shaped.{source_format}"
    rows = [["H1", "H2", "H3"], ["Row1", "", ""], ["Row2", "", "UNIQUE-VALUE"]]
    if source_format == "csv":
        import csv

        with source.open("w", encoding="utf-8", newline="") as handle:
            csv.writer(handle).writerows(rows)
    else:
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        assert sheet is not None
        for row in rows:
            sheet.append(row)
        workbook.save(source)
        workbook.close()

    staging = tmp_path / f"staging-{source_format}"
    staging.mkdir()
    result = SpreadsheetToMarkdownConverter().convert(
        _build_fake_context(str(source), str(staging), source_format=source_format)
    )
    assert result.success, result.error
    content = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
    assert content.count("UNIQUE-VALUE") == 1


def test_multisheet_ocr_sidecars_have_distinct_names_and_references(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_core.text.ocr import OcrOutcome, OcrStatus
    from docwen_plugin_spreadsheet.to_markdown.converter import SpreadsheetToMarkdownConverter

    source = _make_xlsx_with_images(tmp_path, images_per_sheet=[1, 1])
    monkeypatch.setattr(
        "docwen_core.text.ocr.run_ocr_outcome",
        lambda *_args, **_kwargs: OcrOutcome(OcrStatus.SUCCESS, text="OCR recognized"),
    )
    monkeypatch.setattr(
        "docwen_core.text.table_recognition.enrich_ocr_table_structure",
        lambda _path, outcome, **_kwargs: (
            outcome,
            SimpleNamespace(fallback_required=False, diagnostic_code=""),
        ),
    )

    staging = tmp_path / "staging-ocr"
    staging.mkdir()
    context = _build_fake_context(
        str(source),
        str(staging),
        options={"to_md_enable_ocr": True, "ocr_placement": "image_md"},
    )
    result = SpreadsheetToMarkdownConverter().convert(context)
    assert result.success, result.error
    sidecars = [
        artifact for artifact in context.workspace._artifacts
        if artifact.media_type == "text/markdown" and artifact.metadata.get("ocr")
    ]
    assert len(sidecars) == 2
    names = [artifact.suggested_name for artifact in sidecars]
    assert len(set(names)) == 2
    content = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
    for name in names:
        assert name in content

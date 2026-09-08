"""Internal workbook paths must not become the exported document's title."""

from pathlib import Path

import pytest

from docwen_core.models.document_node import ConversionIdentity
from docwen_plugin_spreadsheet.format_conversion.converter import SmartSheetConverter

from ._csv_xlsx_support import _build_fake_context, _write_workbook

pytestmark = pytest.mark.contract


@pytest.mark.parametrize("source_format", ["xls", "ods", "et"])
def test_legacy_spreadsheet_markdown_keeps_frozen_source_title(tmp_path, monkeypatch, source_format):
    source = tmp_path / f"input.{source_format}"
    source.write_bytes(b"legacy input represented by the bridge fixture")
    hub = tmp_path / "auxiliary_1.xlsx"
    _write_workbook(hub, [["Name", "Value"], ["Example", 42]])
    staging = tmp_path / "staging"
    staging.mkdir()
    context = _build_fake_context(str(source), str(staging), target_format="md")
    identity = ConversionIdentity.create(
        task_id=context.request.request_id,
        source_stem="Quarterly report",
        source_format=source_format,
    )
    context.request.conversion_identity = identity

    def prepare(_self, _context, _path, _format):
        return str(hub), "fixture bridge"

    monkeypatch.setattr(SmartSheetConverter, "_prepare_hub_xlsx", prepare)
    result = SmartSheetConverter().convert(context)

    assert result.success, result.error
    primary = next(artifact for artifact in result.artifacts if artifact.kind == "primary")
    markdown = Path(primary.staging_path).read_text(encoding="utf-8")
    assert "Quarterly report" in markdown
    assert "auxiliary_1" not in markdown
    assert "Example" in markdown and "42" in markdown
    assert context.request.conversion_identity is identity
    assert context.request.input_refs[0].path == str(source)

"""The public CSV→ODS route feeds a faithful real XLSX into its Office boundary."""

from __future__ import annotations

from pathlib import Path
from typing import Any
from zipfile import ZIP_DEFLATED, ZIP_STORED, ZipFile

import pytest
from lxml import etree

from docwen_application.bundle_mapping import build_bundle_draft
from docwen_core.office_bridge import BridgeResult
from docwen_plugin_spreadsheet.format_conversion.converter import SmartSheetConverter
from docwen_runtime.output.artifact_bundle import ArtifactBundleCommitter
from tests.integration.test_spreadsheet_final_content import _primary, _request, _saved_values

pytestmark = [pytest.mark.integration, pytest.mark.pr_gate]

ODS_MEDIA_TYPE = "application/vnd.oasis.opendocument.spreadsheet"
ODS_NAMESPACES = {
    "office": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
    "table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
    "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
}
MANIFEST_NAMESPACE = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"


def _write_minimal_ods(path: str, values: list[list[str]]) -> None:
    """A standards-shaped ODF fixture using only frozen lxml and ZipFile."""
    office = ODS_NAMESPACES["office"]
    table_ns = ODS_NAMESPACES["table"]
    text_ns = ODS_NAMESPACES["text"]
    content = etree.Element(f"{{{office}}}document-content", nsmap=ODS_NAMESPACES)
    content.set(f"{{{office}}}version", "1.2")
    etree.SubElement(content, f"{{{office}}}automatic-styles")
    body = etree.SubElement(content, f"{{{office}}}body")
    spreadsheet = etree.SubElement(body, f"{{{office}}}spreadsheet")
    table = etree.SubElement(spreadsheet, f"{{{table_ns}}}table", {f"{{{table_ns}}}name": "Sheet1"})
    for source_row in values:
        row = etree.SubElement(table, f"{{{table_ns}}}table-row")
        for value in source_row:
            cell = etree.SubElement(row, f"{{{table_ns}}}table-cell", {f"{{{office}}}value-type": "string"})
            etree.SubElement(cell, f"{{{text_ns}}}p").text = value
    manifest = etree.Element(f"{{{MANIFEST_NAMESPACE}}}manifest", nsmap={"manifest": MANIFEST_NAMESPACE})
    manifest.set(f"{{{MANIFEST_NAMESPACE}}}version", "1.2")
    for name, media_type in (("/", ODS_MEDIA_TYPE), ("content.xml", "text/xml")):
        etree.SubElement(
            manifest,
            f"{{{MANIFEST_NAMESPACE}}}file-entry",
            {f"{{{MANIFEST_NAMESPACE}}}full-path": name, f"{{{MANIFEST_NAMESPACE}}}media-type": media_type},
        )
    with ZipFile(path, "w", ZIP_DEFLATED) as package:
        package.writestr("mimetype", ODS_MEDIA_TYPE, compress_type=ZIP_STORED)
        package.writestr("content.xml", etree.tostring(content, encoding="utf-8", xml_declaration=True))
        package.writestr("META-INF/manifest.xml", etree.tostring(manifest, encoding="utf-8", xml_declaration=True))


@pytest.mark.parametrize("encoding", ["utf-8-sig", "utf-16", "gbk"])
def test_public_hub_route_receives_exact_string_matrix_and_publishes_legal_ods(
    tmp_path: Path, round_trip_runtime: Any, monkeypatch: pytest.MonkeyPatch, encoding: str
) -> None:
    expected = [["Name", "Amount"], ["Alice", "7"], ["Bob", "9"]]
    source = tmp_path / "semicolon.csv"
    source.write_bytes("Name;Amount\nAlice;7\nBob;9\n".encode(encoding))
    observed_hub: list[list[list[Any]]] = []

    def office_boundary(
        _self: SmartSheetConverter,
        _context: Any,
        input_path: str,
        target: str,
        output_path: str,
        *,
        cancel: object | None = None,
    ) -> BridgeResult:
        assert target == "ods" and cancel is not None
        # Real saved hub package is reopened and every value/type compared to
        # the independent input matrix. This is an Office-boundary substitute,
        # not an installed Office conversion or an OOXML format-name mock.
        matrix = _saved_values(Path(input_path))
        observed_hub.append(matrix)
        assert matrix == expected
        _write_minimal_ods(output_path, expected)
        return BridgeResult(success=True, output_path=output_path, backend="test-typed-ods-boundary")

    monkeypatch.setattr(SmartSheetConverter, "_convert_xlsx_to_target", office_boundary)
    output = tmp_path / "published"
    result = _request(round_trip_runtime, source, output, "ods")
    assert result.success, result.error
    assert observed_hub == [expected]
    draft = build_bundle_draft(
        profile="single_document",
        output_media_type="application/vnd.oasis.opendocument.spreadsheet",
        artifacts=result.artifacts,
    )
    bundle = ArtifactBundleCommitter().commit(task_id="spreadsheet-hub-content", staging_root=str(output), draft=draft)
    with ZipFile(_primary(bundle, output)) as package:
        assert package.read("mimetype") == ODS_MEDIA_TYPE.encode("ascii")
        assert package.namelist()[0] == "mimetype"
        assert package.getinfo("mimetype").compress_type == ZIP_STORED
        root = etree.fromstring(package.read("content.xml"))
        manifest = etree.fromstring(package.read("META-INF/manifest.xml"))
    assert root.tag == f"{{{ODS_NAMESPACES['office']}}}document-content"
    assert root.get(f"{{{ODS_NAMESPACES['office']}}}version") == "1.2"
    assert {
        (entry.get(f"{{{MANIFEST_NAMESPACE}}}full-path"), entry.get(f"{{{MANIFEST_NAMESPACE}}}media-type"))
        for entry in manifest
    } == {("/", ODS_MEDIA_TYPE), ("content.xml", "text/xml")}
    namespaces = ODS_NAMESPACES
    tables = root.findall("office:body/office:spreadsheet/table:table", namespaces)
    assert len(tables) == 1
    rows = []
    for row in tables[0].findall("table:table-row", namespaces):
        values = []
        for cell in row.findall("table:table-cell", namespaces):
            assert cell.get(f"{{{namespaces['office']}}}value-type") == "string"
            values.append("".join(cell.itertext()))
        rows.append(values)
    assert rows == expected

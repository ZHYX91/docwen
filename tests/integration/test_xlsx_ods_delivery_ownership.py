"""Saved ODS ownership/repetition boundaries through the real publisher.

Only the external Office backend is a fixture. Malformed packages below are
intentional rejection inputs, not claims of native Office interoperability.
"""

from __future__ import annotations

import copy
import hashlib
import xml.etree.ElementTree as ET
from decimal import Decimal
from pathlib import Path
from typing import Any
from zipfile import ZIP_DEFLATED, ZIP_STORED, ZipFile

import pytest
from openpyxl.utils.cell import get_column_letter

from docwen_application.bundle_mapping import build_bundle_draft
from docwen_runtime.output.artifact_bundle import ArtifactBundleCommitter
from tests.integration._xlsx_ods_typed_support import ODS_MEDIA_TYPE, OFFICE, TABLE, ods_cell, ods_row, write_ods
from tests.integration.test_xlsx_ods_typed_publication import _backend, _request, _source

pytestmark = [pytest.mark.integration, pytest.mark.pr_gate]

OFFICE_TAG = "{" + OFFICE + "}"
TABLE_TAG = "{" + TABLE + "}"


def _saved_root(path: Path) -> ET.Element:
    with ZipFile(path) as package:
        return ET.fromstring(package.read("content.xml"))


def _replace_content(path: Path, root: ET.Element) -> None:
    with ZipFile(path) as package:
        parts = {name: package.read(name) for name in package.namelist()}
    parts["content.xml"] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    with ZipFile(path, "w") as package:
        for name, payload in parts.items():
            package.writestr(name, payload, compress_type=ZIP_STORED if name == "mimetype" else ZIP_DEFLATED)


def _owned_numeric_matrix(path: Path) -> dict[str, list[list[Decimal]]]:
    """Independent ODF consumer for these small numeric row-group fixtures."""
    root = _saved_root(path)
    assert root.tag == OFFICE_TAG + "document-content"
    [body] = root.findall(OFFICE_TAG + "body")
    [spreadsheet] = body.findall(OFFICE_TAG + "spreadsheet")
    allowed = {TABLE_TAG + "table-row-group", TABLE_TAG + "table-header-rows", TABLE_TAG + "table-rows"}

    def rows(parent: ET.Element) -> list[ET.Element]:
        result = []
        for child in parent:
            if child.tag == TABLE_TAG + "table-row":
                result.append(child)
            elif child.tag in allowed:
                result.extend(rows(child))
        return result

    result = {}
    for table in spreadsheet.findall(TABLE_TAG + "table"):
        matrix = []
        for row in rows(table):
            values = []
            for cell in row.findall(TABLE_TAG + "table-cell"):
                assert cell.get(OFFICE_TAG + "value-type") == "float"
                count = int(cell.get(TABLE_TAG + "number-columns-repeated", "1"))
                assert 1 <= count <= 512
                values.extend([Decimal(cell.attrib[OFFICE_TAG + "value"])] * count)
            repeat = int(row.get(TABLE_TAG + "number-rows-repeated", "1"))
            assert 1 <= repeat <= 3
            matrix.extend([list(values) for _ in range(repeat)])
        name = table.attrib[TABLE_TAG + "name"]
        assert name not in result
        result[name] = matrix
    return result


def _mutate_owner(path: Path, defect: str) -> None:
    root = _saved_root(path)
    body = root.find(OFFICE_TAG + "body")
    assert body is not None
    spreadsheet = body.find(OFFICE_TAG + "spreadsheet")
    assert spreadsheet is not None
    table = spreadsheet.find(TABLE_TAG + "table")
    assert table is not None
    row = table.find(TABLE_TAG + "table-row")
    assert row is not None
    if defect == "row-in-columns":
        table.remove(row)
        ET.SubElement(table, TABLE_TAG + "table-columns").append(row)
    elif defect == "spreadsheet-outside-body":
        body.remove(spreadsheet)
        styles = ET.Element(OFFICE_TAG + "automatic-styles")
        styles.append(spreadsheet)
        root.insert(0, styles)
    elif defect == "wrong-content-root":
        root.tag = OFFICE_TAG + "document-styles"
    elif defect == "cell-without-row":
        [cell] = row.findall(TABLE_TAG + "table-cell")
        table.remove(row)
        table.append(cell)
    elif defect == "duplicate-sheet":
        spreadsheet.append(copy.deepcopy(table))
    else:
        raise AssertionError(defect)
    _replace_content(path, root)


@pytest.mark.parametrize(
    ("defect", "root_name", "owned_names", "direct_row_counts", "all_values"),
    [
        ("row-in-columns", "document-content", ["Data"], [0], ["30"]),
        ("spreadsheet-outside-body", "document-content", [], [], ["30"]),
        ("wrong-content-root", "document-styles", ["Data"], [1], ["30"]),
        ("cell-without-row", "document-content", ["Data"], [0], ["30"]),
        ("duplicate-sheet", "document-content", ["Data", "Data"], [1, 1], ["30", "30"]),
    ],
)
def test_cached_values_outside_real_ods_ownership_never_publish(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    round_trip_runtime: Any,
    defect: str,
    root_name: str,
    owned_names: list[str],
    direct_row_counts: list[int],
    all_values: list[str],
) -> None:
    source = tmp_path / "linked.xlsx"
    cells = {"Data": {"A1": ("n", "30")}}
    original = _source(source, cells)
    delivery = tmp_path / "unowned.ods"
    write_ods(delivery, [("Data", ods_row(ods_cell("number", "30")))])
    _mutate_owner(delivery, defect)
    root = _saved_root(delivery)
    assert root.tag == OFFICE_TAG + root_name
    owned = root.findall(f"{OFFICE_TAG}body/{OFFICE_TAG}spreadsheet/{TABLE_TAG}table")
    assert [item.get(TABLE_TAG + "name") for item in owned] == owned_names
    assert [len(item.findall(TABLE_TAG + "table-row")) for item in owned] == direct_row_counts
    assert [cell.get(OFFICE_TAG + "value") for cell in root.iter(TABLE_TAG + "table-cell")] == all_values
    _backend(monkeypatch, delivery, cells, tmp_path / "private.xlsx")
    output = tmp_path / "published"
    output.mkdir()
    sentinel = output / "previous.txt"
    sentinel.write_bytes(b"previous")
    result = round_trip_runtime.execute(_request(source, output, "ods-ownership-negative"))
    assert not result.success
    assert result.artifacts == []
    assert "EXTERNAL_LINK_FLATTENING_NOT_DELIVERED" in {item.code for item in result.diagnostics}
    assert list(output.iterdir()) == [sentinel]
    assert sentinel.read_bytes() == b"previous"
    assert source.read_bytes() == original


@pytest.mark.parametrize("wrapper", ["table-rows", "table-row-group", "table-header-rows"])
def test_legal_ods_row_groups_preserve_complete_numeric_matrix_and_publish(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, round_trip_runtime: Any, wrapper: str
) -> None:
    source = tmp_path / "linked.xlsx"
    cells = {"Data": {"A1": ("n", "30"), "B1": ("n", "0"), "A2": ("n", "30"), "B2": ("n", "0")}}
    original = _source(source, cells)
    delivery = tmp_path / "grouped.ods"
    row = ods_row(ods_cell("number", "30") + ods_cell("number", "0"), repeat=2)
    write_ods(delivery, [("Data", f"<table:{wrapper}>{row}</table:{wrapper}>")])
    expected = {"Data": [[Decimal(30), Decimal(0)], [Decimal(30), Decimal(0)]]}
    assert _owned_numeric_matrix(delivery) == expected
    _backend(monkeypatch, delivery, cells, tmp_path / "private.xlsx")
    output = tmp_path / "published"
    request = _request(source, output, "ods-ownership-positive")
    result = round_trip_runtime.execute(request)
    assert result.success, result.error
    bundle = ArtifactBundleCommitter().commit(
        task_id=request.request_id,
        staging_root=str(output),
        draft=build_bundle_draft(
            profile="single_document", output_media_type=ODS_MEDIA_TYPE, artifacts=result.artifacts
        ),
    )
    [entry] = bundle.entries
    artifact = next(item for item in bundle.artifacts if item.artifact_id == entry.artifact_id)
    final = output / artifact.locator
    assert _owned_numeric_matrix(final) == expected
    assert final.read_bytes() == delivery.read_bytes()
    assert artifact.sha256 == hashlib.sha256(final.read_bytes()).hexdigest()
    assert source.read_bytes() == original


def test_wide_repeated_ods_row_retains_every_typed_coordinate_after_publication(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, round_trip_runtime: Any
) -> None:
    width = 256
    source = tmp_path / "linked.xlsx"
    cells = {
        "Data": {
            f"{get_column_letter(column)}{row}": ("n", str(column))
            for row in range(1, 4)
            for column in range(1, width + 1)
        }
    }
    original = _source(source, cells)
    delivery = tmp_path / "wide.ods"
    row = ods_row("".join(ods_cell("number", str(column)) for column in range(1, width + 1)), repeat=3)
    write_ods(delivery, [("Data", row)])
    expected = {"Data": [[Decimal(column) for column in range(1, width + 1)] for _ in range(3)]}
    assert _owned_numeric_matrix(delivery) == expected
    _backend(monkeypatch, delivery, cells, tmp_path / "private.xlsx")
    output = tmp_path / "published"
    request = _request(source, output, "ods-wide-positive")
    result = round_trip_runtime.execute(request)
    assert result.success, result.error
    bundle = ArtifactBundleCommitter().commit(
        task_id=request.request_id,
        staging_root=str(output),
        draft=build_bundle_draft(
            profile="single_document", output_media_type=ODS_MEDIA_TYPE, artifacts=result.artifacts
        ),
    )
    [entry] = bundle.entries
    artifact = next(item for item in bundle.artifacts if item.artifact_id == entry.artifact_id)
    final = output / artifact.locator
    assert _owned_numeric_matrix(final) == expected
    assert final.read_bytes() == delivery.read_bytes()
    assert artifact.sha256 == hashlib.sha256(final.read_bytes()).hexdigest()
    assert source.read_bytes() == original

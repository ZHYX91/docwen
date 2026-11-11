"""OOXML preflight follows typed relationships and bounds allocation inputs."""

from __future__ import annotations

from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

import pytest

from docwen_core.cancellation import CancellationToken
from docwen_core.errors import CancellationRequested
from docwen_core.spreadsheet_package import ordinary_worksheet_parts, preflight_worksheet_allocations
from docwen_core.spreadsheet_projection import NoOrdinaryWorksheetsError, SpreadsheetProjectionLimitError

pytestmark = pytest.mark.contract

_S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_P = "http://schemas.openxmlformats.org/package/2006/relationships"


def _package(path: Path, sheets: list[tuple[str, str, str, str]]) -> dict[str, bytes]:
    names, relationships = [], []
    parts = {}
    for index, (name, target, kind, body) in enumerate(sheets):
        names.append(f'<sheet name="{name}" sheetId="{index + 1}" r:id="rId{index}"/>')
        relationships.append(f'<Relationship Id="rId{index}" Type="{_R}/{kind}" Target="{target}"/>')
        parts["xl/" + target.replace("%20", " ")] = f'<worksheet xmlns="{_S}">{body}</worksheet>'.encode()
    parts["xl/workbook.xml"] = (
        f'<workbook xmlns="{_S}" xmlns:r="{_R}"><sheets>{"".join(names)}</sheets></workbook>'.encode()
    )
    parts["xl/_rels/workbook.xml.rels"] = (
        f'<Relationships xmlns="{_P}">{"".join(relationships)}</Relationships>'.encode()
    )
    _write(path, parts)
    return parts


def _write(path: Path, parts: dict[str, bytes]) -> None:
    with ZipFile(path, "w", compression=ZIP_DEFLATED) as package:
        for name, payload in parts.items():
            package.writestr(name, payload)


def test_inventory_uses_relationships_and_skips_chart_parts(tmp_path: Path) -> None:
    path = tmp_path / "source.xlsx"
    _package(
        path,
        [
            ("Plot", "chartsheets/plot.xml", "chartsheet", '<mergeCells><mergeCell ref="A1:XFD1048576"/></mergeCells>'),
            ("First", "custom/sheet%20a.xml", "worksheet", "<sheetData><row><c/></row></sheetData>"),
            ("Second", "custom/other.xml", "worksheet", "<sheetData><row><c/></row></sheetData>"),
        ],
    )
    with ZipFile(path) as package:
        assert list(ordinary_worksheet_parts(package)) == [
            ("First", "xl/custom/sheet a.xml"),
            ("Second", "xl/custom/other.xml"),
        ]
    preflight_worksheet_allocations(path, cancel_check=lambda: None, limit=2)


@pytest.mark.parametrize("kind", ["stored", "merged"])
def test_cumulative_allocation_limits_reject_before_native_loading(tmp_path: Path, kind: str) -> None:
    path = tmp_path / "source.xlsx"
    body = (
        "<sheetData><row><c/><c/><c/><c/></row></sheetData>"
        if kind == "stored"
        else '<mergeCells><mergeCell ref="A1:B2"/></mergeCells>'
    )
    _package(path, [("First", "a.xml", "worksheet", body), ("Second", "b.xml", "worksheet", body)])
    preflight_worksheet_allocations(path, cancel_check=lambda: None, limit=8)
    with pytest.raises(SpreadsheetProjectionLimitError, match="Second"):
        preflight_worksheet_allocations(path, cancel_check=lambda: None, limit=7)


def test_chart_only_package_has_no_ordinary_output(tmp_path: Path) -> None:
    path = tmp_path / "source.xlsx"
    _package(path, [("Plot", "plot.xml", "chartsheet", "")])
    with pytest.raises(NoOrdinaryWorksheetsError):
        preflight_worksheet_allocations(path, cancel_check=lambda: None)


def test_ambiguous_relationship_ids_cannot_hide_an_unchecked_sheet(tmp_path: Path) -> None:
    path = tmp_path / "source.xlsx"
    parts = _package(
        path,
        [("First", "a.xml", "worksheet", "<sheetData/>"), ("Second", "b.xml", "worksheet", "<sheetData/>")],
    )
    relations = "xl/_rels/workbook.xml.rels"
    parts[relations] = parts[relations].replace(b'Id="rId1"', b'Id="rId0"')
    _write(path, parts)
    with pytest.raises(ValueError, match="duplicate or missing identities"):
        preflight_worksheet_allocations(path, cancel_check=lambda: None)


def test_external_worksheet_target_is_never_opened(tmp_path: Path) -> None:
    path = tmp_path / "source.xlsx"
    parts = _package(path, [("External", "absent.xml", "worksheet", "")])
    parts.pop("xl/absent.xml")
    relations = "xl/_rels/workbook.xml.rels"
    parts[relations] = parts[relations].replace(
        b'Target="absent.xml"', b'TargetMode="External" Target="https://example.test/data.xml"'
    )
    _write(path, parts)
    with pytest.raises(NoOrdinaryWorksheetsError):
        preflight_worksheet_allocations(path, cancel_check=lambda: None)


@pytest.mark.parametrize("encoding", ["utf-8", "utf-16"])
@pytest.mark.parametrize("part", ["xl/workbook.xml", "xl/a.xml"])
def test_dtd_and_entities_are_rejected_in_all_supported_xml_encodings(tmp_path: Path, part: str, encoding: str) -> None:
    path = tmp_path / "source.xlsx"
    parts = _package(path, [("Data", "a.xml", "worksheet", "")])
    original = parts[part].decode()
    parts[part] = ('<!DOCTYPE worksheet [<!ENTITY expanded "untrusted">]>' + original).encode(encoding)
    _write(path, parts)
    with pytest.raises(ValueError, match="document type"):
        preflight_worksheet_allocations(path, cancel_check=lambda: None)


@pytest.mark.parametrize(
    ("constant", "limit", "message"),
    [
        ("_MAX_PACKAGE_MEMBERS", 2, "member-count"),
        ("_MAX_MEMBER_BYTES", 128, "uncompressed-byte"),
        ("_MAX_PACKAGE_BYTES", 128, "uncompressed-byte"),
        ("_MAX_METADATA_BYTES", 32, "metadata"),
    ],
)
def test_package_member_and_decompression_budgets(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, constant: str, limit: int, message: str
) -> None:
    from docwen_core import spreadsheet_package

    path = tmp_path / "source.xlsx"
    _package(path, [("Data", "a.xml", "worksheet", "<sheetData/>" + " " * 256)])
    monkeypatch.setattr(spreadsheet_package, constant, limit)
    with pytest.raises(SpreadsheetProjectionLimitError, match=message):
        preflight_worksheet_allocations(path, cancel_check=lambda: None)


def test_preflight_propagates_cancellation_during_streaming(tmp_path: Path) -> None:
    path = tmp_path / "source.xlsx"
    _package(path, [("Data", "a.xml", "worksheet", "<sheetData><row>" + "<c/>" * 3000 + "</row></sheetData>")])
    token = CancellationToken()
    checks = 0

    def check() -> None:
        nonlocal checks
        checks += 1
        if checks == 3:
            token.cancel("streaming")
        token.check()

    with pytest.raises(CancellationRequested):
        preflight_worksheet_allocations(path, cancel_check=check)
    assert checks == 3

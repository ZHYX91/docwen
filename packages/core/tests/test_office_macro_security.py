"""Pure file checks for the XLM exception to Office AutomationSecurity.

The synthetic records contain only framing and sheet metadata, never executable
macros. These tests do not open an Office application or certify host behavior.
"""

import struct
import zipfile
from pathlib import Path

import pytest

from docwen_core import office_macro_security as security

pytestmark = pytest.mark.contract

_CONTENT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_OFFICE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
_WORKBOOK_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
_SHEET_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"


def _record(kind: int, data: bytes = b"") -> bytes:
    return struct.pack("<HH", kind, len(data)) + data


def _bof(kind: int, version: int = 0x0600) -> bytes:
    return _record(0x0809, struct.pack("<HH", version, kind) + bytes(12 if version == 0x0600 else 4))


def _bound_sheet(offset: int, kind: int = 0, version: int = 0x0600) -> bytes:
    name = b"\x01\x00S" if version == 0x0600 else b"\x01S"
    return _record(0x0085, struct.pack("<IBB", offset, 0, kind) + name)


def _biff_workbook(
    *, version: int = 0x0600, declared_kind: int = 0, actual_kind: int = 0x0010, extra_globals: bytes = b""
) -> bytes:
    globals_size = len(_bof(5, version) + _bound_sheet(0, declared_kind, version) + extra_globals + _record(0x000A))
    return (
        _bof(5, version)
        + _bound_sheet(globals_size, declared_kind, version)
        + extra_globals
        + _record(0x000A)
        + _bof(actual_kind, version)
        + _record(0x000A)
    )


def _ole_workbook(payload: bytes, *, stream_name: str = "Workbook", extra_stream: str | None = None) -> bytes:
    """Minimal CFB v3 with regular 4096-byte streams, avoiding a mini FAT."""
    assert len(payload) <= 4096
    free, end, fat = 0xFFFFFFFF, 0xFFFFFFFE, 0xFFFFFFFD
    header = bytearray(512)
    header[:8] = bytes.fromhex("d0cf11e0a1b11ae1")
    struct.pack_into("<HHHHH", header, 24, 0x003E, 3, 0xFFFE, 9, 6)
    struct.pack_into("<IIIIIIIII", header, 40, 0, 1, 0, 0, 4096, end, 0, end, 0)
    struct.pack_into("<109I", header, 76, 9, *([free] * 108))

    def directory_entry(
        name: str, kind: int, *, child: int = free, right: int = free, start: int = end, size: int = 0
    ) -> bytes:
        entry = bytearray(128)
        encoded_name = (name + "\0").encode("utf-16-le")
        entry[: len(encoded_name)] = encoded_name
        struct.pack_into("<HBBIII", entry, 64, len(encoded_name), kind, 1, free, right, child)
        struct.pack_into("<IQ", entry, 116, start, size)
        return bytes(entry)

    directory = directory_entry("Root Entry", 5, child=1) + directory_entry(
        stream_name, 2, right=2 if extra_stream else free, start=1, size=4096
    )
    if extra_stream:
        directory += directory_entry(extra_stream, 2, start=10, size=4096)
    directory = directory.ljust(512, b"\0")
    fat_entries = [end, 2, 3, 4, 5, 6, 7, 8, end, fat] + [free] * 118
    if extra_stream:
        fat_entries[10:18] = [11, 12, 13, 14, 15, 16, 17, end]
    return (
        bytes(header)
        + directory
        + payload.ljust(4096, b"\0")
        + struct.pack("<128I", *fat_entries)
        + (bytes(4096) if extra_stream else b"")
    )


def _xlsx_entries() -> dict[str, bytes]:
    return {
        "[Content_Types].xml": (
            f'<Types xmlns="{_CONTENT_NS}"><Default Extension="rels" '
            'ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            f'<Override PartName="/xl/workbook.xml" ContentType="{_WORKBOOK_TYPE}"/>'
            f'<Override PartName="/xl/worksheets/sheet1.xml" ContentType="{_SHEET_TYPE}"/></Types>'
        ).encode(),
        "_rels/.rels": (
            f'<Relationships xmlns="{_REL_NS}"><Relationship Id="rId1" '
            f'Type="{_OFFICE_REL}officeDocument" Target="xl/workbook.xml"/></Relationships>'
        ).encode(),
        "xl/workbook.xml": (
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            f'xmlns:r="{_OFFICE_REL[:-1]}"><sheets><sheet name="S" sheetId="1" r:id="rId1"/>'
            "</sheets></workbook>"
        ).encode(),
        "xl/_rels/workbook.xml.rels": (
            f'<Relationships xmlns="{_REL_NS}"><Relationship Id="rId1" '
            f'Type="{_OFFICE_REL}worksheet" Target="worksheets/sheet1.xml"/></Relationships>'
        ).encode(),
        "xl/worksheets/sheet1.xml": (
            b'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>'
        ),
    }


def _write_zip(path: Path, entries: dict[str, bytes]) -> Path:
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name, data in entries.items():
            archive.writestr(name, data)
    return path


@pytest.mark.parametrize("version", [0x0500, 0x0600])
@pytest.mark.parametrize("padding", [b"", bytes(32)])
def test_biff_admits_complete_plain_workbook(version: int, padding: bytes) -> None:
    security._check_biff_workbook(_biff_workbook(version=version) + padding)


@pytest.mark.parametrize("declared_kind,actual_kind", [(1, 0x0010), (0, 0x0040), (1, 0x0040)])
def test_biff_rejects_declared_or_actual_xlm_sheet(declared_kind: int, actual_kind: int) -> None:
    with pytest.raises(security.OfficeMacroPolicyError, match=r"Excel 4\.0"):
        security._check_biff_workbook(_biff_workbook(declared_kind=declared_kind, actual_kind=actual_kind))


def test_biff_rejects_nested_macro_substream() -> None:
    payload = _biff_workbook()[:-4] + _bof(0x0040) + _record(0x000A) + _record(0x000A)
    with pytest.raises(security.OfficeMacroPolicyError, match=r"Excel 4\.0"):
        security._check_biff_workbook(payload)


def test_biff_rejects_encryption_even_without_macro_sheet() -> None:
    with pytest.raises(security.OfficeMacroPolicyError, match="Encrypted"):
        security._check_biff_workbook(_biff_workbook(extra_globals=_record(0x002F, b"\x00\x00")))


@pytest.mark.parametrize(
    "payload",
    [
        b"",
        b"\x09\x08",
        _record(0x0809, b"\x00\x06\x05\x00"),
        _record(0x0809, struct.pack("<HH", 0x0400, 5) + bytes(12)),
        _record(0x0809, struct.pack("<HH", 0x0600, 5) + bytes(4)),
        _bof(0x0010) + _record(0x000A),
        _bof(5) + struct.pack("<HH", 0x0042, 10) + b"short",
        _bof(5) + _record(0x0042, bytes(8225)),
        _biff_workbook()[:-4],
        _biff_workbook()[:-4] + _record(0x000A, b"\x00"),
        _biff_workbook() + _record(0x0042),
        _biff_workbook() + b"\x01",
        _biff_workbook(declared_kind=3),
        _biff_workbook(actual_kind=0x7777),
    ],
)
def test_biff_rejects_unknown_or_incomplete_framing(payload: bytes) -> None:
    with pytest.raises(security.OfficeMacroPolicyError):
        security._check_biff_workbook(payload)


def test_biff_rejects_mismatched_sheet_offset() -> None:
    payload = bytearray(_biff_workbook())
    struct.pack_into("<I", payload, len(_bof(5)) + 4, 123)
    with pytest.raises(security.OfficeMacroPolicyError, match="match"):
        security._check_biff_workbook(bytes(payload))


def test_biff_rejects_duplicate_sheet_declaration() -> None:
    payload = _bof(5) + _bound_sheet(50) + _bound_sheet(50) + _record(0x000A)
    with pytest.raises(security.OfficeMacroPolicyError, match="duplicate"):
        security._check_biff_workbook(payload)


@pytest.mark.parametrize(
    "version,name",
    [
        (0x0600, b"\x00\x00"),
        (0x0600, b"\x20\x00" + b"S" * 32),
        (0x0600, b"\x01"),
        (0x0600, b"\x02\x00S"),
        (0x0600, b"\x01\x01S"),
        (0x0600, b"\x01\x02S"),
        (0x0500, b"\x00"),
        (0x0500, b"\x02S"),
    ],
)
def test_biff_rejects_incomplete_or_invalid_sheet_names(version: int, name: bytes) -> None:
    declaration = struct.pack("<IBB", 0, 0, 0) + name
    sheet_offset = len(_bof(5, version) + _record(0x0085, declaration) + _record(0x000A))
    declaration = struct.pack("<I", sheet_offset) + declaration[4:]
    payload = (
        _bof(5, version) + _record(0x0085, declaration) + _record(0x000A) + _bof(0x0010, version) + _record(0x000A)
    )
    with pytest.raises(security.OfficeMacroPolicyError):
        security._check_biff_workbook(payload)


def test_biff_rejects_substream_version_mismatch() -> None:
    payload = bytearray(_biff_workbook())
    sheet_offset = struct.unpack_from("<I", payload, len(_bof(5)) + 4)[0]
    struct.pack_into("<H", payload, sheet_offset + 4, 0x0500)
    with pytest.raises(security.OfficeMacroPolicyError):
        security._check_biff_workbook(bytes(payload))


@pytest.mark.parametrize("budget,value", [("_MAX_WORKBOOK_BYTES", 16), ("_MAX_BIFF_RECORDS", 4)])
def test_biff_enforces_inspection_budgets(monkeypatch: pytest.MonkeyPatch, budget: str, value: int) -> None:
    monkeypatch.setattr(security, budget, value)
    with pytest.raises(security.OfficeMacroPolicyError, match="budget"):
        security._check_biff_workbook(_biff_workbook())


@pytest.mark.parametrize("stream_name", ["Workbook", "Book"])
def test_ole_reads_actual_legacy_workbook_stream(tmp_path: Path, stream_name: str) -> None:
    path = tmp_path / "input.xls"
    path.write_bytes(_ole_workbook(_biff_workbook(), stream_name=stream_name))
    security.ensure_excel_input_has_no_xlm(path)


def test_ole_macro_signature_takes_precedence_over_csv_suffix(tmp_path: Path) -> None:
    path = tmp_path / "input.csv"
    path.write_bytes(_ole_workbook(_biff_workbook(declared_kind=1)))
    with pytest.raises(security.OfficeMacroPolicyError, match=r"Excel 4\.0"):
        security.ensure_excel_input_has_no_xlm(path)


def test_ole_rejects_encrypted_package_without_workbook(tmp_path: Path) -> None:
    path = tmp_path / "input.xls"
    path.write_bytes(_ole_workbook(b"encryption metadata", stream_name="EncryptionInfo"))
    with pytest.raises(security.OfficeMacroPolicyError):
        security.ensure_excel_input_has_no_xlm(path)


@pytest.mark.parametrize("encrypted_stream", ["EncryptedPackage", "EncryptionInfo"])
def test_ole_rejects_encryption_metadata_alongside_plain_workbook(tmp_path: Path, encrypted_stream: str) -> None:
    path = tmp_path / "input.xls"
    path.write_bytes(_ole_workbook(_biff_workbook(), extra_stream=encrypted_stream))
    with pytest.raises(security.OfficeMacroPolicyError, match="Encrypted"):
        security.ensure_excel_input_has_no_xlm(path)


def test_ole_rejects_incomplete_fat(tmp_path: Path) -> None:
    path = tmp_path / "input.xls"
    path.write_bytes(_ole_workbook(_biff_workbook())[:-512])
    with pytest.raises(security.OfficeMacroPolicyError):
        security.ensure_excel_input_has_no_xlm(path)


def test_ole_checks_declared_stream_size_before_reading(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    path = tmp_path / "input.xls"
    path.write_bytes(_ole_workbook(_biff_workbook()))
    monkeypatch.setattr(security, "_MAX_WORKBOOK_BYTES", 4095)
    with pytest.raises(security.OfficeMacroPolicyError, match="budget"):
        security.ensure_excel_input_has_no_xlm(path)


@pytest.mark.parametrize("suffix", [".xlsx", ".csv"])
def test_ooxml_admits_plain_workbook_by_container_identity(tmp_path: Path, suffix: str) -> None:
    security.ensure_excel_input_has_no_xlm(_write_zip(tmp_path / f"input{suffix}", _xlsx_entries()))


@pytest.mark.parametrize("source", ["part-name", "content-type", "relationship"])
def test_ooxml_rejects_xlm_in_each_metadata_channel(tmp_path: Path, source: str) -> None:
    entries = _xlsx_entries()
    if source == "part-name":
        entries["xl/macrosheets/sheet1.xml"] = b"<x/>"
    elif source == "content-type":
        entries["[Content_Types].xml"] = entries["[Content_Types].xml"].replace(
            b"</Types>",
            b'<Override PartName="/xl/hidden.xml" ContentType="application/vnd.ms-excel.macrosheet+xml"/></Types>',
        )
    else:
        entries["xl/_rels/workbook.xml.rels"] = entries["xl/_rels/workbook.xml.rels"].replace(
            b"/worksheet", b"/xlMacrosheet"
        )
    path = _write_zip(tmp_path / "input.xlsx", entries)
    with pytest.raises(security.OfficeMacroPolicyError, match=r"Excel 4\.0"):
        security.ensure_excel_input_has_no_xlm(path)


@pytest.mark.parametrize("invalid_name", ["XL/workbook.XML", "xl/../hidden.xml", "/hidden.xml", "xl\\hidden.xml"])
def test_ooxml_rejects_ambiguous_package_paths(tmp_path: Path, invalid_name: str) -> None:
    entries = _xlsx_entries()
    entries[invalid_name] = b"<x/>"
    with pytest.raises(security.OfficeMacroPolicyError, match="Ambiguous"):
        security.ensure_excel_input_has_no_xlm(_write_zip(tmp_path / "input.xlsx", entries))


@pytest.mark.parametrize(
    "entry,payload",
    [
        ("[Content_Types].xml", b"<Types>"),
        ("[Content_Types].xml", b"<Types/>"),
        ("[Content_Types].xml", f'<Types xmlns="{_CONTENT_NS}"/>'.encode()),
        ("[Content_Types].xml", b'<!DOCTYPE x [<!ENTITY a "unsafe">]><Types/>'),
        ("xl/_rels/workbook.xml.rels", b"<Relationships>"),
        ("xl/_rels/workbook.xml.rels", b"<Relationships/>"),
    ],
)
def test_ooxml_rejects_unreadable_or_unknown_security_metadata(tmp_path: Path, entry: str, payload: bytes) -> None:
    entries = _xlsx_entries()
    entries[entry] = payload
    with pytest.raises(security.OfficeMacroPolicyError):
        security.ensure_excel_input_has_no_xlm(_write_zip(tmp_path / "input.xlsx", entries))


@pytest.mark.parametrize("encoding", ["utf-16", "utf-32"])
def test_ooxml_rejects_dtd_in_non_utf8_security_metadata(tmp_path: Path, encoding: str) -> None:
    entries = _xlsx_entries()
    declaration = f'<?xml version="1.0" encoding="{encoding}"?><!DOCTYPE Types [<!ENTITY a "inert">]>'
    entries["[Content_Types].xml"] = (declaration + entries["[Content_Types].xml"].decode()).encode(encoding)
    with pytest.raises(security.OfficeMacroPolicyError):
        security.ensure_excel_input_has_no_xlm(_write_zip(tmp_path / "input.xlsx", entries))


def test_ooxml_rejects_missing_declared_workbook(tmp_path: Path) -> None:
    entries = _xlsx_entries()
    del entries["xl/workbook.xml"]
    with pytest.raises(security.OfficeMacroPolicyError, match="unique"):
        security.ensure_excel_input_has_no_xlm(_write_zip(tmp_path / "input.xlsx", entries))


@pytest.mark.parametrize(
    "budget,value", [("_MAX_ZIP_ENTRIES", 4), ("_MAX_XML_BYTES", 16), ("_MAX_TOTAL_XML_BYTES", 500)]
)
def test_ooxml_enforces_metadata_budgets(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, budget: str, value: int
) -> None:
    path = _write_zip(tmp_path / "input.xlsx", _xlsx_entries())
    monkeypatch.setattr(security, budget, value)
    with pytest.raises(security.OfficeMacroPolicyError, match="budget"):
        security.ensure_excel_input_has_no_xlm(path)


def test_ooxml_rejects_encrypted_security_metadata(tmp_path: Path) -> None:
    path = _write_zip(tmp_path / "input.xlsx", _xlsx_entries())
    payload = bytearray(path.read_bytes())
    for signature, flag_offset in [(b"PK\x03\x04", 6), (b"PK\x01\x02", 8)]:
        offset = 0
        while (offset := payload.find(signature, offset)) != -1:
            flags = struct.unpack_from("<H", payload, offset + flag_offset)[0]
            struct.pack_into("<H", payload, offset + flag_offset, flags | 1)
            offset += 4
    path.write_bytes(payload)
    with pytest.raises(security.OfficeMacroPolicyError):
        security.ensure_excel_input_has_no_xlm(path)


@pytest.mark.parametrize("encoding", ["utf-8", "utf-8-sig", "utf-16"])
def test_plain_csv_is_admitted(tmp_path: Path, encoding: str) -> None:
    path = tmp_path / "input.csv"
    path.write_text("name,value\nA,1\n", encoding=encoding)
    security.ensure_excel_input_has_no_xlm(path)


@pytest.mark.parametrize(
    "payload",
    [
        b"ID;P\nC;X1;Y1;K1\nE\n",
        b"ID,name\n1,A\n",
        '<?xml version="1.0"?><Workbook/>'.encode("utf-16"),
        b" \t\n<Workbook/>",
        b"{\\rtf1 inert}",
        b"name,value\nA,\x00\n",
        b" \r\n",
    ],
)
def test_csv_rejects_text_that_may_select_a_different_parser(tmp_path: Path, payload: bytes) -> None:
    path = tmp_path / "input.csv"
    path.write_bytes(payload)
    with pytest.raises(security.OfficeMacroPolicyError):
        security.ensure_excel_input_has_no_xlm(path)


@pytest.mark.parametrize("budget", ["_MAX_CONTAINER_BYTES", "_MAX_WORKBOOK_BYTES"])
def test_csv_checks_size_before_decoding(tmp_path: Path, monkeypatch: pytest.MonkeyPatch, budget: str) -> None:
    path = tmp_path / "input.csv"
    path.write_bytes(b"name,value\nA,1\n")
    monkeypatch.setattr(security, budget, 8)
    with pytest.raises(security.OfficeMacroPolicyError, match="budget"):
        security.ensure_excel_input_has_no_xlm(path)


def test_ods_requires_exact_spreadsheet_mime(tmp_path: Path) -> None:
    path = _write_zip(tmp_path / "input.ods", {"mimetype": b"application/vnd.oasis.opendocument.spreadsheet"})
    security.ensure_excel_input_has_no_xlm(path)


@pytest.mark.parametrize("mime", [b"application/zip", b"application/vnd.oasis.opendocument.text", b"x" * 257])
def test_zip_with_other_mime_is_rejected(tmp_path: Path, mime: bytes) -> None:
    with pytest.raises(security.OfficeMacroPolicyError):
        security.ensure_excel_input_has_no_xlm(_write_zip(tmp_path / "input.ods", {"mimetype": mime}))


@pytest.mark.parametrize("payload", [b"unknown", b"PKbroken", bytes.fromhex("d0cf11e0a1b11ae1")])
def test_unknown_or_truncated_container_is_rejected(tmp_path: Path, payload: bytes) -> None:
    path = tmp_path / "input.xls"
    path.write_bytes(payload)
    with pytest.raises(security.OfficeMacroPolicyError):
        security.ensure_excel_input_has_no_xlm(path)


def test_missing_input_fails_with_macro_policy_reason(tmp_path: Path) -> None:
    with pytest.raises(security.OfficeMacroPolicyError, match="Cannot inspect"):
        security.ensure_excel_input_has_no_xlm(tmp_path / "missing.xlsx")

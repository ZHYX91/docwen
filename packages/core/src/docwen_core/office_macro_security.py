"""Bounded input checks for the Excel 4.0 exception to COM macro disabling.

These checks do not certify a document as safe. They admit only recognized,
unencrypted spreadsheet framing before an owned COM server opens the input
with AutomationSecurity=ForceDisable. Unsupported inputs may use an explicitly
macro-disabled LibreOffice backend instead.
"""

from __future__ import annotations

import struct
import zipfile
from pathlib import Path
from xml.etree import ElementTree

from docwen_core.detection import delimited_text_encodings

_OLE_SIGNATURE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
_MAX_WORKBOOK_BYTES = 64 * 1024 * 1024
_MAX_CONTAINER_BYTES = 128 * 1024 * 1024
_MAX_BIFF_RECORDS = 1_000_000
_MAX_ZIP_ENTRIES = 10_000
_MAX_XML_BYTES = 4 * 1024 * 1024
_MAX_TOTAL_XML_BYTES = 16 * 1024 * 1024
_BOF = 0x0809
_EOF = 0x000A
_FILEPASS = 0x002F
_BOUNDSHEET = 0x0085
_MACRO_SUBSTREAM = 0x0040
_SHEET_TYPES = {0: 0x0010, 2: 0x0020, 6: 0x0006}


class OfficeMacroPolicyError(ValueError):
    """The input cannot be admitted under the external-Office macro policy."""


class _MacroPolicyTreeBuilder(ElementTree.TreeBuilder):
    def doctype(self, name: str, pubid: str | None, system: str | None) -> None:
        # A parser callback covers UTF-16 as well as UTF-8 metadata. A byte
        # search for <!DOCTYPE misses other encodings before entity expansion.
        raise OfficeMacroPolicyError("Spreadsheet security metadata must not declare XML entities.")


def ensure_excel_input_has_no_xlm(path: Path) -> None:
    """Reject XLM, encryption and unrecognized workbook framing before COM."""
    try:
        if path.stat().st_size > _MAX_CONTAINER_BYTES:
            raise OfficeMacroPolicyError("Spreadsheet container exceeds the macro inspection budget.")
        with path.open("rb") as stream:
            signature = stream.read(8)
        if signature == _OLE_SIGNATURE:
            _check_ole_workbook(path)
        elif signature.startswith(b"PK"):
            _check_zip_workbook(path)
        elif path.suffix.casefold() == ".csv":
            _check_csv_text(path)
        else:
            raise OfficeMacroPolicyError("Unrecognized spreadsheet framing; cannot exclude Excel 4.0 macros.")
    except OfficeMacroPolicyError:
        raise
    except (ImportError, OSError, ValueError, RuntimeError, zipfile.BadZipFile, ElementTree.ParseError) as exc:
        raise OfficeMacroPolicyError(f"Cannot inspect spreadsheet macro policy: {exc}") from exc


def _check_csv_text(path: Path) -> None:
    # Excel chooses parsers from content even when the filename says CSV.
    # In particular, SYLK beginning with ID may contain Excel 4.0 macros.
    if path.stat().st_size > _MAX_WORKBOOK_BYTES:
        raise OfficeMacroPolicyError("CSV exceeds the macro inspection budget.")
    payload = path.read_bytes()
    for encoding in delimited_text_encodings(str(path)):
        try:
            text = payload.decode(encoding, errors="strict")
        except UnicodeDecodeError:
            continue
        if any(ord(character) < 32 and character not in "\t\r\n" for character in text):
            raise OfficeMacroPolicyError("Binary control characters are not admitted as CSV to COM.")
        start = text.lstrip("\ufeff \t\r\n")
        if not start or start.startswith(("<", "{", "ID")):
            raise OfficeMacroPolicyError("Ambiguous CSV content may select another Office parser.")
        return
    raise OfficeMacroPolicyError("CSV cannot be decoded as strictly validated text.")


def _check_ole_workbook(path: Path) -> None:
    import olefile

    with olefile.OleFileIO(str(path), raise_defects=olefile.DEFECT_INCORRECT) as container:
        streams = container.listdir()
        if any(parts[-1].casefold() in {"encryptioninfo", "encryptedpackage"} for parts in streams if parts):
            raise OfficeMacroPolicyError("Encrypted Office containers cannot be inspected for Excel 4.0 macros.")
        candidates = [parts for parts in streams if len(parts) == 1 and parts[0].casefold() in {"workbook", "book"}]
        if len(candidates) != 1:
            raise OfficeMacroPolicyError("A unique legacy workbook stream is required for COM conversion.")
        name = candidates[0]
        stream_size = container.get_size(name)
        if stream_size > _MAX_WORKBOOK_BYTES:
            raise OfficeMacroPolicyError("Legacy workbook exceeds the macro inspection budget.")
        with container.openstream(name) as stream:
            payload = stream.read(_MAX_WORKBOOK_BYTES + 1)
        parsing_issues = getattr(container, "parsing_issues", None)
        if parsing_issues is None or parsing_issues or len(payload) != stream_size:
            raise OfficeMacroPolicyError("Malformed legacy Office container cannot be inspected safely.")
    _check_biff_workbook(payload)


def _check_biff_workbook(payload: bytes) -> None:
    """Inspect BIFF5/8 records, including declared and actual sheet kinds.

    MS-XLS BoundSheet8.dt=1 and BOF.dt=0x0040 denote macro sheets.
    FilePass and unfamiliar framing are rejected, not interpreted as no macros.
    """
    if not payload or len(payload) > _MAX_WORKBOOK_BYTES:
        raise OfficeMacroPolicyError("Legacy workbook exceeds the macro inspection budget or is empty.")
    offset = 0
    record_count = 0
    stack: list[int] = []
    declared_sheets: dict[int, int] = {}
    actual_sheets: dict[int, int] = {}
    globals_finished = False
    globals_version: int | None = None
    while offset < len(payload):
        # OLE workbook streams can include zero padding after their final EOF.
        if not stack and globals_finished and payload[offset : offset + 4] == b"\0\0\0\0":
            if any(payload[offset:]):
                raise OfficeMacroPolicyError("Nonzero data follows legacy workbook padding.")
            break
        if len(payload) - offset < 4:
            raise OfficeMacroPolicyError("Truncated legacy workbook record header.")
        record, length = struct.unpack_from("<HH", payload, offset)
        end = offset + 4 + length
        record_count += 1
        if record_count > _MAX_BIFF_RECORDS or length > 8224 or end > len(payload):
            raise OfficeMacroPolicyError("Invalid legacy workbook record length or inspection budget exceeded.")
        data = payload[offset + 4 : end]
        if record == _FILEPASS:
            raise OfficeMacroPolicyError("Encrypted legacy workbooks cannot be inspected for Excel 4.0 macros.")
        if record == _BOF:
            if length < 8:
                raise OfficeMacroPolicyError("Incomplete legacy workbook BOF.")
            version, kind = struct.unpack_from("<HH", data)
            if version not in {0x0500, 0x0600} or (version == 0x0600 and length < 16):
                raise OfficeMacroPolicyError("Only complete BIFF5/8 framing is admitted to Excel COM.")
            if kind == _MACRO_SUBSTREAM:
                raise OfficeMacroPolicyError("Excel 4.0 macro sheets are not opened by COM conversion.")
            if offset == 0:
                if kind != 0x0005:
                    raise OfficeMacroPolicyError("Legacy workbook must begin with its globals substream.")
                globals_version = version
            elif kind == 0x0005 or not globals_finished or kind not in _SHEET_TYPES.values():
                raise OfficeMacroPolicyError("Unrecognized legacy workbook substream.")
            elif version != globals_version:
                raise OfficeMacroPolicyError("Legacy workbook substreams use inconsistent BIFF versions.")
            if not stack and offset:
                actual_sheets[offset] = kind
            stack.append(kind)
        elif not stack:
            raise OfficeMacroPolicyError("Legacy workbook record is outside a substream.")
        elif record == _EOF:
            if length:
                raise OfficeMacroPolicyError("Invalid legacy workbook EOF record.")
            if stack.pop() == 0x0005:
                globals_finished = True
        elif record == _BOUNDSHEET:
            if stack != [0x0005] or length < 7:
                raise OfficeMacroPolicyError("Invalid legacy workbook sheet declaration.")
            name_length = data[6]
            if not 1 <= name_length <= 31:
                raise OfficeMacroPolicyError("Invalid legacy workbook sheet name length.")
            if globals_version == 0x0600:
                if length < 8 or data[7] & ~1:
                    raise OfficeMacroPolicyError("Invalid legacy workbook sheet name encoding.")
                expected_length = 8 + name_length * (2 if data[7] & 1 else 1)
            else:
                expected_length = 7 + name_length
            if length != expected_length:
                raise OfficeMacroPolicyError("Incomplete legacy workbook sheet name.")
            sheet_offset = struct.unpack_from("<I", data)[0]
            sheet_type = data[5]
            if sheet_type == 1:
                raise OfficeMacroPolicyError("Excel 4.0 macro sheets are not opened by COM conversion.")
            if sheet_type not in _SHEET_TYPES or sheet_offset in declared_sheets:
                raise OfficeMacroPolicyError("Unrecognized or duplicate legacy workbook sheet declaration.")
            declared_sheets[sheet_offset] = _SHEET_TYPES[sheet_type]
        offset = end
    if stack or not globals_finished or not declared_sheets or declared_sheets != actual_sheets:
        raise OfficeMacroPolicyError("Legacy workbook sheets do not match complete declared substreams.")


def _check_zip_workbook(path: Path) -> None:
    with zipfile.ZipFile(path) as archive:
        entries = archive.infolist()
        if len(entries) > _MAX_ZIP_ENTRIES:
            raise OfficeMacroPolicyError("Spreadsheet package exceeds the macro inspection budget.")
        names: dict[str, zipfile.ZipInfo] = {}
        for entry in entries:
            name = entry.filename
            key = name.casefold()
            if key in names or "\\" in name or name.startswith("/") or ".." in name.split("/"):
                raise OfficeMacroPolicyError("Ambiguous spreadsheet package paths cannot be inspected safely.")
            names[key] = entry
            if "macrosheets" in name.casefold().split("/"):
                raise OfficeMacroPolicyError("Excel 4.0 macro sheets are not opened by COM conversion.")

        def read_xml(entry: zipfile.ZipInfo) -> ElementTree.Element:
            nonlocal xml_bytes
            xml_bytes += entry.file_size
            if entry.file_size > _MAX_XML_BYTES or xml_bytes > _MAX_TOTAL_XML_BYTES:
                raise OfficeMacroPolicyError("Spreadsheet XML exceeds the macro inspection budget.")
            payload = archive.read(entry)
            return ElementTree.fromstring(payload, parser=ElementTree.XMLParser(target=_MacroPolicyTreeBuilder()))

        xml_bytes = 0
        if "[content_types].xml" not in names:
            # ODS has no Excel 4.0 sheet type; VBA-style automation remains
            # disabled in the COM server. Do not accept arbitrary ZIP inputs.
            mime = names.get("mimetype")
            if mime is None or mime.file_size > 256:
                raise OfficeMacroPolicyError("Unrecognized spreadsheet package.")
            if archive.read(mime) != b"application/vnd.oasis.opendocument.spreadsheet":
                raise OfficeMacroPolicyError("Unrecognized spreadsheet package.")
            return
        content_types = read_xml(names["[content_types].xml"])
        if content_types.tag != "{http://schemas.openxmlformats.org/package/2006/content-types}Types":
            raise OfficeMacroPolicyError("Invalid spreadsheet content type declarations.")
        workbook_types = {
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
            "application/vnd.ms-excel.sheet.macroenabled.main+xml",
        }
        workbooks: list[str] = []
        for item in content_types:
            content_type = item.get("ContentType", "").casefold()
            if "macrosheet" in content_type:
                raise OfficeMacroPolicyError("Excel 4.0 macro sheets are not opened by COM conversion.")
            if content_type in workbook_types:
                workbooks.append(item.get("PartName", "").lstrip("/").casefold())
        if len(workbooks) != 1 or workbooks[0] not in names:
            raise OfficeMacroPolicyError("A unique XML workbook is required for COM conversion.")
        for name, entry in names.items():
            if name.endswith(".rels"):
                relationships = read_xml(entry)
                if relationships.tag != "{http://schemas.openxmlformats.org/package/2006/relationships}Relationships":
                    raise OfficeMacroPolicyError("Invalid spreadsheet relationship declarations.")
                if any("macrosheet" in relation.get("Type", "").casefold() for relation in relationships):
                    raise OfficeMacroPolicyError("Excel 4.0 macro sheets are not opened by COM conversion.")

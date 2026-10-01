"""Native-provider derived samples and bounded CFB admission regressions."""

from __future__ import annotations

import hashlib
import io
import json
import struct
import zipfile
from pathlib import Path

import pytest
from PySide6.QtCore import QMimeData

from docwen_core.models.clipboard_document import ClipboardParagraph, ClipboardTable, ClipboardText
from docwen_gui import clipboard_office_provider as provider
from docwen_gui.clipboard_capture import freeze_clipboard_mime
from docwen_gui.clipboard_rich_document import project_frozen_rich_document

pytestmark = pytest.mark.contract

_SAMPLES = Path(__file__).resolve().parents[4] / "tests/fixtures/files/clipboard-office"
_FREE = 0xFFFFFFFF
_END = 0xFFFFFFFE
_FAT = 0xFFFFFFFD
_DIFAT = 0xFFFFFFFC
_IMAGE_HASHES = {
    "66d68226db981e8856cd373abedf046917cd5d801498754efdf1b4ecf70d6633",
    "020d53090eb9e490942c26afbfee75b8b568059ad5955aad1f7616dfaf28e8c3",
}


def _sample(name: str) -> bytes:
    data = (_SAMPLES / name).read_bytes()
    facts = json.loads((_SAMPLES / "provenance.json").read_text(encoding="utf-8"))["files"][name]
    assert len(data) == facts["bytes"]
    assert hashlib.sha256(data).hexdigest() == facts["sha256"]
    return data


def _directory(package_size: int, package_start: int, *, root_start: int = _END, root_size: int = 0) -> bytes:
    result = bytearray(512)
    for index, name, kind, start, size in [
        (0, "Root Entry", 5, root_start, root_size),
        (1, "Package", 2, package_start, package_size),
    ]:
        offset = index * 128
        label = (name + "\0").encode("utf-16le")
        result[offset : offset + len(label)] = label
        struct.pack_into("<HBBIII", result, offset + 64, len(label), kind, 1, _FREE, _FREE, 1 if index == 0 else _FREE)
        struct.pack_into("<IQ", result, offset + 116, start, size)
    return bytes(result)


def _ole(package: bytes, *, fat_count: int = 1, major: int = 3) -> bytes:
    """Build a complete control, including actual extended DIFAT chains."""
    assert len(package) >= 4096
    sector_size = 512 if major == 3 else 4096
    capacity = sector_size // 4
    difat_count = max(0, (fat_count - 109 + capacity - 2) // (capacity - 1))
    package_count = (len(package) + sector_size - 1) // sector_size
    package_start = 1 + fat_count + difat_count
    header = bytearray(sector_size)
    header[:8] = bytes.fromhex("d0cf11e0a1b11ae1")
    struct.pack_into("<HHHHH", header, 24, 0x3E, major, 0xFFFE, 9 if major == 3 else 12, 6)
    struct.pack_into(
        "<IIIIIIIII",
        header,
        40,
        0 if major == 3 else 1,
        fat_count,
        0,
        0,
        4096,
        _END,
        0,
        1 + fat_count if difat_count else _END,
        difat_count,
    )
    fat_sectors = list(range(1, fat_count + 1))
    struct.pack_into("<109I", header, 76, *(fat_sectors[:109] + [_FREE] * max(0, 109 - fat_count)))
    fat = [_FREE] * (fat_count * capacity)
    fat[0] = _END
    for sector in fat_sectors:
        fat[sector] = _FAT
    for sector in range(1 + fat_count, package_start):
        fat[sector] = _DIFAT
    for index in range(package_count):
        sector = package_start + index
        fat[sector] = sector + 1 if index < package_count - 1 else _END
    parts = [bytes(header), _directory(len(package), package_start).ljust(sector_size, b"\0")]
    for index in range(fat_count):
        parts.append(struct.pack(f"<{capacity}I", *fat[index * capacity : (index + 1) * capacity]))
    remaining = fat_sectors[109:]
    for index in range(difat_count):
        values = remaining[index * (capacity - 1) : (index + 1) * (capacity - 1)]
        next_sector = 1 + fat_count + index + 1 if index < difat_count - 1 else _END
        parts.append(struct.pack(f"<{capacity}I", *(values + [_FREE] * (capacity - 1 - len(values)) + [next_sector])))
    return b"".join(parts) + package


def _patch_u32(payload: bytes, offset: int, value: int) -> bytes:
    modified = bytearray(payload)
    struct.pack_into("<I", modified, offset, value)
    return bytes(modified)


def _package() -> bytes:
    return _sample("word-derived.ole")[1536:]


def _named_stream(name: str, stream: bytes) -> bytes:
    payload = bytearray(_ole(stream))
    label = (name + "\0").encode("utf-16le")
    payload[640:704] = label.ljust(64, b"\0")
    struct.pack_into("<H", payload, 704, len(label))
    return bytes(payload)


def _biff_workbook() -> bytes:
    bof = struct.pack("<HHHH", 0x0809, 4, 0x0600, 0x0005)
    record = struct.pack("<HH", 0x0018, 4096) + b"\0" * 4096
    return bof + record + struct.pack("<HH", 0x000A, 0)


def _xlsx_package(*, root: str = "workbook", include_word: bool = False) -> bytes:
    target = io.BytesIO()
    with zipfile.ZipFile(target, "w", compression=zipfile.ZIP_STORED) as archive:
        archive.writestr(
            "xl/workbook.xml",
            f'<{root} xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>',
        )
        archive.writestr("padding.bin", b"\0" * 4096)
        if include_word:
            archive.writestr("word/document.xml", b"<document/>")
    return target.getvalue()


@pytest.mark.parametrize("name", ["Workbook", "Book"])
def test_generic_valid_biff_source_does_not_displace_html_table(name: str, qapp) -> None:
    mime = QMimeData()
    mime.setText("00123\t2\r\n")
    mime.setHtml("<table><tr><td>00123</td><td>2</td></tr></table>")
    mime.setData(provider.WORD_EMBED_SOURCE_MIME, _named_stream(name, _biff_workbook()))
    decision = project_frozen_rich_document(freeze_clipboard_mime(mime))
    assert decision.projection is not None and not decision.plain_fallback
    assert decision.projection.provider == "" and not decision.projection.resources
    table, trailing = decision.projection.document.blocks
    assert isinstance(table, ClipboardTable)
    assert (table.row_count, table.column_count) == (1, 2)
    assert table.cells[0].blocks == (ClipboardParagraph((ClipboardText("00123"),)),)
    assert table.cells[1].blocks == (ClipboardParagraph((ClipboardText("2"),)),)
    assert trailing == ClipboardParagraph((ClipboardText("\r\n"),))


@pytest.mark.parametrize("name", ["Package", "package"])
def test_generic_valid_ooxml_sheet_does_not_displace_html_table(name: str, qapp) -> None:
    mime = QMimeData()
    mime.setText("00123\t2\r\n")
    mime.setHtml("<table><tr><td>00123</td><td>2</td></tr></table>")
    mime.setData(provider.WORD_EMBED_SOURCE_MIME, _named_stream(name, _xlsx_package()))
    decision = project_frozen_rich_document(freeze_clipboard_mime(mime))
    assert decision.projection is not None and not decision.plain_fallback
    assert not decision.projection.resources
    table, trailing = decision.projection.document.blocks
    assert isinstance(table, ClipboardTable)
    assert (table.row_count, table.column_count) == (1, 2)
    assert trailing == ClipboardParagraph((ClipboardText("\r\n"),))


@pytest.mark.parametrize("defect", ["missing_eof", "truncated", "bad_version", "wrong_root", "mixed"])
def test_generic_malformed_or_ambiguous_sheet_source_preserves_full_plain(defect: str, qapp) -> None:
    name, stream = "Workbook", _biff_workbook()
    if defect == "missing_eof":
        stream = stream[:-4]
    elif defect == "truncated":
        stream = stream[:-5]
    elif defect == "bad_version":
        value = bytearray(stream)
        struct.pack_into("<H", value, 4, 0x0000)
        stream = bytes(value)
    else:
        name = "package"
        stream = _xlsx_package(root="wrong" if defect == "wrong_root" else "workbook", include_word=defect == "mixed")
    mime = QMimeData()
    plain = " 00123\t2\r\nUNCHANGED  \t\n"
    mime.setText(plain)
    mime.setHtml("<table><tr><td>00123</td><td>2</td></tr></table>")
    mime.setData(provider.WORD_EMBED_SOURCE_MIME, _named_stream(name, stream))
    frozen = freeze_clipboard_mime(mime)
    decision = project_frozen_rich_document(frozen)
    assert decision.projection is None and decision.plain_fallback
    assert frozen.plain_text == plain


def test_generic_word_source_still_binds_original_images() -> None:
    projection = provider.parse_embed_source_images(_sample("word-derived.ole"))
    assert projection is not None and projection.provider == "word"
    assert len(projection.occurrences) == 3 and len(projection.resources) == 2


@pytest.mark.parametrize("source", ["word", "wps_writer"])
def test_native_derived_provider_retains_order_extent_and_original_png_bytes(source: str) -> None:
    if source == "word":
        projection = provider.parse_word_embed_source(_sample("word-derived.ole"))
    else:
        projection = provider.parse_wps_writer(_sample("wps-writer-derived.zip"), _sample("wps-images.bin"))
    assert projection.provider == source
    assert len(projection.resources) == 2
    assert len(projection.occurrences) == 3
    first, nested, last = projection.occurrences
    assert first.identifier == last.identifier != nested.identifier
    assert [item.container_path.count("/table:") for item in projection.occurrences] == [0, 2, 0]
    assert [(item.extent_cx_emu, item.extent_cy_emu) for item in projection.occurrences] == [
        (914400, 609600),
        (457200, 304800),
        (914400, 609600),
    ]
    assert {(item.pixel_width, item.pixel_height) for item in projection.resources} == {(96, 64)}
    assert {item.sha256 for item in projection.resources} == _IMAGE_HASHES
    assert {hashlib.sha256(item[3]).hexdigest() for item in projection.resource_bytes} == _IMAGE_HASHES


def test_word_logical_tail_needs_no_sector_padding() -> None:
    payload = _sample("word-derived.ole")
    assert len(payload) % 512 != 0
    package = provider._read_cfb_package(payload)
    assert package == _package()
    with zipfile.ZipFile(io.BytesIO(package)) as archive:
        assert "word/document.xml" in archive.namelist()
        assert all(not name.startswith("docProps/") for name in archive.namelist())


@pytest.mark.parametrize("major,fat_count", [(3, 1), (3, 2), (3, 110), (3, 237), (4, 1)])
def test_cfb_complete_controls_including_extended_difat(major: int, fat_count: int) -> None:
    package = _package()
    assert provider._read_cfb_package(_ole(package, fat_count=fat_count, major=major)) == package


@pytest.mark.parametrize("offset,value", [(44, _FREE), (64, _FREE), (72, _FREE), (40, 1), (68, 0)])
def test_cfb_inconsistent_counts_rejected_before_sector_traversal(offset: int, value: int, monkeypatch) -> None:
    payload = _patch_u32(_sample("word-derived.ole"), offset, value)
    original = provider.struct.unpack

    def guard(fmt, buffer):
        if fmt.endswith("I"):
            pytest.fail("Malformed header reached sector traversal")
        return original(fmt, buffer)

    monkeypatch.setattr(provider.struct, "unpack", guard)
    with pytest.raises(provider.ClipboardOfficeProviderError) as failure:
        provider._read_cfb_package(payload)
    assert failure.value.code == "clipboard.provider_ole_invalid"


def test_cfb_cyclic_difat_fails_before_rereading_sector(monkeypatch) -> None:
    payload = _ole(_package(), fat_count=237)
    first_difat = 238
    modified = _patch_u32(payload, (first_difat + 1) * 512 + 508, first_difat)
    original = provider.struct.unpack
    reads = 0

    def guard(fmt, buffer):
        nonlocal reads
        if fmt == "<128I":
            reads += 1
            if reads > 1:
                pytest.fail("Cyclic DIFAT was reread instead of rejected")
        return original(fmt, buffer)

    monkeypatch.setattr(provider.struct, "unpack", guard)
    with pytest.raises(provider.ClipboardOfficeProviderError):
        provider._read_cfb_package(modified)
    assert reads == 1


@pytest.mark.parametrize(
    "defect", ["duplicate_fat", "fat_difat_overlap", "directory_fat_overlap", "unterminated_difat"]
)
def test_cfb_metadata_overlap_or_duplicate_rejected(defect: str) -> None:
    payload = _ole(_package(), fat_count=110 if "difat" in defect else 2)
    if defect == "duplicate_fat":
        payload = _patch_u32(payload, 80, 1)
    elif defect == "fat_difat_overlap":
        payload = _patch_u32(payload, 76, 111)
    elif defect == "directory_fat_overlap":
        payload = _patch_u32(payload, 48, 1)
    else:
        payload = _patch_u32(payload, 112 * 512 + 508, 111)
    with pytest.raises(provider.ClipboardOfficeProviderError):
        provider._read_cfb_package(payload)


@pytest.mark.parametrize(
    "defect",
    ["missing_byte", "directory_tail", "fat_tail", "cycle", "free_termination", "extra_sector", "duplicate_package"],
)
def test_cfb_truncation_and_bad_streams_fail_closed(defect: str) -> None:
    payload = _sample("word-derived.ole")
    last_sector = (len(payload) - 512 + 511) // 512 - 1
    if defect == "missing_byte":
        payload = payload[:-1]
    elif defect == "directory_tail":
        payload = _patch_u32(payload, 48, last_sector)
    elif defect == "fat_tail":
        payload = _patch_u32(payload, 76, last_sector)
    elif defect == "cycle":
        payload = _patch_u32(payload, 512 + last_sector * 4, 2)
    elif defect == "free_termination":
        payload = _patch_u32(payload, 512 + last_sector * 4, _FREE)
    elif defect == "extra_sector":
        payload = payload.ljust((last_sector + 2) * 512, b"\0") + bytes(512)
        payload = _patch_u32(payload, 512 + last_sector * 4, last_sector + 1)
        payload = _patch_u32(payload, 512 + (last_sector + 1) * 4, _END)
    else:
        modified = bytearray(payload)
        modified[1280:1408] = modified[1152:1280]
        payload = bytes(modified)
    with pytest.raises(provider.ClipboardOfficeProviderError) as failure:
        provider._read_cfb_package(payload)
    assert failure.value.code == "clipboard.provider_ole_invalid"


def test_cfb_small_package_uses_bounded_ministream() -> None:
    package = b"PK\x03\x04" + bytes(331)
    root_size = (len(package) + 63) // 64 * 64
    header = bytearray(_ole(_package())[:512])
    struct.pack_into("<I", header, 48, 1)
    struct.pack_into("<I", header, 76, 0)
    struct.pack_into("<II", header, 60, 2, 1)
    fat = [_FAT, _END, _END, _END] + [_FREE] * 124
    mini_count = root_size // 64
    mini_fat = list(range(1, mini_count)) + [_END] + [_FREE] * (128 - mini_count)
    payload = (
        bytes(header)
        + struct.pack("<128I", *fat)
        + _directory(len(package), 0, root_start=3, root_size=root_size)
        + struct.pack("<128I", *mini_fat)
        + package.ljust(root_size, b"\0")
    )
    assert provider._read_cfb_package(payload) == package
    with pytest.raises(provider.ClipboardOfficeProviderError):
        provider._read_cfb_package(payload[: -(root_size - len(package) + 1)])

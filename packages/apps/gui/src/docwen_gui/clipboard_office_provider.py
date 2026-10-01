"""Bounded native clipboard adapters for controlled Word and WPS writer MIME.

The adapters consume frozen bytes only. They never dereference HTML paths,
SourceURL/base values, external OOXML relationships, or execute payload data.
"""

from __future__ import annotations

import base64
import io
import posixpath
import struct
import zipfile
from collections.abc import Callable
from dataclasses import dataclass
from pathlib import PurePosixPath
from typing import NoReturn

from lxml import etree

from docwen_core.models.clipboard_document import ClipboardResource
from docwen_gui.clipboard_image_bytes import FrozenPng, inspect_png_bytes

WORD_EMBED_SOURCE_MIME = 'application/x-qt-windows-mime;value="Embed Source"'
WPS_DOCUMENT_MIME = 'application/x-qt-windows-mime;value="Kingsoft WPS 9.0 Format"'
WPS_IMAGE_DATA_MIME = 'application/x-qt-windows-mime;value="Kingsoft Image Data"'

_MAX_PROVIDER_BYTES = 32 * 1024 * 1024
_MAX_ZIP_ENTRIES = 512
_MAX_ZIP_UNCOMPRESSED = 64 * 1024 * 1024
_MAX_CFB_SECTORS = 131072

_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_IMAGE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
_NS = {"w": _W_NS, "a": _A_NS, "r": _R_NS, "wp": _WP_NS}

_CFB_SIGNATURE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
_CFB_FREE = 0xFFFFFFFF
_CFB_END = 0xFFFFFFFE
_CFB_FAT = 0xFFFFFFFD
_CFB_DIFAT = 0xFFFFFFFC


class ClipboardOfficeProviderError(ValueError):
    """Stable provider parsing failure."""

    def __init__(self, code: str, message: str) -> None:
        self.code = code
        super().__init__(message)


@dataclass(frozen=True, slots=True)
class ProviderImageOccurrence:
    identifier: str
    container_path: str
    previous_text: str
    next_text: str
    extent_cx_emu: int
    extent_cy_emu: int


@dataclass(frozen=True, slots=True)
class ProviderImageResource:
    identifier: str
    resource_id: str
    logical_path: str
    png: FrozenPng


@dataclass(frozen=True, slots=True)
class ProviderImageProjection:
    provider: str
    occurrences: tuple[ProviderImageOccurrence, ...]
    resources: tuple[ClipboardResource, ...]
    resource_bytes: tuple[tuple[str, str, str, bytes], ...]
    identifier_to_resource_id: tuple[tuple[str, str], ...]


def _fail(code: str, message: str) -> NoReturn:
    raise ClipboardOfficeProviderError(code, message)


def _bounded_chain(start: int, table: list[int], *, max_items: int) -> list[int]:
    if start == _CFB_END:
        return []
    chain: list[int] = []
    seen: set[int] = set()
    current = start
    while current != _CFB_END:
        if current >= 0xFFFFFFFA or current < 0 or current >= len(table):
            _fail("clipboard.provider_ole_invalid", "Word clipboard package has an invalid sector chain.")
        if current in seen or len(chain) >= max_items:
            _fail("clipboard.provider_ole_invalid", "Word clipboard package has a cyclic sector chain.")
        seen.add(current)
        chain.append(current)
        current = table[current]
    return chain


def _read_cfb_package(payload: bytes) -> bytes:
    """Read the unique Package stream from one bounded CFB payload."""

    if not isinstance(payload, bytes) or len(payload) < 512 or len(payload) > _MAX_PROVIDER_BYTES:
        _fail("clipboard.provider_ole_invalid", "Word clipboard package is invalid.")
    if payload[:8] != _CFB_SIGNATURE:
        _fail("clipboard.provider_ole_invalid", "Word clipboard package is not a supported OLE container.")

    _minor, major, byte_order, sector_shift, mini_shift = struct.unpack_from("<HHHHH", payload, 24)
    if byte_order != 0xFFFE or (major, sector_shift) not in {(3, 9), (4, 12)} or mini_shift != 6:
        _fail("clipboard.provider_ole_invalid", "Word clipboard package uses an unsupported OLE layout.")
    sector_size = 1 << sector_shift
    if len(payload) < sector_size:
        _fail("clipboard.provider_ole_invalid", "Word clipboard package is truncated.")
    # Native Word clipboard data can omit padding after the last logical
    # Package byte. Metadata sectors must still be physically complete.
    sector_count = (len(payload) - sector_size + sector_size - 1) // sector_size
    if sector_count < 1 or sector_count > _MAX_CFB_SECTORS:
        _fail("clipboard.provider_ole_invalid", "Word clipboard package exceeds the supported sector budget.")

    (
        num_directory_sectors,
        num_fat_sectors,
        first_directory_sector,
        _transaction_signature,
        mini_cutoff,
        first_mini_fat_sector,
        num_mini_fat_sectors,
        first_difat_sector,
        num_difat_sectors,
    ) = struct.unpack_from("<IIIIIIIII", payload, 40)
    if mini_cutoff != 4096:
        _fail("clipboard.provider_ole_invalid", "Word clipboard package has an unsupported mini-stream cutoff.")

    entries_per_sector = sector_size // 4
    expected_difat = max(0, (num_fat_sectors - 109 + entries_per_sector - 2) // (entries_per_sector - 1))
    if (
        not 1 <= num_fat_sectors <= sector_count
        or num_mini_fat_sectors > sector_count
        or num_directory_sectors > sector_count
        or (major == 3 and num_directory_sectors != 0)
        or num_difat_sectors > sector_count
        or num_difat_sectors != expected_difat
        or (num_difat_sectors == 0 and first_difat_sector != _CFB_END)
    ):
        _fail("clipboard.provider_ole_invalid", "Word clipboard package has inconsistent sector counts.")

    def sector_bytes(sector: int, *, complete: bool = False) -> bytes:
        if sector < 0 or sector >= sector_count:
            _fail("clipboard.provider_ole_invalid", "Word clipboard package references an invalid sector.")
        start = sector_size + sector * sector_size
        data = payload[start : start + sector_size]
        if complete and len(data) != sector_size:
            _fail("clipboard.provider_ole_invalid", "Word clipboard package has truncated metadata.")
        return data

    difat = list(struct.unpack_from("<109I", payload, 76))
    next_difat = first_difat_sector
    difat_sectors: set[int] = set()
    for _ in range(num_difat_sectors):
        if next_difat >= sector_count or next_difat in difat_sectors:
            _fail("clipboard.provider_ole_invalid", "Word clipboard package has an invalid DIFAT chain.")
        difat_sectors.add(next_difat)
        values = struct.unpack(f"<{entries_per_sector}I", sector_bytes(next_difat, complete=True))
        difat.extend(values[:-1])
        next_difat = values[-1]
    if next_difat != _CFB_END:
        _fail("clipboard.provider_ole_invalid", "Word clipboard package has an unterminated DIFAT chain.")
    fat_sectors = [value for value in difat if value != _CFB_FREE]
    if (
        len(fat_sectors) != num_fat_sectors
        or len(set(fat_sectors)) != len(fat_sectors)
        or difat_sectors.intersection(fat_sectors)
        or any(value >= sector_count for value in fat_sectors)
    ):
        _fail("clipboard.provider_ole_invalid", "Word clipboard package references an invalid FAT sector.")

    fat: list[int] = []
    for sector in fat_sectors:
        values = struct.unpack(f"<{entries_per_sector}I", sector_bytes(sector, complete=True))
        # The input bounds the physical sector count, including its logical
        # tail. Surplus FAT capacity must not amplify the in-memory table.
        fat.extend(values[: max(0, sector_count - len(fat))])
    if len(fat) != sector_count or any(fat[sector] != _CFB_FAT for sector in fat_sectors):
        _fail("clipboard.provider_ole_invalid", "Word clipboard package has inconsistent FAT metadata.")
    if any(fat[sector] != _CFB_DIFAT for sector in difat_sectors):
        _fail("clipboard.provider_ole_invalid", "Word clipboard package has inconsistent DIFAT metadata.")

    claimed_sectors = set(fat_sectors) | difat_sectors

    def claim_chain(start: int, *, max_items: int = sector_count) -> list[int]:
        chain = _bounded_chain(start, fat, max_items=max_items)
        if claimed_sectors.intersection(chain):
            _fail("clipboard.provider_ole_invalid", "Word clipboard package has overlapping sector chains.")
        claimed_sectors.update(chain)
        return chain

    def read_stream(start: int, size: int) -> bytes:
        if not 1 <= size <= _MAX_PROVIDER_BYTES:
            _fail("clipboard.provider_ole_invalid", "Word clipboard stream exceeds the supported size.")
        expected_count = (size + sector_size - 1) // sector_size
        chain = claim_chain(start, max_items=expected_count)
        if len(chain) != expected_count:
            _fail("clipboard.provider_ole_invalid", "Word clipboard stream has an inconsistent sector count.")
        chunks = [sector_bytes(sector, complete=index < len(chain) - 1) for index, sector in enumerate(chain)]
        stream = b"".join(chunks)
        if len(stream) < size:
            _fail("clipboard.provider_ole_invalid", "Word clipboard Package stream is truncated.")
        return stream[:size]

    directory_chain = claim_chain(first_directory_sector)
    if not directory_chain or (major == 4 and len(directory_chain) != num_directory_sectors):
        _fail("clipboard.provider_ole_invalid", "Word clipboard package has an invalid directory chain.")
    directory = b"".join(sector_bytes(sector, complete=True) for sector in directory_chain)
    entries: list[tuple[str, int, int, int]] = []
    root_start = _CFB_END
    root_size = 0
    root_entries = 0
    for offset in range(0, len(directory) - 127, 128):
        entry = directory[offset : offset + 128]
        name_length = struct.unpack_from("<H", entry, 64)[0]
        entry_type = entry[66]
        if entry_type == 0:
            continue
        if (
            name_length < 2
            or name_length > 64
            or name_length % 2
            or entry[name_length - 2 : name_length] != b"\0\0"
            or entry_type not in {1, 2, 5}
        ):
            _fail("clipboard.provider_ole_invalid", "Word clipboard directory contains an invalid name.")
        try:
            name = entry[: name_length - 2].decode("utf-16le", errors="strict")
        except UnicodeDecodeError:
            _fail("clipboard.provider_ole_invalid", "Word clipboard directory contains an invalid name.")
        start_sector = struct.unpack_from("<I", entry, 116)[0]
        size = struct.unpack_from("<Q", entry, 120)[0]
        if major == 3 and size > 0xFFFFFFFF:
            _fail("clipboard.provider_ole_invalid", "Word clipboard stream has an invalid size.")
        if entry_type == 5:
            root_entries += 1
            root_start, root_size = start_sector, size
        entries.append((name, entry_type, start_sector, size))

    packages = [entry for entry in entries if entry[0] == "Package" and entry[1] == 2]
    if root_entries != 1 or len(packages) != 1:
        _fail("clipboard.provider_ole_invalid", "Word clipboard package must contain exactly one Package stream.")
    _name, _entry_type, package_start, package_size = packages[0]
    if package_size < 1 or package_size > _MAX_PROVIDER_BYTES:
        _fail("clipboard.provider_ole_invalid", "Word clipboard Package stream exceeds the supported size.")

    if package_size >= mini_cutoff:
        return read_stream(package_start, package_size)

    if root_size < 1 or root_start in {_CFB_END, _CFB_FREE}:
        _fail("clipboard.provider_ole_invalid", "Word clipboard mini stream is missing.")
    root_stream = read_stream(root_start, root_size)
    mini_fat_chain = claim_chain(first_mini_fat_sector, max_items=num_mini_fat_sectors)
    if not mini_fat_chain or len(mini_fat_chain) != num_mini_fat_sectors:
        _fail("clipboard.provider_ole_invalid", "Word clipboard mini FAT is inconsistent.")
    mini_fat_bytes = b"".join(sector_bytes(sector, complete=True) for sector in mini_fat_chain)
    mini_fat = list(struct.unpack_from(f"<{len(mini_fat_bytes) // 4}I", mini_fat_bytes, 0))
    mini_count = (root_size + 63) // 64
    mini_fat = mini_fat[:mini_count]
    expected_count = (package_size + 63) // 64
    mini_chain = _bounded_chain(package_start, mini_fat, max_items=expected_count)
    if len(mini_chain) != expected_count:
        _fail("clipboard.provider_ole_invalid", "Word clipboard mini stream has an inconsistent sector count.")
    stream = b"".join(root_stream[index * 64 : (index + 1) * 64] for index in mini_chain)
    if len(stream) < package_size:
        _fail("clipboard.provider_ole_invalid", "Word clipboard Package stream is truncated.")
    return stream[:package_size]


def _safe_zip_parts(payload: bytes) -> dict[str, bytes]:
    if not payload.startswith(b"PK\x03\x04") or len(payload) > _MAX_PROVIDER_BYTES:
        _fail("clipboard.provider_zip_invalid", "Clipboard document package is not a supported OOXML ZIP.")
    try:
        archive = zipfile.ZipFile(io.BytesIO(payload))
    except (OSError, zipfile.BadZipFile):
        _fail("clipboard.provider_zip_invalid", "Clipboard document package is not a valid OOXML ZIP.")

    infos = archive.infolist()
    if not infos or len(infos) > _MAX_ZIP_ENTRIES:
        _fail("clipboard.provider_zip_invalid", "Clipboard OOXML package exceeds the entry budget.")
    names: set[str] = set()
    total = 0
    for info in infos:
        name = info.filename
        path = PurePosixPath(name)
        if (
            not name
            or "\\" in name
            or name.startswith("/")
            or ".." in path.parts
            or name in names
            or info.flag_bits & 0x1
            or info.compress_type not in {zipfile.ZIP_STORED, zipfile.ZIP_DEFLATED}
        ):
            _fail("clipboard.provider_zip_invalid", "Clipboard OOXML package contains an unsafe entry.")
        names.add(name)
        total += int(info.file_size)
        if total > _MAX_ZIP_UNCOMPRESSED:
            _fail("clipboard.provider_zip_invalid", "Clipboard OOXML package exceeds the uncompressed byte budget.")

    parts: dict[str, bytes] = {}
    try:
        for info in infos:
            if info.is_dir():
                continue
            parts[info.filename] = archive.read(info)
    except (OSError, RuntimeError, zipfile.BadZipFile):
        _fail("clipboard.provider_zip_invalid", "Clipboard OOXML package could not be read safely.")
    return parts


def _parse_xml(payload: bytes, *, code: str, message: str):
    try:
        parser = etree.XMLParser(resolve_entities=False, no_network=True, recover=False)
        root = etree.fromstring(payload, parser)
    except (etree.XMLSyntaxError, ValueError):
        _fail(code, message)
    if getattr(root.getroottree().docinfo, "doctype", ""):
        _fail(code, message)
    return root


def _paragraph_text(paragraph) -> str:
    return "".join(paragraph.xpath(".//w:t/text()", namespaces=_NS))


def _direct_paragraph_context(children: list[object], index: int) -> tuple[str, str]:
    previous = ""
    following = ""
    for candidate in reversed(children[:index]):
        if getattr(candidate, "tag", None) == f"{{{_W_NS}}}p":
            text = _paragraph_text(candidate)
            if text:
                previous = text
                break
    for candidate in children[index + 1 :]:
        if getattr(candidate, "tag", None) == f"{{{_W_NS}}}p":
            text = _paragraph_text(candidate)
            if text:
                following = text
                break
    return previous, following


def _extent_for_blip(blip) -> tuple[int, int]:
    drawings = blip.xpath("ancestor::w:drawing[1]", namespaces=_NS)
    if not drawings:
        _fail("clipboard.provider_image_invalid", "Clipboard document image has no DrawingML extent.")
    extents = drawings[0].xpath(".//wp:extent[1]", namespaces=_NS)
    if len(extents) != 1:
        _fail("clipboard.provider_image_invalid", "Clipboard document image has no unique DrawingML extent.")
    try:
        cx = int(extents[0].get("cx") or "0")
        cy = int(extents[0].get("cy") or "0")
    except ValueError:
        _fail("clipboard.provider_image_invalid", "Clipboard document image extent is invalid.")
    if cx < 1 or cy < 1 or cx > 2_147_483_647 or cy > 2_147_483_647:
        _fail("clipboard.provider_image_invalid", "Clipboard document image extent is out of range.")
    return cx, cy


def _collect_occurrences(
    document_xml: bytes, resolve_identifier: Callable[[str], str]
) -> tuple[ProviderImageOccurrence, ...]:
    root = _parse_xml(
        document_xml,
        code="clipboard.provider_document_invalid",
        message="Clipboard document XML is invalid.",
    )
    body = root.find(f"{{{_W_NS}}}body")
    if body is None:
        _fail("clipboard.provider_document_invalid", "Clipboard document XML has no body.")
    output: list[ProviderImageOccurrence] = []

    def walk_container(container, path: str) -> None:
        children = [child for child in container if isinstance(getattr(child, "tag", None), str)]
        table_index = 0
        for index, child in enumerate(children):
            if child.tag == f"{{{_W_NS}}}p":
                previous, following = _direct_paragraph_context(children, index)
                blips = child.xpath(".//a:blip", namespaces=_NS)
                for blip in blips:
                    raw_identifier = blip.get(f"{{{_R_NS}}}embed")
                    if not raw_identifier:
                        _fail(
                            "clipboard.provider_image_invalid", "Clipboard document image has no embedded identifier."
                        )
                    cx, cy = _extent_for_blip(blip)
                    output.append(
                        ProviderImageOccurrence(
                            identifier=resolve_identifier(raw_identifier),
                            container_path=path,
                            previous_text=previous,
                            next_text=following,
                            extent_cx_emu=cx,
                            extent_cy_emu=cy,
                        )
                    )
            elif child.tag == f"{{{_W_NS}}}tbl":
                walk_table(child, f"{path}/table:{table_index}")
                table_index += 1

    def walk_table(table, path: str) -> None:
        rows = [child for child in table if child.tag == f"{{{_W_NS}}}tr"]
        for row_index, row in enumerate(rows):
            column = 0
            for cell in [child for child in row if child.tag == f"{{{_W_NS}}}tc"]:
                tc_pr = cell.find(f"{{{_W_NS}}}tcPr")
                span = 1
                if tc_pr is not None:
                    grid_span = tc_pr.find(f"{{{_W_NS}}}gridSpan")
                    if grid_span is not None:
                        try:
                            span = max(1, int(grid_span.get(f"{{{_W_NS}}}val") or "1"))
                        except ValueError:
                            _fail("clipboard.provider_document_invalid", "Clipboard table span is invalid.")
                walk_container(cell, f"{path}/cell:{row_index}:{column}")
                column += span

    walk_container(body, "body")
    return tuple(output)


def _relationship_targets(parts: dict[str, bytes]) -> dict[str, str]:
    rel_payload = parts.get("word/_rels/document.xml.rels")
    if rel_payload is None:
        _fail("clipboard.provider_document_invalid", "Clipboard Word package is missing document relationships.")
    root = _parse_xml(
        rel_payload,
        code="clipboard.provider_document_invalid",
        message="Clipboard Word relationships are invalid.",
    )
    relationships: dict[str, str] = {}
    for relation in root:
        if relation.tag != f"{{{_REL_NS}}}Relationship":
            continue
        rel_id = relation.get("Id") or ""
        rel_type = relation.get("Type") or ""
        if rel_type != _IMAGE_REL:
            continue
        if relation.get("TargetMode") == "External":
            _fail("clipboard.provider_external_relationship", "External clipboard image relationships are not allowed.")
        target = relation.get("Target") or ""
        normalized = posixpath.normpath(posixpath.join("word", target))
        if not rel_id or normalized.startswith(("../", "/")) or "\\" in normalized or normalized not in parts:
            _fail("clipboard.provider_document_invalid", "Clipboard Word image relationship target is invalid.")
        if rel_id in relationships:
            _fail("clipboard.provider_document_invalid", "Clipboard Word image relationship is duplicated.")
        relationships[rel_id] = normalized
    return relationships


def _resource_projection(
    provider: str,
    occurrences: tuple[ProviderImageOccurrence, ...],
    identifier_payloads: dict[str, bytes],
) -> ProviderImageProjection:
    identifier_to_resource: dict[str, str] = {}
    resources_by_sha: dict[str, tuple[ClipboardResource, bytes]] = {}
    for identifier, payload in identifier_payloads.items():
        png = inspect_png_bytes(payload)
        resource_id = f"img-{png.payload_sha256[:24]}"
        logical_path = f"resources/{png.payload_sha256[:24]}.png"
        identifier_to_resource[identifier] = resource_id
        resources_by_sha.setdefault(
            png.payload_sha256,
            (
                ClipboardResource(
                    resource_id=resource_id,
                    logical_path=logical_path,
                    media_type="image/png",
                    size_bytes=len(payload),
                    sha256=png.payload_sha256,
                    pixel_width=png.width,
                    pixel_height=png.height,
                    rgba_sha256=png.rgba_sha256,
                ),
                payload,
            ),
        )

    referenced = {occurrence.identifier for occurrence in occurrences}
    if referenced != set(identifier_payloads):
        _fail(
            "clipboard.provider_binding_invalid",
            "Clipboard document image identifiers do not exactly match provided image resources.",
        )

    ordered = sorted(resources_by_sha.values(), key=lambda item: item[0].resource_id)
    return ProviderImageProjection(
        provider=provider,
        occurrences=occurrences,
        resources=tuple(item[0] for item in ordered),
        resource_bytes=tuple(
            (item[0].resource_id, item[0].logical_path, item[0].media_type, item[1]) for item in ordered
        ),
        identifier_to_resource_id=tuple(sorted(identifier_to_resource.items())),
    )


def parse_word_embed_source(payload: bytes) -> ProviderImageProjection:
    package = _read_cfb_package(payload)
    parts = _safe_zip_parts(package)
    document_xml = parts.get("word/document.xml")
    if document_xml is None:
        _fail("clipboard.provider_document_invalid", "Clipboard Word package is missing document XML.")
    relationships = _relationship_targets(parts)

    def resolve(identifier: str) -> str:
        if identifier not in relationships:
            _fail("clipboard.provider_binding_invalid", "Clipboard Word image relationship is unresolved.")
        return identifier

    occurrences = _collect_occurrences(document_xml, resolve)
    identifiers = {occurrence.identifier for occurrence in occurrences}
    payloads = {identifier: parts[relationships[identifier]] for identifier in identifiers}
    return _resource_projection("word", occurrences, payloads)


def _parse_wps_image_data(payload: bytes) -> dict[str, bytes]:
    if not isinstance(payload, bytes) or not payload or len(payload) > _MAX_PROVIDER_BYTES:
        _fail("clipboard.provider_wps_image_data_invalid", "WPS clipboard image data is invalid.")
    records: dict[str, bytes] = {}
    offset = 0
    while offset < len(payload):
        if len(payload) - offset < 4:
            _fail("clipboard.provider_wps_image_data_invalid", "WPS clipboard image data has trailing bytes.")
        record_length = struct.unpack_from("<I", payload, offset)[0]
        offset += 4
        if record_length <= 16 or record_length > len(payload) - offset:
            _fail("clipboard.provider_wps_image_data_invalid", "WPS clipboard image record length is invalid.")
        record = payload[offset : offset + record_length]
        offset += record_length
        identifier_raw = record[:16]
        image_payload = record[16:]
        identifier = base64.b64encode(identifier_raw).decode("ascii")
        if identifier in records:
            _fail("clipboard.provider_wps_image_data_invalid", "WPS clipboard image identifier is duplicated.")
        inspect_png_bytes(image_payload)
        records[identifier] = image_payload
    if offset != len(payload):
        _fail("clipboard.provider_wps_image_data_invalid", "WPS clipboard image data was not fully consumed.")
    return records


def parse_wps_writer(document_payload: bytes, image_data_payload: bytes) -> ProviderImageProjection:
    parts = _safe_zip_parts(document_payload)
    document_xml = parts.get("word/document.xml")
    if document_xml is None:
        _fail("clipboard.provider_document_invalid", "WPS clipboard package is missing document XML.")
    image_payloads = _parse_wps_image_data(image_data_payload)

    def resolve(identifier: str) -> str:
        try:
            raw = base64.b64decode(identifier, validate=True)
        except Exception:
            _fail("clipboard.provider_binding_invalid", "WPS clipboard image identifier is invalid.")
        if len(raw) != 16 or identifier not in image_payloads:
            _fail("clipboard.provider_binding_invalid", "WPS clipboard image identifier is unresolved.")
        return identifier

    occurrences = _collect_occurrences(document_xml, resolve)
    return _resource_projection("wps_writer", occurrences, image_payloads)


__all__ = [
    "WORD_EMBED_SOURCE_MIME",
    "WPS_DOCUMENT_MIME",
    "WPS_IMAGE_DATA_MIME",
    "ClipboardOfficeProviderError",
    "ProviderImageOccurrence",
    "ProviderImageProjection",
    "parse_word_embed_source",
    "parse_wps_writer",
]

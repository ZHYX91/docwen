"""Content-based file format detection — magic bytes, ZIP inspection,
OLE compound document parsing, text/binary classification, and text-format
sniffing.

This module serves as the single shared location for file detection
utilities previously scattered across the old ``file_type_sniffing.py``
and ``file_type_utils.py`` modules.

It MUST NOT import from any plugin, runtime, or application package.
"""

from __future__ import annotations

import codecs
import csv
import io
import os
import posixpath
import re
import zipfile
from collections.abc import Mapping
from types import MappingProxyType
from xml.etree import ElementTree

from docwen_core.models.file_inspection import (
    ContentDetection,
    DetectionConfidence,
    DetectionMethod,
    StructureStatus,
)

# ── Magic-byte signatures ─────────────────────────────────────────────

# Tuples of (signature, format_name).  Signatures that need further
# disambiguation (ZIP, OLE) return None for the format.
_SIGNATURES: list[tuple[bytes, str | None]] = [
    (b"\xff\xd8\xff", "jpeg"),
    (b"\x89PNG\r\n\x1a\n", "png"),
    (b"GIF87a", "gif"),
    (b"GIF89a", "gif"),
    (b"BM", "bmp"),
    (b"II*\x00", "tiff"),  # little-endian TIFF
    (b"MM\x00*", "tiff"),  # big-endian TIFF
    (b"%PDF", "pdf"),
    (b"{\\rtf", "rtf"),
    (b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1", None),  # OLE2/CFB → further inspection
    (b"\xd0\xcf\x11\xe0", None),  # OLE2 (short) → further inspection
    (b"PK\x03\x04", None),  # ZIP local-file header → further inspection
    (b"PK\x05\x06", None),  # Empty ZIP / end-of-central-directory record
    (b"PK\x07\x08", None),  # ZIP spanned/data-descriptor record
    (b"{\\rtf1", "rtf"),
]

# Signatures checked at specific offsets. HEIF permits several major brands;
# the public ``heic`` format name represents this whole decodable family.
_HEIF_FTYP_SIGNATURES: tuple[bytes, ...] = tuple(
    b"ftyp" + brand for brand in (b"heic", b"heix", b"hevc", b"hevx", b"mif1", b"msf1")
)
_OFFSET_SIGNATURES: list[tuple[int, bytes, str]] = [
    *((4, signature, "heic") for signature in _HEIF_FTYP_SIGNATURES),
    (8, b"WEBP", "webp"),  # WebP FourCC at offset 8 in RIFF container
]

# Signature read size
_SIG_READ_SIZE = 16
# Additional read size for WebP / ftyp detection
_OFFSET_READ_SIZE = 24
# Minimum read to cover both header and offset-based signatures
_SIG_MIN_READ = max(_SIG_READ_SIZE, _OFFSET_READ_SIZE)
_TEXT_READ_SIZE = 64 * 1024
_DELIMITED_SNIFF_SIZE = 16 * 1024
_DELIMITED_RECORD_LIMIT = 20

# ── ZIP-based container detection ─────────────────────────────────────

_OOXML_MAIN_PARTS: dict[str, tuple[str, str, str]] = {
    "docx": (
        "word/document.xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
        "document",
    ),
    "xlsx": (
        "xl/workbook.xml",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
        "workbook",
    ),
    "pptx": (
        "ppt/presentation.xml",
        "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
        "presentation",
    ),
}

# ZIP mimetype-based detection (ODT/ODS)
_ZIP_MIMETYPE_MAP: dict[str, str] = {
    "application/vnd.oasis.opendocument.text": "odt",
    "application/vnd.oasis.opendocument.spreadsheet": "ods",
    "application/vnd.oasis.opendocument.presentation": "odp",
}

# XML parts used for package admission are deliberately bounded.  This keeps
# a forged ZIP member from turning content sniffing into an unbounded
# decompression or XML-allocation path.  Admission fails closed above the
# centralized limit instead of attempting an unsafe partial parse.
_MAX_PACKAGE_XML_BYTES = 32 * 1024 * 1024
_MAX_PACKAGE_MIMETYPE_BYTES = 256
_XPS_FIXED_DOCUMENT_SEQUENCE_CONTENT_TYPE = "application/vnd.ms-package.xps-fixeddocumentsequence+xml"
_XPS_FIXED_DOCUMENT_SEQUENCE_NAMESPACE = "http://schemas.microsoft.com/xps/2005/06"

# ── OLE2 stream-based detection ───────────────────────────────────────

_OLE_STREAM_SIGNATURES: list[tuple[list[str], str]] = [
    (["WordDocument", "1Table", "0Table", "Data"], "doc"),
    (["Workbook", "Book", "BOOK"], "xls"),
    (["PowerPoint Document"], "ppt"),
]

# ── Supported filename declarations ──────────────────────────────────

# Public, immutable registry of known filename declarations and their concrete
# format names. Content inspection remains authoritative: this registry powers
# presentation filters and suffix/content comparison, never execution
# admission by itself.
SUPPORTED_EXTENSION_FORMATS: Mapping[str, str] = MappingProxyType(
    {
        ".jpg": "jpeg",
        ".jpeg": "jpeg",
        ".png": "png",
        ".gif": "gif",
        ".bmp": "bmp",
        ".tif": "tiff",
        ".tiff": "tiff",
        ".webp": "webp",
        ".heic": "heic",
        ".heif": "heic",
        ".pdf": "pdf",
        ".rtf": "rtf",
        ".doc": "doc",
        ".docx": "docx",
        ".odt": "odt",
        ".wps": "wps",
        ".xls": "xls",
        ".xlsx": "xlsx",
        ".ods": "ods",
        ".et": "et",
        ".csv": "csv",
        ".tsv": "tsv",
        ".ppt": "ppt",
        ".pptx": "pptx",
        ".ofd": "ofd",
        ".xps": "xps",
        ".md": "markdown",
        ".markdown": "markdown",
        ".txt": "txt",
        ".html": "html",
        ".htm": "html",
        ".mhtml": "mhtml",
        ".mht": "mhtml",
        ".epub": "epub",
        ".enex": "enex",
    }
)


# ── Public API ────────────────────────────────────────────────────────


def detect_content_format(file_path: str) -> ContentDetection:
    """Inspect content without ever trusting the filename extension.

    This function never falls back to the suffix. Ambiguous containers and
    unknown binary data remain explicit so callers can block them instead of
    accidentally routing a disguised or corrupt file.
    """

    if not os.path.isfile(file_path):
        raise FileNotFoundError(f"File not found: {file_path}")

    try:
        with open(file_path, "rb") as fh:
            head = fh.read(_SIG_MIN_READ)
    except OSError as exc:
        return ContentDetection(
            format="unknown",
            method=DetectionMethod.UNKNOWN,
            confidence=DetectionConfidence.UNVERIFIED,
            structure_status=StructureStatus.UNVERIFIED,
            detail_code="FILE_READ_ERROR",
            detail_message=str(exc),
        )

    if not head:
        return ContentDetection(
            format="unknown",
            method=DetectionMethod.UNKNOWN,
            confidence=DetectionConfidence.UNVERIFIED,
            structure_status=StructureStatus.INVALID,
            detail_code="FILE_EMPTY",
            detail_message="The input file is empty.",
        )

    for sig, fmt in _SIGNATURES:
        if not head.startswith(sig):
            continue
        if fmt is not None:
            return ContentDetection(
                format=fmt,
                method=DetectionMethod.SIGNATURE,
                confidence=DetectionConfidence.CERTAIN,
            )
        if sig in (b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1", b"\xd0\xcf\x11\xe0"):
            ole_format = _detect_ole_type(file_path)
            if ole_format:
                return ContentDetection(
                    format=ole_format,
                    method=DetectionMethod.CONTAINER,
                    confidence=DetectionConfidence.CERTAIN,
                    structure_status=StructureStatus.VALID,
                )
            return ContentDetection(
                format="ole",
                method=DetectionMethod.CONTAINER,
                confidence=DetectionConfidence.CERTAIN,
                structure_status=StructureStatus.UNVERIFIED,
                detail_code="FILE_CONTAINER_UNRECOGNIZED",
                detail_message="The OLE container does not contain a supported DOC, XLS, or PPT stream.",
            )
        if sig in {b"PK\x03\x04", b"PK\x05\x06", b"PK\x07\x08"}:
            return _inspect_zip_container(file_path)

    if len(head) >= _OFFSET_READ_SIZE:
        for offset, sig, fmt in _OFFSET_SIGNATURES:
            if head[offset : offset + len(sig)] == sig:
                return ContentDetection(
                    format=fmt,
                    method=DetectionMethod.SIGNATURE,
                    confidence=DetectionConfidence.CERTAIN,
                )

    if is_text_file(file_path):
        return ContentDetection(
            format=detect_text_format(file_path),
            method=DetectionMethod.TEXT_SNIFF,
            confidence=DetectionConfidence.PROBABLE,
        )

    return ContentDetection(
        format="unknown",
        method=DetectionMethod.UNKNOWN,
        confidence=DetectionConfidence.UNVERIFIED,
        structure_status=StructureStatus.UNVERIFIED,
        detail_code="FILE_CONTENT_UNRECOGNIZED",
        detail_message="The input content does not match a supported file signature or readable text format.",
    )


def has_known_signature(file_path: str) -> bool:
    """Return ``True`` if the file has a recognized binary magic-byte signature."""
    try:
        with open(file_path, "rb") as fh:
            head = fh.read(_SIG_MIN_READ)
    except OSError:
        return False

    if not head:
        return False

    for sig, _fmt in _SIGNATURES:
        if head.startswith(sig):
            return True

    if len(head) >= _OFFSET_READ_SIZE:
        for offset, sig, _fmt in _OFFSET_SIGNATURES:
            if head[offset : offset + len(sig)] == sig:
                return True

    return False


def is_text_file(file_path: str, sample_size: int = 8192) -> bool:
    """Return ``True`` if the file appears to be a text file (not binary).

    Uses a content-sample heuristic: reads *sample_size* bytes, checks the
    null-byte ratio (binary indicator), attempts UTF-8/GBK decoding, and
    measures printable-character ratio.
    """
    try:
        file_size = os.path.getsize(file_path)
    except OSError:
        return False

    if file_size == 0:
        return True  # Empty file is text

    read_size = min(sample_size, file_size)
    try:
        with open(file_path, "rb") as fh:
            data = fh.read(read_size)
    except OSError:
        return False

    if not data:
        return True

    decoded = _decode_text_bytes(data, final=read_size == file_size)
    if decoded is None:
        return False
    text, _encoding = decoded
    if not text:
        return True

    # Printable character ratio: > 85% printable → text
    printable = sum(1 for c in text if c.isprintable() or c.isspace())
    return printable / len(text) > 0.85


def detect_text_format(file_path: str) -> str:
    """Detect the text format of a file by content analysis.

    Reads the file content and inspects it for format-specific markers:
    Markdown (headings, YAML front matter, wiki/md links, code blocks),
    ENEX (``<en-export>``), MHTML (MIME headers), HTML (tags), CSV/TSV
    (delimiter analysis), falling back to ``"txt"``.

    Returns a format string: ``"markdown"``, ``"csv"``, ``"tsv"``,
    ``"html"``, ``"enex"``, ``"mhtml"``, or ``"txt"``.
    """
    # Try to read the file as text
    text = _read_text_file(file_path)
    if not text:
        # Text sniffing is content-only. Empty inputs are rejected by
        # ``detect_content_format`` before reaching this helper; callers using
        # it directly receive the neutral plain-text result rather than a
        # filename-derived guess.
        return "txt"

    sample = text[:4096]  # First 4K for format detection
    sample_lower = sample.lower()

    # ENEX — Evernote export format
    if "<en-export" in sample_lower or ('<?xml version="1.0" encoding="UTF-8"?>' in sample and "<en-export" in sample):
        return "enex"

    # MHTML — check for MIME headers
    if any(
        h in sample_lower for h in ("mime-version:", "multipart/related", "content-type:", "content-transfer-encoding:")
    ) and ("<html" in sample_lower or "boundary=" in sample_lower):
        return "mhtml"

    # HTML — check for HTML markers
    html_markers = ("<!doctype html", "<html", "<head", "<body", "<meta", "<div", "<p>")
    if any(m in sample_lower for m in html_markers):
        return "html"

    # Markdown — check for Markdown markers
    md_markers: list[str] = []
    for line in text.splitlines()[:50]:
        stripped = line.strip()
        if stripped:
            md_markers.append(stripped)

    md_score = 0
    for line_index, line in enumerate(md_markers):
        if line.startswith(("#", "```")):  # heading
            md_score += 2
        elif line.startswith("---") and len(line) <= 5:  # YAML front matter / HR
            md_score += 1
        elif line.startswith("|") and "|" in line[1:]:  # table
            md_score += 2
        elif (
            line_index > 0
            and "|" in md_markers[line_index - 1]
            and re.fullmatch(r"\|?\s*:?-{3,}:?\s*(?:\|\s*:?-{3,}:?\s*)+\|?", line)
        ):
            # Borderless Markdown tables are common and their delimiter row is
            # the only content-level discriminator from a pipe-delimited file.
            # Require an adjacent pipe-bearing header and the strict GFM
            # delimiter grammar before the generic CSV sniffer so a valid
            # embedded table is not rerouted as CSV, while a separator-shaped
            # value elsewhere in ordinary pipe data cannot promote the file.
            md_score += 2
        elif line.startswith(("- ", "* ", "+ ")):  # list
            md_score += 1
        elif "](" in line or "[[" in line:  # links / wiki links
            md_score += 2
        elif line.startswith("> "):  # blockquote
            md_score += 1

    if md_score >= 2:
        return "markdown"

    delimited_format = _detect_delimited_format(text)
    if delimited_format is not None:
        return delimited_format

    return "txt"


# ── Internal helpers ──────────────────────────────────────────────────


def _is_text_content(head: bytes) -> bool:
    """Check if the first 16 bytes appear to be text (printable ASCII only)."""
    return all(32 <= b <= 126 or b in (9, 10, 13) for b in head)


def _read_text_file(file_path: str, max_size: int = _TEXT_READ_SIZE) -> str | None:
    """Read a bounded text sample using the canonical strict decoder."""
    try:
        with open(file_path, "rb") as fh:
            data = fh.read(max_size + 4)
    except OSError:
        return None

    if not data:
        return ""

    truncated = len(data) > max_size
    sample = data[:max_size]
    decoded = _decode_text_bytes(sample, final=not truncated)
    return decoded[0] if decoded is not None else None


def _decode_text_bytes(data: bytes, *, final: bool) -> tuple[str, str] | None:
    """Strictly decode one bounded sample, recognizing Unicode BOMs first."""
    bom_encodings = (
        (codecs.BOM_UTF32_LE, "utf-32"),
        (codecs.BOM_UTF32_BE, "utf-32"),
        (codecs.BOM_UTF8, "utf-8-sig"),
        (codecs.BOM_UTF16_LE, "utf-16"),
        (codecs.BOM_UTF16_BE, "utf-16"),
    )
    for bom, encoding in bom_encodings:
        if not data.startswith(bom):
            continue
        try:
            decoder = codecs.getincrementaldecoder(encoding)(errors="strict")
            return decoder.decode(data, final=final), encoding
        except (UnicodeDecodeError, LookupError):
            return None

    # Without a Unicode BOM, a high NUL ratio remains a fail-closed binary
    # signal.  This prevents arbitrary binary data from being routed as text.
    if data and data.count(0) / len(data) > 0.1:
        return None

    for encoding in ("utf-8", "gbk"):
        try:
            decoder = codecs.getincrementaldecoder(encoding)(errors="strict")
            return decoder.decode(data, final=final), encoding
        except (UnicodeDecodeError, LookupError):
            continue

    return None


def _detect_delimited_format(text: str) -> str | None:
    """Recognize bounded CSV/TSV content using logical, quote-aware records."""
    sample = text[:_DELIMITED_SNIFF_SIZE]
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
        reader = csv.reader(io.StringIO(sample), dialect)
        records: list[list[str]] = []
        for row in reader:
            if not row or not any(cell.strip() for cell in row):
                continue
            if len(row) == 1 and row[0].lstrip().startswith("#"):
                continue
            records.append(row)
            if len(records) >= _DELIMITED_RECORD_LIMIT:
                break
    except (csv.Error, UnicodeError):
        return None

    if len(records) < 2:
        return None
    widths = [len(row) for row in records]
    if any(width < 2 for width in widths):
        return None
    if len(set(widths)) > 2 and max(widths) - min(widths) > 3:
        return None
    return "tsv" if dialect.delimiter == "\t" else "csv"


def _inspect_zip_container(file_path: str) -> ContentDetection:
    """Return a content-only result for a ZIP-based input.

    A marker filename is not sufficient evidence.  Every supported package is
    admitted only after its minimum relationship graph and XML parts have been
    read and parsed from the archive itself.  Nothing is extracted to disk.
    """

    candidate_format = ""
    try:
        with zipfile.ZipFile(file_path, "r") as package:
            index = _zip_part_index(package)
            candidate_format = _zip_candidate_format(package, index)
            if candidate_format:
                _validate_zip_document_package(package, index, candidate_format)
                return ContentDetection(
                    format=candidate_format,
                    method=DetectionMethod.CONTAINER,
                    confidence=DetectionConfidence.CERTAIN,
                    structure_status=StructureStatus.VALID,
                )
    except (
        EOFError,
        NotImplementedError,
        OSError,
        RuntimeError,
        ValueError,
        zipfile.BadZipFile,
        ElementTree.ParseError,
    ) as exc:
        label = candidate_format.upper() if candidate_format else "ZIP-based"
        return ContentDetection(
            format=candidate_format or "unknown",
            method=DetectionMethod.CONTAINER,
            confidence=DetectionConfidence.UNVERIFIED,
            structure_status=StructureStatus.INVALID,
            detail_code="FILE_CONTAINER_INVALID",
            detail_message=f"The {label} container is corrupt or structurally invalid: {exc}",
        )

    return ContentDetection(
        format="zip",
        method=DetectionMethod.CONTAINER,
        confidence=DetectionConfidence.CERTAIN,
        structure_status=StructureStatus.VALID,
        detail_code="FILE_CONTAINER_UNSUPPORTED",
        detail_message="The ZIP container does not contain a supported document package structure.",
    )


def _zip_part_index(package: zipfile.ZipFile) -> dict[str, tuple[zipfile.ZipInfo, ...]]:
    """Index exact member names while preserving duplicate entries."""

    parts: dict[str, list[zipfile.ZipInfo]] = {}
    for info in package.infolist():
        if info.is_dir():
            continue
        parts.setdefault(info.filename, []).append(info)
    return {name: tuple(entries) for name, entries in parts.items()}


def _zip_candidate_format(
    package: zipfile.ZipFile,
    index: dict[str, tuple[zipfile.ZipInfo, ...]],
) -> str:
    """Return the format whose identifying package marker must be validated."""

    if "OFD.xml" in index:
        return "ofd"

    for file_format, (main_part, _content_type, _root_name) in _OOXML_MAIN_PARTS.items():
        if main_part in index:
            return file_format

    mime = _optional_mimetype(package, index)
    if mime == "application/epub+zip" or "META-INF/container.xml" in index:
        return "epub"
    if mime in _ZIP_MIMETYPE_MAP:
        return _ZIP_MIMETYPE_MAP[mime]
    if "FixedDocumentSequence.fdseq" in index:
        return "xps"
    return ""


def _validate_zip_document_package(
    package: zipfile.ZipFile,
    index: dict[str, tuple[zipfile.ZipInfo, ...]],
    file_format: str,
) -> None:
    if file_format in _OOXML_MAIN_PARTS:
        _validate_ooxml_package(package, index, file_format)
    elif file_format in _ZIP_MIMETYPE_MAP.values():
        _validate_odf_package(package, index, file_format)
    elif file_format == "epub":
        _validate_epub_package(package, index)
    elif file_format == "ofd":
        _validate_ofd_package(package, index)
    elif file_format == "xps":
        _validate_xps_package(package, index)
    else:  # pragma: no cover - guarded by _zip_candidate_format
        raise ValueError(f"unsupported document package candidate: {file_format}")


def _validate_ooxml_package(
    package: zipfile.ZipFile,
    index: dict[str, tuple[zipfile.ZipInfo, ...]],
    file_format: str,
) -> None:
    main_part, expected_content_type, expected_root_name = _OOXML_MAIN_PARTS[file_format]
    content_types = _read_xml_part(package, index, "[Content_Types].xml")
    relationships = _read_xml_part(package, index, "_rels/.rels")
    main_root = _read_xml_part(package, index, main_part)

    _require_xml_root(content_types, "Types", "[Content_Types].xml")
    _require_xml_root(relationships, "Relationships", "_rels/.rels")
    _require_xml_root(main_root, expected_root_name, main_part)

    matching_overrides = [
        node
        for node in content_types.iter()
        if _xml_local_name(node.tag) == "Override"
        and _xml_attribute(node, "PartName") == f"/{main_part}"
        and _xml_attribute(node, "ContentType") == expected_content_type
    ]
    if len(matching_overrides) != 1:
        raise ValueError(f"[Content_Types].xml must declare exactly one content type for {main_part}")

    office_targets: list[str] = []
    for node in relationships.iter():
        if _xml_local_name(node.tag) != "Relationship":
            continue
        relationship_type = _xml_attribute(node, "Type")
        if not relationship_type.endswith("/officeDocument"):
            continue
        if _xml_attribute(node, "TargetMode").lower() == "external":
            raise ValueError("the officeDocument relationship must not be external")
        target = _resolve_package_reference(_xml_attribute(node, "Target"), allow_package_absolute=True)
        office_targets.append(target)

    if office_targets != [main_part]:
        raise ValueError(f"_rels/.rels must reference exactly one {main_part} officeDocument part")


def _validate_odf_package(
    package: zipfile.ZipFile,
    index: dict[str, tuple[zipfile.ZipInfo, ...]],
    file_format: str,
) -> None:
    expected_mimetype = next(mime for mime, mapped_format in _ZIP_MIMETYPE_MAP.items() if mapped_format == file_format)
    mimetype = _required_mimetype(package, index)
    if mimetype != expected_mimetype:
        raise ValueError(f"mimetype does not declare {file_format.upper()}")

    manifest = _read_xml_part(package, index, "META-INF/manifest.xml")
    content = _read_xml_part(package, index, "content.xml")
    _require_xml_root(manifest, "manifest", "META-INF/manifest.xml")
    _require_xml_root(content, "document-content", "content.xml")

    content_entries = [
        node
        for node in manifest.iter()
        if _xml_local_name(node.tag) == "file-entry" and _xml_attribute(node, "full-path") == "content.xml"
    ]
    if len(content_entries) != 1:
        raise ValueError("META-INF/manifest.xml must reference content.xml exactly once")


def _validate_epub_package(
    package: zipfile.ZipFile,
    index: dict[str, tuple[zipfile.ZipInfo, ...]],
) -> None:
    if _required_mimetype(package, index) != "application/epub+zip":
        raise ValueError("mimetype must be application/epub+zip")

    container = _read_xml_part(package, index, "META-INF/container.xml")
    _require_xml_root(container, "container", "META-INF/container.xml")
    rootfiles = [
        node
        for node in container.iter()
        if _xml_local_name(node.tag) == "rootfile"
        and _xml_attribute(node, "media-type") == "application/oebps-package+xml"
    ]
    if not rootfiles:
        raise ValueError("META-INF/container.xml does not reference an EPUB package rootfile")

    referenced_parts = {_resolve_package_reference(_xml_attribute(node, "full-path")) for node in rootfiles}
    for rootfile in referenced_parts:
        package_root = _read_xml_part(package, index, rootfile)
        _require_xml_root(package_root, "package", rootfile)


def _validate_ofd_package(
    package: zipfile.ZipFile,
    index: dict[str, tuple[zipfile.ZipInfo, ...]],
) -> None:
    descriptor = _read_xml_part(package, index, "OFD.xml")
    _require_xml_root(descriptor, "OFD", "OFD.xml")
    document_roots = [
        (node.text or "").strip()
        for node in descriptor.iter()
        if _xml_local_name(node.tag) == "DocRoot" and (node.text or "").strip()
    ]
    if not document_roots:
        raise ValueError("OFD.xml does not reference a DocRoot")

    referenced_parts = {_resolve_package_reference(target) for target in document_roots}
    for document_root in referenced_parts:
        document = _read_xml_part(package, index, document_root)
        _require_xml_root(document, "Document", document_root)


def _validate_xps_package(
    package: zipfile.ZipFile,
    index: dict[str, tuple[zipfile.ZipInfo, ...]],
) -> None:
    content_types = _read_xml_part(package, index, "[Content_Types].xml")
    sequence = _read_xml_part(package, index, "FixedDocumentSequence.fdseq")
    _require_xml_root(content_types, "Types", "[Content_Types].xml")
    _require_xml_root(sequence, "FixedDocumentSequence", "FixedDocumentSequence.fdseq")

    declarations = [
        node
        for node in content_types.iter()
        if (
            _xml_local_name(node.tag) == "Override"
            and _xml_attribute(node, "PartName") == "/FixedDocumentSequence.fdseq"
            and _xml_attribute(node, "ContentType") == _XPS_FIXED_DOCUMENT_SEQUENCE_CONTENT_TYPE
        )
        or (
            _xml_local_name(node.tag) == "Default"
            and _xml_attribute(node, "Extension").lower() == "fdseq"
            and _xml_attribute(node, "ContentType") == _XPS_FIXED_DOCUMENT_SEQUENCE_CONTENT_TYPE
        )
    ]
    if len(declarations) != 1:
        raise ValueError("[Content_Types].xml must declare FixedDocumentSequence.fdseq exactly once")
    if _xml_namespace(sequence.tag) != _XPS_FIXED_DOCUMENT_SEQUENCE_NAMESPACE:
        raise ValueError(
            "FixedDocumentSequence.fdseq must use the Microsoft XPS 2005/06 namespace; OpenXPS is unsupported"
        )


def _optional_mimetype(
    package: zipfile.ZipFile,
    index: dict[str, tuple[zipfile.ZipInfo, ...]],
) -> str:
    if "mimetype" not in index:
        return ""
    return _required_mimetype(package, index)


def _required_mimetype(
    package: zipfile.ZipFile,
    index: dict[str, tuple[zipfile.ZipInfo, ...]],
) -> str:
    payload = _read_zip_part(package, index, "mimetype", max_bytes=_MAX_PACKAGE_MIMETYPE_BYTES)
    try:
        return payload.decode("ascii")
    except UnicodeDecodeError as exc:
        raise ValueError("mimetype is not ASCII text") from exc


def _read_xml_part(
    package: zipfile.ZipFile,
    index: dict[str, tuple[zipfile.ZipInfo, ...]],
    part_name: str,
) -> ElementTree.Element:
    payload = _read_zip_part(package, index, part_name, max_bytes=_MAX_PACKAGE_XML_BYTES)
    try:
        return ElementTree.fromstring(payload)
    except ElementTree.ParseError as exc:
        raise ValueError(f"{part_name} is not well-formed XML: {exc}") from exc


def _read_zip_part(
    package: zipfile.ZipFile,
    index: dict[str, tuple[zipfile.ZipInfo, ...]],
    part_name: str,
    *,
    max_bytes: int,
) -> bytes:
    entries = index.get(part_name, ())
    if not entries:
        raise ValueError(f"missing required part: {part_name}")
    if len(entries) != 1:
        raise ValueError(f"duplicate required part: {part_name}")
    info = entries[0]
    if info.file_size > max_bytes:
        raise ValueError(f"required part exceeds validation limit: {part_name}")
    try:
        with package.open(info, "r") as stream:
            payload = stream.read(max_bytes + 1)
    except (EOFError, NotImplementedError, OSError, RuntimeError, zipfile.BadZipFile) as exc:
        raise ValueError(f"required part cannot be read: {part_name}: {exc}") from exc
    if len(payload) > max_bytes:
        raise ValueError(f"required part exceeds validation limit: {part_name}")
    return payload


def _resolve_package_reference(reference: str, *, allow_package_absolute: bool = False) -> str:
    raw_reference = reference.strip()
    if not raw_reference:
        raise ValueError("package reference is empty")
    if "\\" in raw_reference or "\x00" in raw_reference or "?" in raw_reference or "#" in raw_reference:
        raise ValueError(f"package reference is not a plain internal path: {raw_reference}")
    if ":" in raw_reference.split("/", 1)[0]:
        raise ValueError(f"external package reference is not allowed: {raw_reference}")
    if raw_reference.startswith("/"):
        if not allow_package_absolute or raw_reference.startswith("//"):
            raise ValueError(f"absolute package reference is not allowed: {raw_reference}")
        raw_reference = raw_reference[1:]

    normalized = posixpath.normpath(raw_reference)
    if normalized in {"", ".", ".."} or normalized.startswith(("../", "/")):
        raise ValueError(f"package reference escapes the archive: {reference}")
    return normalized


def _require_xml_root(root: ElementTree.Element, expected_local_name: str, part_name: str) -> None:
    if _xml_local_name(root.tag) != expected_local_name:
        raise ValueError(f"{part_name} has an unexpected XML root element")


def _xml_local_name(name: str) -> str:
    return name.rsplit("}", 1)[-1]


def _xml_namespace(name: str) -> str:
    if name.startswith("{") and "}" in name:
        return name[1:].split("}", 1)[0]
    return ""


def _xml_attribute(node: ElementTree.Element, local_name: str) -> str:
    for name, value in node.attrib.items():
        if _xml_local_name(name) == local_name:
            return value.strip()
    return ""


def _detect_ole_type(file_path: str) -> str | None:
    """Inspect OLE2 compound document streams to distinguish DOC/XLS/PPT."""
    try:
        import olefile
    except ImportError:
        return None

    try:
        if not olefile.isOleFile(file_path):
            return None
        ole = olefile.OleFileIO(file_path)
        streams = ole.listdir()
        flat = {s[0] for s in streams if s}
    except Exception:
        return None

    for stream_names, fmt in _OLE_STREAM_SIGNATURES:
        if any(s in flat for s in stream_names):
            return fmt

    return None

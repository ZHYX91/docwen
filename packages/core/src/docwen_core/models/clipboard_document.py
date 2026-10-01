"""Versioned recursive clipboard-document model for managed GUI snapshots.

This is an internal application/runtime contract, not a public Machine Protocol
surface. Ordinary JSON files are never admitted as this format; only the
dedicated managed-snapshot inspector may produce `clipboard_document` input.
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass
from typing import TYPE_CHECKING, Any, Literal

if TYPE_CHECKING:
    from collections.abc import Sequence

    from docwen_core.models.file_ref import FileRef

from docwen_core.models.semantic_document import (
    SemanticDocument,
    SemanticDocumentValidationError,
    SemanticTable,
    SemanticTableCell,
    derive_table_header_shape,
    validate_semantic_document,
)

CLIPBOARD_DOCUMENT_SCHEMA = "docwen.clipboard_document.v2"
CLIPBOARD_DOCUMENT_FORMAT = "clipboard_document"
CLIPBOARD_DOCUMENT_MEDIA_TYPE = "application/vnd.docwen.clipboard-document+json"

MAX_CLIPBOARD_DOCUMENT_BYTES = 4 * 1024 * 1024
MAX_CLIPBOARD_BLOCKS = 4096
MAX_CLIPBOARD_TABLES = 256
MAX_CLIPBOARD_CELLS = 65536
MAX_CLIPBOARD_DEPTH = 8
MAX_CLIPBOARD_INLINES = 262144
MAX_CLIPBOARD_TEXT_CODEPOINTS = 4_000_000
MAX_CLIPBOARD_RESOURCES = 256
MAX_CLIPBOARD_RESOURCE_BYTES = 64 * 1024 * 1024
MAX_CLIPBOARD_IMAGE_PIXELS = 64 * 1024 * 1024
MAX_CLIPBOARD_IMAGE_SIDE = 32768
MAX_CLIPBOARD_IMAGE_PIXELS_PER_RESOURCE = 32 * 1024 * 1024

_RESOURCE_ID_RE = re.compile(r"^[A-Za-z][A-Za-z0-9._-]{0,127}$")
_HEADER_SCOPES = frozenset({"", "row", "col", "rowgroup", "colgroup"})


class ClipboardDocumentError(ValueError):
    """A structured clipboard payload violates its closed internal contract."""

    def __init__(self, code: str, message: str) -> None:
        self.code = code
        super().__init__(message)


@dataclass(frozen=True, slots=True)
class ClipboardText:
    value: str


@dataclass(frozen=True, slots=True)
class ClipboardHardBreak:
    pass


@dataclass(frozen=True, slots=True)
class ClipboardImageRef:
    """One image occurrence; missing nodes keep their original document position."""

    resource_id: str | None
    alt: str
    missing_reason: str = ""
    extent_cx_emu: int | None = None
    extent_cy_emu: int | None = None


type ClipboardInline = ClipboardText | ClipboardHardBreak | ClipboardImageRef


@dataclass(frozen=True, slots=True)
class ClipboardParagraph:
    inlines: tuple[ClipboardInline, ...]


@dataclass(frozen=True, slots=True)
class ClipboardTableCell:
    row: int
    column: int
    row_span: int
    column_span: int
    blocks: tuple[ClipboardBlock, ...]
    header: bool = False
    scope: str = ""
    headers: tuple[str, ...] = ()
    html_id: str = ""


@dataclass(frozen=True, slots=True)
class ClipboardTable:
    row_count: int
    column_count: int
    cells: tuple[ClipboardTableCell, ...]


type ClipboardBlock = ClipboardParagraph | ClipboardTable


@dataclass(frozen=True, slots=True)
class ClipboardResource:
    resource_id: str
    logical_path: str
    media_type: str
    size_bytes: int
    sha256: str
    pixel_width: int | None = None
    pixel_height: int | None = None
    rgba_sha256: str = ""


@dataclass(frozen=True, slots=True)
class ClipboardDocument:
    blocks: tuple[ClipboardBlock, ...]
    resources: tuple[ClipboardResource, ...] = ()


def _xml_text(value: str) -> bool:
    return all(
        character in {"\t", "\n", "\r"}
        or "\u0020" <= character <= "\ud7ff"
        or "\ue000" <= character <= "\ufffd"
        or "\U00010000" <= character <= "\U0010ffff"
        for character in value
    )


def _require_keys(value: dict[str, Any], allowed: set[str], required: set[str], *, where: str) -> None:
    extra = set(value) - allowed
    missing = required - set(value)
    if extra:
        raise ClipboardDocumentError("clipboard.unknown_field", f"{where} has unsupported field(s): {sorted(extra)}")
    if missing:
        raise ClipboardDocumentError("clipboard.missing_field", f"{where} is missing field(s): {sorted(missing)}")


def _require_int(value: object, *, minimum: int, maximum: int, where: str) -> int:
    if type(value) is not int or not minimum <= value <= maximum:
        raise ClipboardDocumentError("clipboard.integer_invalid", f"{where} must be an integer in range.")
    return value


def _require_string(value: object, *, where: str, allow_empty: bool = True) -> str:
    if type(value) is not str or (not allow_empty and not value):
        raise ClipboardDocumentError("clipboard.string_invalid", f"{where} must be a string.")
    if not _xml_text(value):
        raise ClipboardDocumentError("clipboard.text_invalid", f"{where} contains unsupported control characters.")
    return value


def _parse_inline(data: object, counters: dict[str, int], *, where: str) -> ClipboardInline:
    if not isinstance(data, dict):
        raise ClipboardDocumentError("clipboard.inline_invalid", f"{where} must be an object.")
    kind = data.get("type")
    counters["inlines"] += 1
    if counters["inlines"] > MAX_CLIPBOARD_INLINES:
        raise ClipboardDocumentError("clipboard.budget_exceeded", "Clipboard inline budget exceeded.")
    if kind == "text":
        _require_keys(data, {"type", "value"}, {"type", "value"}, where=where)
        value = _require_string(data["value"], where=f"{where}.value")
        counters["text"] += len(value)
        if counters["text"] > MAX_CLIPBOARD_TEXT_CODEPOINTS:
            raise ClipboardDocumentError("clipboard.budget_exceeded", "Clipboard text budget exceeded.")
        return ClipboardText(value)
    if kind == "hard_break":
        _require_keys(data, {"type"}, {"type"}, where=where)
        return ClipboardHardBreak()
    if kind == "image":
        _require_keys(
            data,
            {"type", "resourceId", "alt", "missingReason", "extentCxEmu", "extentCyEmu"},
            {"type", "alt"},
            where=where,
        )
        resource_id = data.get("resourceId")
        if resource_id is not None:
            resource_id = _require_string(resource_id, where=f"{where}.resourceId", allow_empty=False)
            if _RESOURCE_ID_RE.fullmatch(resource_id) is None:
                raise ClipboardDocumentError("clipboard.resource_id_invalid", f"{where}.resourceId is invalid.")
        alt = _require_string(data.get("alt", ""), where=f"{where}.alt")
        reason = _require_string(data.get("missingReason", ""), where=f"{where}.missingReason")
        counters["text"] += len(alt)
        if counters["text"] > MAX_CLIPBOARD_TEXT_CODEPOINTS:
            raise ClipboardDocumentError("clipboard.budget_exceeded", "Clipboard text budget exceeded.")
        if resource_id is None and not reason:
            raise ClipboardDocumentError(
                "clipboard.image_binding_invalid", "An image without resource bytes must carry a missing reason."
            )
        if resource_id is not None and reason:
            raise ClipboardDocumentError(
                "clipboard.image_binding_invalid", "An image cannot be both resource-bound and missing."
            )
        raw_cx = data.get("extentCxEmu")
        raw_cy = data.get("extentCyEmu")
        if (raw_cx is None) != (raw_cy is None):
            raise ClipboardDocumentError(
                "clipboard.image_extent_invalid", "Image extent must provide both EMU dimensions."
            )
        extent_cx = (
            None
            if raw_cx is None
            else _require_int(raw_cx, minimum=1, maximum=2_147_483_647, where=f"{where}.extentCxEmu")
        )
        extent_cy = (
            None
            if raw_cy is None
            else _require_int(raw_cy, minimum=1, maximum=2_147_483_647, where=f"{where}.extentCyEmu")
        )
        return ClipboardImageRef(resource_id, alt, reason, extent_cx, extent_cy)
    raise ClipboardDocumentError("clipboard.inline_type_invalid", f"{where} has unsupported inline type.")


def _parse_block(data: object, counters: dict[str, int], *, depth: int, where: str) -> ClipboardBlock:
    if depth > MAX_CLIPBOARD_DEPTH:
        raise ClipboardDocumentError("clipboard.depth_exceeded", "Clipboard nesting depth exceeded.")
    if not isinstance(data, dict):
        raise ClipboardDocumentError("clipboard.block_invalid", f"{where} must be an object.")
    counters["blocks"] += 1
    if counters["blocks"] > MAX_CLIPBOARD_BLOCKS:
        raise ClipboardDocumentError("clipboard.budget_exceeded", "Clipboard block budget exceeded.")
    kind = data.get("type")
    if kind == "paragraph":
        _require_keys(data, {"type", "inlines"}, {"type", "inlines"}, where=where)
        raw_inlines = data["inlines"]
        if not isinstance(raw_inlines, list):
            raise ClipboardDocumentError("clipboard.inlines_invalid", f"{where}.inlines must be an array.")
        return ClipboardParagraph(
            tuple(
                _parse_inline(item, counters, where=f"{where}.inlines[{index}]")
                for index, item in enumerate(raw_inlines)
            )
        )
    if kind != "table":
        raise ClipboardDocumentError("clipboard.block_type_invalid", f"{where} has unsupported block type.")

    _require_keys(
        data,
        {"type", "rowCount", "columnCount", "cells"},
        {"type", "rowCount", "columnCount", "cells"},
        where=where,
    )
    counters["tables"] += 1
    if counters["tables"] > MAX_CLIPBOARD_TABLES:
        raise ClipboardDocumentError("clipboard.budget_exceeded", "Clipboard table budget exceeded.")
    row_count = _require_int(data["rowCount"], minimum=1, maximum=MAX_CLIPBOARD_CELLS, where=f"{where}.rowCount")
    column_count = _require_int(
        data["columnCount"], minimum=1, maximum=MAX_CLIPBOARD_CELLS, where=f"{where}.columnCount"
    )
    if row_count * column_count > MAX_CLIPBOARD_CELLS:
        raise ClipboardDocumentError("clipboard.budget_exceeded", "Clipboard table grid budget exceeded.")
    raw_cells = data["cells"]
    if not isinstance(raw_cells, list):
        raise ClipboardDocumentError("clipboard.cells_invalid", f"{where}.cells must be an array.")
    counters["cells"] += len(raw_cells)
    if counters["cells"] > MAX_CLIPBOARD_CELLS:
        raise ClipboardDocumentError("clipboard.budget_exceeded", "Clipboard cell budget exceeded.")
    cells: list[ClipboardTableCell] = []
    for index, raw_cell in enumerate(raw_cells):
        cell_where = f"{where}.cells[{index}]"
        if not isinstance(raw_cell, dict):
            raise ClipboardDocumentError("clipboard.cell_invalid", f"{cell_where} must be an object.")
        _require_keys(
            raw_cell,
            {"row", "column", "rowSpan", "columnSpan", "blocks", "header", "scope", "headers", "htmlId"},
            {"row", "column", "rowSpan", "columnSpan", "blocks"},
            where=cell_where,
        )
        raw_blocks = raw_cell["blocks"]
        if not isinstance(raw_blocks, list):
            raise ClipboardDocumentError("clipboard.blocks_invalid", f"{cell_where}.blocks must be an array.")
        header = raw_cell.get("header", False)
        if type(header) is not bool:
            raise ClipboardDocumentError("clipboard.header_invalid", f"{cell_where}.header must be a boolean.")
        scope = _require_string(raw_cell.get("scope", ""), where=f"{cell_where}.scope").lower()
        if scope not in _HEADER_SCOPES:
            raise ClipboardDocumentError("clipboard.header_scope_invalid", f"{cell_where}.scope is invalid.")
        raw_headers = raw_cell.get("headers", [])
        if not isinstance(raw_headers, list):
            raise ClipboardDocumentError("clipboard.headers_invalid", f"{cell_where}.headers must be an array.")
        headers = tuple(_require_string(item, where=f"{cell_where}.headers", allow_empty=False) for item in raw_headers)
        html_id = _require_string(raw_cell.get("htmlId", ""), where=f"{cell_where}.htmlId")
        cells.append(
            ClipboardTableCell(
                row=_require_int(raw_cell["row"], minimum=0, maximum=row_count - 1, where=f"{cell_where}.row"),
                column=_require_int(
                    raw_cell["column"], minimum=0, maximum=column_count - 1, where=f"{cell_where}.column"
                ),
                row_span=_require_int(raw_cell["rowSpan"], minimum=1, maximum=row_count, where=f"{cell_where}.rowSpan"),
                column_span=_require_int(
                    raw_cell["columnSpan"], minimum=1, maximum=column_count, where=f"{cell_where}.columnSpan"
                ),
                blocks=tuple(
                    _parse_block(item, counters, depth=depth + 1, where=f"{cell_where}.blocks[{block_index}]")
                    for block_index, item in enumerate(raw_blocks)
                ),
                header=header,
                scope=scope,
                headers=headers,
                html_id=html_id,
            )
        )
    table = ClipboardTable(row_count, column_count, tuple(cells))
    _validate_table_geometry(table, where=where)
    return table


def _parse_resource(data: object, *, where: str) -> ClipboardResource:
    if not isinstance(data, dict):
        raise ClipboardDocumentError("clipboard.resource_invalid", f"{where} must be an object.")
    _require_keys(
        data,
        {
            "resourceId",
            "logicalPath",
            "mediaType",
            "sizeBytes",
            "sha256",
            "pixelWidth",
            "pixelHeight",
            "rgbaSha256",
        },
        {"resourceId", "logicalPath", "mediaType", "sizeBytes", "sha256"},
        where=where,
    )
    resource_id = _require_string(data["resourceId"], where=f"{where}.resourceId", allow_empty=False)
    if _RESOURCE_ID_RE.fullmatch(resource_id) is None:
        raise ClipboardDocumentError("clipboard.resource_id_invalid", f"{where}.resourceId is invalid.")
    logical_path = _require_string(data["logicalPath"], where=f"{where}.logicalPath", allow_empty=False)
    segments = logical_path.split("/")
    if logical_path.startswith("/") or "\\" in logical_path or any(segment in {"", ".", ".."} for segment in segments):
        raise ClipboardDocumentError("clipboard.resource_path_invalid", f"{where}.logicalPath is invalid.")
    media_type = _require_string(data["mediaType"], where=f"{where}.mediaType", allow_empty=False)
    size_bytes = _require_int(
        data["sizeBytes"], minimum=0, maximum=MAX_CLIPBOARD_RESOURCE_BYTES, where=f"{where}.sizeBytes"
    )
    sha256 = _require_string(data["sha256"], where=f"{where}.sha256", allow_empty=False).lower()
    if re.fullmatch(r"[0-9a-f]{64}", sha256) is None:
        raise ClipboardDocumentError("clipboard.resource_hash_invalid", f"{where}.sha256 is invalid.")
    raw_width = data.get("pixelWidth")
    raw_height = data.get("pixelHeight")
    rgba_sha256 = _require_string(data.get("rgbaSha256", ""), where=f"{where}.rgbaSha256")
    if (raw_width is None) != (raw_height is None):
        raise ClipboardDocumentError("clipboard.resource_pixels_invalid", "Image pixel facts must provide both dimensions.")
    pixel_width = (
        None
        if raw_width is None
        else _require_int(raw_width, minimum=1, maximum=MAX_CLIPBOARD_IMAGE_SIDE, where=f"{where}.pixelWidth")
    )
    pixel_height = (
        None
        if raw_height is None
        else _require_int(raw_height, minimum=1, maximum=MAX_CLIPBOARD_IMAGE_SIDE, where=f"{where}.pixelHeight")
    )
    if pixel_width is not None and pixel_height is not None:
        if pixel_width * pixel_height > MAX_CLIPBOARD_IMAGE_PIXELS_PER_RESOURCE:
            raise ClipboardDocumentError("clipboard.budget_exceeded", "Clipboard image pixel budget exceeded.")
        if media_type != "image/png" or re.fullmatch(r"[0-9a-f]{64}", rgba_sha256) is None:
            raise ClipboardDocumentError(
                "clipboard.resource_pixels_invalid",
                "Clipboard image pixel facts require PNG media and an RGBA SHA-256.",
            )
    elif rgba_sha256:
        raise ClipboardDocumentError(
            "clipboard.resource_pixels_invalid",
            "RGBA SHA-256 requires image pixel dimensions.",
        )
    return ClipboardResource(
        resource_id,
        logical_path,
        media_type,
        size_bytes,
        sha256,
        pixel_width,
        pixel_height,
        rgba_sha256,
    )


def load_clipboard_document_bytes(payload: bytes) -> ClipboardDocument:
    """Load and strictly validate a managed structured clipboard snapshot."""

    if not isinstance(payload, bytes):
        raise TypeError("structured clipboard payload must be bytes")
    if not payload or len(payload) > MAX_CLIPBOARD_DOCUMENT_BYTES:
        raise ClipboardDocumentError("clipboard.payload_size_invalid", "Structured clipboard payload size is invalid.")
    try:
        data = json.loads(payload.decode("utf-8", errors="strict"))
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ClipboardDocumentError(
            "clipboard.json_invalid", "Structured clipboard payload is not valid UTF-8 JSON."
        ) from exc
    if not isinstance(data, dict):
        raise ClipboardDocumentError("clipboard.root_invalid", "Structured clipboard payload must be an object.")
    _require_keys(data, {"schema", "blocks", "resources"}, {"schema", "blocks"}, where="document")
    if data.get("schema") != CLIPBOARD_DOCUMENT_SCHEMA:
        raise ClipboardDocumentError("clipboard.schema_invalid", "Unsupported structured clipboard schema.")
    raw_blocks = data["blocks"]
    if not isinstance(raw_blocks, list):
        raise ClipboardDocumentError("clipboard.blocks_invalid", "document.blocks must be an array.")
    raw_resources = data.get("resources", [])
    if not isinstance(raw_resources, list) or len(raw_resources) > MAX_CLIPBOARD_RESOURCES:
        raise ClipboardDocumentError(
            "clipboard.resources_invalid", "document.resources exceeds the supported resource budget."
        )
    counters = {"blocks": 0, "tables": 0, "cells": 0, "inlines": 0, "text": 0}
    document = ClipboardDocument(
        blocks=tuple(
            _parse_block(item, counters, depth=1, where=f"document.blocks[{index}]")
            for index, item in enumerate(raw_blocks)
        ),
        resources=tuple(
            _parse_resource(item, where=f"document.resources[{index}]") for index, item in enumerate(raw_resources)
        ),
    )
    resource_ids = {item.resource_id for item in document.resources}
    if len(resource_ids) != len(document.resources):
        raise ClipboardDocumentError("clipboard.resource_duplicate", "Resource IDs must be unique.")
    if len({item.logical_path for item in document.resources}) != len(document.resources):
        raise ClipboardDocumentError("clipboard.resource_path_duplicate", "Resource logical paths must be unique.")
    if sum(item.size_bytes for item in document.resources) > MAX_CLIPBOARD_RESOURCE_BYTES:
        raise ClipboardDocumentError("clipboard.budget_exceeded", "Clipboard resource byte budget exceeded.")
    total_pixels = sum(
        item.pixel_width * item.pixel_height
        for item in document.resources
        if item.pixel_width is not None and item.pixel_height is not None
    )
    if total_pixels > MAX_CLIPBOARD_IMAGE_PIXELS:
        raise ClipboardDocumentError("clipboard.budget_exceeded", "Clipboard total image pixel budget exceeded.")
    for inline in iter_clipboard_inlines(document):
        if (
            isinstance(inline, ClipboardImageRef)
            and inline.resource_id is not None
            and inline.resource_id not in resource_ids
        ):
            raise ClipboardDocumentError(
                "clipboard.resource_missing", "Image reference points to an undeclared resource."
            )
    return document


def clipboard_document_to_bytes(document: ClipboardDocument) -> bytes:
    """Serialize one validated clipboard document deterministically."""

    payload = json.dumps(
        clipboard_document_to_dict(document),
        ensure_ascii=False,
        separators=(",", ":"),
        sort_keys=True,
    ).encode("utf-8")
    load_clipboard_document_bytes(payload)
    return payload


def clipboard_document_to_dict(document: ClipboardDocument) -> dict[str, Any]:
    return {
        "schema": CLIPBOARD_DOCUMENT_SCHEMA,
        "blocks": [_block_to_dict(block) for block in document.blocks],
        "resources": [
            {
                "resourceId": item.resource_id,
                "logicalPath": item.logical_path,
                "mediaType": item.media_type,
                "sizeBytes": item.size_bytes,
                "sha256": item.sha256,
                **({"pixelWidth": item.pixel_width} if item.pixel_width is not None else {}),
                **({"pixelHeight": item.pixel_height} if item.pixel_height is not None else {}),
                **({"rgbaSha256": item.rgba_sha256} if item.rgba_sha256 else {}),
            }
            for item in document.resources
        ],
    }


def _block_to_dict(block: ClipboardBlock) -> dict[str, Any]:
    if isinstance(block, ClipboardParagraph):
        return {"type": "paragraph", "inlines": [_inline_to_dict(item) for item in block.inlines]}
    return {
        "type": "table",
        "rowCount": block.row_count,
        "columnCount": block.column_count,
        "cells": [
            {
                "row": cell.row,
                "column": cell.column,
                "rowSpan": cell.row_span,
                "columnSpan": cell.column_span,
                "blocks": [_block_to_dict(item) for item in cell.blocks],
                "header": cell.header,
                "scope": cell.scope,
                "headers": list(cell.headers),
                "htmlId": cell.html_id,
            }
            for cell in block.cells
        ],
    }


def _inline_to_dict(inline: ClipboardInline) -> dict[str, Any]:
    if isinstance(inline, ClipboardText):
        return {"type": "text", "value": inline.value}
    if isinstance(inline, ClipboardHardBreak):
        return {"type": "hard_break"}
    return {
        "type": "image",
        "resourceId": inline.resource_id,
        "alt": inline.alt,
        "missingReason": inline.missing_reason,
        **({"extentCxEmu": inline.extent_cx_emu} if inline.extent_cx_emu is not None else {}),
        **({"extentCyEmu": inline.extent_cy_emu} if inline.extent_cy_emu is not None else {}),
    }


def _validate_table_geometry(table: ClipboardTable, *, where: str) -> None:
    semantic = SemanticTable(
        row_count=table.row_count,
        column_count=table.column_count,
        cells=tuple(
            SemanticTableCell(
                row=cell.row,
                column=cell.column,
                text=clipboard_cell_text(cell),
                row_span=cell.row_span,
                column_span=cell.column_span,
            )
            for cell in table.cells
        ),
    )
    diagnostics = validate_semantic_document(SemanticDocument(blocks=(semantic,)))
    if diagnostics:
        message = "; ".join(f"{item.code}@{item.location}" for item in diagnostics[:4])
        raise ClipboardDocumentError(
            "clipboard.table_geometry_invalid", f"{where} has invalid table geometry: {message}"
        )


def _native_header_candidate(table: ClipboardTable, header_rows: int, header_columns: int) -> bool:
    semantic_cells: list[SemanticTableCell] = []
    for cell in table.cells:
        rows = tuple(range(cell.row, cell.row + cell.row_span))
        columns = tuple(range(cell.column, cell.column + cell.column_span))
        in_rows = bool(header_rows) and all(row < header_rows for row in rows)
        in_columns = bool(header_columns) and all(column < header_columns for column in columns)
        if header_rows and any(row < header_rows for row in rows) and any(row >= header_rows for row in rows):
            return False
        if (
            header_columns
            and any(column < header_columns for column in columns)
            and any(column >= header_columns for column in columns)
        ):
            return False
        native_header = in_rows or in_columns
        if not cell.header and native_header:
            return False
        if cell.header and cell.scope in {"col", "colgroup"} and not in_rows:
            return False
        if cell.header and cell.scope in {"row", "rowgroup"} and not in_columns:
            return False
        if cell.header and not cell.scope and not native_header:
            return False
        role: Literal["data", "column_header", "row_header", "corner_header"] = "data"
        if in_rows and in_columns:
            role = "corner_header"
        elif in_rows:
            role = "column_header"
        elif in_columns:
            role = "row_header"
        semantic_cells.append(
            SemanticTableCell(
                row=cell.row,
                column=cell.column,
                text=clipboard_cell_text(cell),
                role=role,
                row_span=cell.row_span,
                column_span=cell.column_span,
            )
        )
    try:
        derived = derive_table_header_shape(SemanticTable(table.row_count, table.column_count, tuple(semantic_cells)))
    except SemanticDocumentValidationError:
        return False
    return derived == (header_rows, header_columns)


def clipboard_table_header_shape(table: ClipboardTable) -> tuple[int, int]:
    """Return native header prefixes justified by explicit th/scope evidence.

    scope=row/rowgroup is row-header evidence, never evidence for a repeated
    Word column-header row. td cells are never promoted merely by position.
    Association graphs that cannot fit contiguous Word row/column prefixes
    return (0, 0) and remain reversible in the dedicated association map.
    """

    coverage: dict[tuple[int, int], ClipboardTableCell] = {}
    for cell in table.cells:
        for row in range(cell.row, cell.row + cell.row_span):
            for column in range(cell.column, cell.column + cell.column_span):
                coverage[(row, column)] = cell

    row_prefix = 0
    for row in range(table.row_count):
        if all(coverage[(row, column)].header for column in range(table.column_count)):
            row_prefix += 1
        else:
            break
    column_prefix = 0
    for column in range(table.column_count):
        if all(coverage[(row, column)].header for row in range(table.row_count)):
            column_prefix += 1
        else:
            break

    candidates = {(0, 0), (row_prefix, 0), (0, column_prefix), (row_prefix, column_prefix)}
    valid: list[tuple[int, int, int]] = []
    for header_rows, header_columns in candidates:
        if not _native_header_candidate(table, header_rows, header_columns):
            continue
        area = header_rows * table.column_count + header_columns * table.row_count - header_rows * header_columns
        valid.append((area, header_columns, header_rows))
    if not valid:
        return 0, 0
    _area, header_columns, header_rows = min(valid, key=lambda item: (item[0], item[1], -item[2]))
    return header_rows, header_columns


def clipboard_inline_text(inline: ClipboardInline) -> str:
    if isinstance(inline, ClipboardText):
        return inline.value
    if isinstance(inline, ClipboardHardBreak):
        return "\n"
    return inline.alt


def clipboard_paragraph_text(paragraph: ClipboardParagraph) -> str:
    return "".join(clipboard_inline_text(item) for item in paragraph.inlines)


def clipboard_cell_text(cell: ClipboardTableCell) -> str:
    return "\n".join(clipboard_paragraph_text(block) for block in cell.blocks if isinstance(block, ClipboardParagraph))


def iter_clipboard_blocks(document: ClipboardDocument):
    def walk(blocks: tuple[ClipboardBlock, ...]):
        for block in blocks:
            yield block
            if isinstance(block, ClipboardTable):
                for cell in block.cells:
                    yield from walk(cell.blocks)

    yield from walk(document.blocks)


def iter_clipboard_inlines(document: ClipboardDocument):
    for block in iter_clipboard_blocks(document):
        if isinstance(block, ClipboardParagraph):
            yield from block.inlines


def validate_clipboard_resource_refs(
    document: ClipboardDocument,
    refs: Sequence[FileRef],
) -> None:
    """Validate one frozen typed-resource set against the document declaration."""

    from docwen_core.models.file_ref import (
        MANAGED_INPUT_SHA256_METADATA_KEY,
        MANAGED_INPUT_SIZE_BYTES_METADATA_KEY,
        MANAGED_RESOURCE_ID_METADATA_KEY,
    )

    declarations = {item.resource_id: item for item in document.resources}
    if len(refs) != len(declarations):
        raise ClipboardDocumentError(
            "clipboard.resource_set_mismatch",
            "Structured clipboard resource count does not match the document declaration.",
        )

    seen: set[str] = set()
    for ref in refs:
        resource_id = ref.metadata.get(MANAGED_RESOURCE_ID_METADATA_KEY)
        if (
            ref.input_role != "linked_resource"
            or ref.input_kind != "resource"
            or not isinstance(resource_id, str)
            or not resource_id
            or resource_id in seen
        ):
            raise ClipboardDocumentError(
                "clipboard.resource_ref_invalid",
                "Structured clipboard linked resource identity is invalid.",
            )
        seen.add(resource_id)
        declaration = declarations.get(resource_id)
        if declaration is None:
            raise ClipboardDocumentError(
                "clipboard.resource_set_mismatch",
                "Structured clipboard linked resource is not declared by the document.",
            )
        managed_size = ref.metadata.get(MANAGED_INPUT_SIZE_BYTES_METADATA_KEY)
        managed_sha = ref.metadata.get(MANAGED_INPUT_SHA256_METADATA_KEY)
        if (
            ref.logical_path != declaration.logical_path
            or ref.media_type != declaration.media_type
            or ref.size_bytes != declaration.size_bytes
            or type(managed_size) is not int
            or managed_size != declaration.size_bytes
            or not isinstance(managed_sha, str)
            or managed_sha != declaration.sha256
        ):
            raise ClipboardDocumentError(
                "clipboard.resource_ref_mismatch",
                "Structured clipboard linked resource metadata does not match its declaration.",
            )
    if seen != set(declarations):
        raise ClipboardDocumentError(
            "clipboard.resource_set_mismatch",
            "Structured clipboard linked resources do not match the document declaration.",
        )


def clipboard_document_preview(document: ClipboardDocument, *, max_chars: int = 240) -> str:
    parts: list[str] = []
    for block in document.blocks:
        if isinstance(block, ClipboardParagraph):
            parts.append(clipboard_paragraph_text(block))
        else:
            parts.append(f"[Table {block.row_count}×{block.column_count}]")
        if sum(len(part) for part in parts) >= max_chars:
            break
    return "\n".join(parts)[:max_chars]


__all__ = [
    "CLIPBOARD_DOCUMENT_FORMAT",
    "CLIPBOARD_DOCUMENT_MEDIA_TYPE",
    "CLIPBOARD_DOCUMENT_SCHEMA",
    "MAX_CLIPBOARD_DOCUMENT_BYTES",
    "MAX_CLIPBOARD_IMAGE_PIXELS",
    "MAX_CLIPBOARD_IMAGE_PIXELS_PER_RESOURCE",
    "MAX_CLIPBOARD_IMAGE_SIDE",
    "ClipboardBlock",
    "ClipboardDocument",
    "ClipboardDocumentError",
    "ClipboardHardBreak",
    "ClipboardImageRef",
    "ClipboardInline",
    "ClipboardParagraph",
    "ClipboardResource",
    "ClipboardTable",
    "ClipboardTableCell",
    "ClipboardText",
    "clipboard_cell_text",
    "clipboard_document_preview",
    "clipboard_document_to_bytes",
    "clipboard_document_to_dict",
    "clipboard_inline_text",
    "clipboard_paragraph_text",
    "clipboard_table_header_shape",
    "iter_clipboard_blocks",
    "iter_clipboard_inlines",
    "load_clipboard_document_bytes",
    "validate_clipboard_resource_refs",
]

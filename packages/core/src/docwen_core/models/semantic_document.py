"""Provider-neutral semantic document model.

These immutable objects are an in-process API.  They intentionally provide no
JSON projection, schema identity, or wire representation.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Literal
from urllib.parse import urlsplit

type SemanticLevel = Literal["info", "warning", "error"]
type SemanticCaptionKind = Literal["figure", "table", "equation", "listing"]
type SemanticTableCellRole = Literal["data", "column_header", "row_header", "corner_header"]
type SemanticRepeatHeader = Literal["inherit", "always", "never"]

# The longest current reversible DOCX name prefix is seven characters.  With
# unpadded base32, 20 ASCII bytes need 32 characters (39 total), while 21 need
# 34 (41 total) and would exceed Word's 40-character bookmark-name limit.
_PORTABLE_ENCODED_NAME_MAX_LENGTH = 40
_PORTABLE_ENCODED_PREFIX_MAX_LENGTH = 7
PORTABLE_TARGET_ID_MAX_LENGTH = (_PORTABLE_ENCODED_NAME_MAX_LENGTH - _PORTABLE_ENCODED_PREFIX_MAX_LENGTH) * 5 // 8
_TARGET_ID_RE = re.compile(r"^[A-Za-z][A-Za-z0-9._-]*$")
_TABLE_ROLES = frozenset({"data", "column_header", "row_header", "corner_header"})
_REPEAT_HEADER_VALUES = frozenset({"inherit", "always", "never"})


@dataclass(frozen=True, slots=True)
class SemanticDiagnostic:
    """One stable diagnostic from neutral validation or DOCX import."""

    level: SemanticLevel
    code: str
    message: str
    location: str = ""


@dataclass(frozen=True, slots=True)
class SemanticText:
    """Plain text inside a neutral paragraph."""

    value: str


@dataclass(frozen=True, slots=True)
class SemanticReference:
    """A target-bound cross-reference with a visible cached result."""

    target_id: str
    cached_result: str


@dataclass(frozen=True, slots=True)
class SemanticCitationItem:
    """One provider-neutral bibliography item referenced by a citation."""

    item_id: str


@dataclass(frozen=True, slots=True)
class SemanticCitationCluster:
    """An ordered citation cluster with a caller-supplied visible result."""

    cluster_id: str
    items: tuple[SemanticCitationItem, ...]
    cached_result: str


type SemanticInline = SemanticText | SemanticReference | SemanticCitationCluster


@dataclass(frozen=True, slots=True)
class SemanticParagraph:
    """A paragraph whose inline semantics do not depend on source grammar."""

    inlines: tuple[SemanticInline, ...]


@dataclass(frozen=True, slots=True)
class SemanticCaption:
    """A caption explicitly bound to its owning semantic object."""

    kind: SemanticCaptionKind
    target_id: str | None
    cached_number: str
    label: str
    content: str


@dataclass(frozen=True, slots=True)
class SemanticTableCell:
    """One anchor cell in a merge-aware neutral table."""

    row: int
    column: int
    text: str
    role: SemanticTableCellRole = "data"
    row_span: int = 1
    column_span: int = 1


@dataclass(frozen=True, slots=True)
class SemanticTable:
    """An anchor-only table with explicit roles and merge geometry."""

    row_count: int
    column_count: int
    cells: tuple[SemanticTableCell, ...]
    repeat_header: SemanticRepeatHeader = "inherit"
    caption: SemanticCaption | None = None


@dataclass(frozen=True, slots=True)
class SemanticBibliographyRun:
    """One already-presented rich run inside a bibliography entry."""

    text: str
    bold: bool = False
    italic: bool = False
    href: str | None = None


@dataclass(frozen=True, slots=True)
class SemanticBibliographyEntry:
    """One already-formatted bibliography entry with a reversible item ID."""

    item_id: str
    runs: tuple[SemanticBibliographyRun, ...]


@dataclass(frozen=True, slots=True)
class SemanticBibliographyFragment:
    """Already-formatted bibliography entries; the heading stays independent."""

    entries: tuple[SemanticBibliographyEntry, ...]


type SemanticBlock = SemanticParagraph | SemanticTable


@dataclass(frozen=True, slots=True)
class SemanticDocument:
    """Provider-neutral semantic blocks plus an optional bibliography fragment."""

    blocks: tuple[SemanticBlock, ...]
    bibliography: SemanticBibliographyFragment | None = None


@dataclass(frozen=True, slots=True)
class SemanticImportResult:
    """Neutral document recovered from DOCX plus fail-closed diagnostics."""

    document: SemanticDocument
    diagnostics: tuple[SemanticDiagnostic, ...] = ()

    @property
    def has_errors(self) -> bool:
        return any(item.level == "error" for item in self.diagnostics)


class SemanticDocumentValidationError(ValueError):
    """Raised when rendering is attempted with invalid neutral semantics."""

    def __init__(self, diagnostics: tuple[SemanticDiagnostic, ...]) -> None:
        self.diagnostics = diagnostics
        super().__init__("neutral semantic document is invalid")


def validate_semantic_document(document: SemanticDocument) -> tuple[SemanticDiagnostic, ...]:
    """Validate object binding, references, and explicit table geometry."""

    diagnostics: list[SemanticDiagnostic] = []
    targets: set[str] = set()
    references: list[tuple[SemanticReference, str]] = []
    citation_cluster_ids: set[str] = set()

    for block_index, block in enumerate(document.blocks):
        location = f"blocks[{block_index}]"
        if isinstance(block, SemanticParagraph):
            for inline_index, inline in enumerate(block.inlines):
                inline_location = f"{location}.inlines[{inline_index}]"
                if isinstance(inline, SemanticReference):
                    references.append((inline, inline_location))
                elif isinstance(inline, SemanticCitationCluster):
                    diagnostics.extend(
                        _validate_citation_cluster(
                            inline,
                            location=inline_location,
                            cluster_ids=citation_cluster_ids,
                        )
                    )
            continue

        if block.caption is not None:
            caption = block.caption
            caption_location = f"{location}.caption"
            if caption.kind != "table":
                diagnostics.append(
                    SemanticDiagnostic(
                        "error",
                        "semantic.caption.object_kind_mismatch",
                        "A SemanticTable caption must use kind 'table'.",
                        caption_location,
                    )
                )
            if caption.target_id is not None:
                if not is_portable_semantic_id(caption.target_id):
                    diagnostics.append(
                        SemanticDiagnostic(
                            "error",
                            "semantic.caption.target_id_invalid",
                            "Caption target_id must be an ASCII portable identifier of at most "
                            f"{PORTABLE_TARGET_ID_MAX_LENGTH} characters when provided.",
                            caption_location,
                        )
                    )
                elif caption.target_id in targets:
                    diagnostics.append(
                        SemanticDiagnostic(
                            "error",
                            "semantic.target.duplicate",
                            f"Target {caption.target_id} occurs more than once.",
                            caption_location,
                        )
                    )
                else:
                    targets.add(caption.target_id)
            if not caption.cached_number:
                diagnostics.append(
                    SemanticDiagnostic(
                        "error",
                        "semantic.caption.cached_number_empty",
                        "Caption cached_number must not be empty.",
                        caption_location,
                    )
                )
            if not caption.label.strip():
                diagnostics.append(
                    SemanticDiagnostic(
                        "error",
                        "semantic.caption.label_empty",
                        "Caption label must not be empty.",
                        caption_location,
                    )
                )

        diagnostics.extend(_validate_table(block, location=location))

    for reference, location in references:
        if not is_portable_semantic_id(reference.target_id):
            diagnostics.append(
                SemanticDiagnostic(
                    "error",
                    "semantic.reference.target_id_invalid",
                    "Reference target_id must be an ASCII portable identifier of at most "
                    f"{PORTABLE_TARGET_ID_MAX_LENGTH} characters.",
                    location,
                )
            )
        elif reference.target_id not in targets:
            diagnostics.append(
                SemanticDiagnostic(
                    "error",
                    "semantic.reference.target_missing",
                    f"Reference target {reference.target_id} does not exist.",
                    location,
                )
            )
        if not reference.cached_result:
            diagnostics.append(
                SemanticDiagnostic(
                    "error",
                    "semantic.reference.cached_result_empty",
                    "Reference cached_result must not be empty.",
                    location,
                )
            )

    if document.bibliography is not None:
        bibliography_item_ids: set[str] = set()
        for entry_index, entry in enumerate(document.bibliography.entries):
            location = f"bibliography.entries[{entry_index}]"
            if not is_portable_semantic_id(entry.item_id):
                diagnostics.append(
                    SemanticDiagnostic(
                        "error",
                        "semantic.bibliography.item_id_invalid",
                        "Bibliography item_id must be an ASCII portable identifier of at most "
                        f"{PORTABLE_TARGET_ID_MAX_LENGTH} characters.",
                        location,
                    )
                )
            elif entry.item_id in bibliography_item_ids:
                diagnostics.append(
                    SemanticDiagnostic(
                        "error",
                        "semantic.bibliography.item_id_duplicate",
                        f"Bibliography item {entry.item_id} occurs more than once.",
                        location,
                    )
                )
            else:
                bibliography_item_ids.add(entry.item_id)
            if not entry.runs:
                diagnostics.append(
                    SemanticDiagnostic(
                        "error",
                        "semantic.bibliography.entry_empty",
                        "Bibliography entries must contain at least one already-formatted run.",
                        location,
                    )
                )
                continue
            for run_index, run in enumerate(entry.runs):
                run_location = f"{location}.runs[{run_index}]"
                if type(run.text) is not str or not run.text or not _is_xml_10_text(run.text):
                    diagnostics.append(
                        SemanticDiagnostic(
                            "error",
                            "semantic.bibliography.run_text_empty",
                            "Bibliography run text must be a non-empty XML 1.0-compatible string.",
                            run_location,
                        )
                    )
                if type(run.bold) is not bool:
                    diagnostics.append(
                        SemanticDiagnostic(
                            "error",
                            "semantic.bibliography.run_bold_invalid",
                            "Bibliography run bold must be a boolean.",
                            run_location,
                        )
                    )
                if type(run.italic) is not bool:
                    diagnostics.append(
                        SemanticDiagnostic(
                            "error",
                            "semantic.bibliography.run_italic_invalid",
                            "Bibliography run italic must be a boolean.",
                            run_location,
                        )
                    )
                if run.href is not None and not _is_absolute_http_url(run.href):
                    diagnostics.append(
                        SemanticDiagnostic(
                            "error",
                            "semantic.bibliography.run_href_invalid",
                            "Bibliography run href must be an absolute HTTP or HTTPS URL.",
                            run_location,
                        )
                    )
            if not any(type(run.text) is str and run.text.strip() for run in entry.runs):
                diagnostics.append(
                    SemanticDiagnostic(
                        "error",
                        "semantic.bibliography.entry_empty",
                        "Bibliography entries must contain visible already-formatted text.",
                        location,
                    )
                )

    return tuple(diagnostics)


def _is_absolute_http_url(value: object) -> bool:
    if not isinstance(value, str) or not value:
        return False
    try:
        parsed = urlsplit(value)
        return parsed.scheme.lower() in {"http", "https"} and parsed.hostname is not None
    except ValueError:
        return False


def _is_xml_10_text(value: str) -> bool:
    return all(
        character in {"\t", "\n", "\r"}
        or "\u0020" <= character <= "\ud7ff"
        or "\ue000" <= character <= "\ufffd"
        or "\U00010000" <= character <= "\U0010ffff"
        for character in value
    )


def _validate_citation_cluster(
    cluster: SemanticCitationCluster,
    *,
    location: str,
    cluster_ids: set[str],
) -> list[SemanticDiagnostic]:
    diagnostics: list[SemanticDiagnostic] = []
    if not is_portable_semantic_id(cluster.cluster_id):
        diagnostics.append(
            SemanticDiagnostic(
                "error",
                "semantic.citation.cluster_id_invalid",
                "Citation cluster_id must be an ASCII portable identifier of at most "
                f"{PORTABLE_TARGET_ID_MAX_LENGTH} characters.",
                location,
            )
        )
    elif cluster.cluster_id in cluster_ids:
        diagnostics.append(
            SemanticDiagnostic(
                "error",
                "semantic.citation.cluster_id_duplicate",
                f"Citation cluster {cluster.cluster_id} occurs more than once.",
                location,
            )
        )
    else:
        cluster_ids.add(cluster.cluster_id)
    if not cluster.items:
        diagnostics.append(
            SemanticDiagnostic(
                "error",
                "semantic.citation.items_empty",
                "Citation clusters must contain at least one ordered item.",
                location,
            )
        )
    for item_index, item in enumerate(cluster.items):
        if not is_portable_semantic_id(item.item_id):
            diagnostics.append(
                SemanticDiagnostic(
                    "error",
                    "semantic.citation.item_id_invalid",
                    "Citation item_id must be an ASCII portable identifier of at most "
                    f"{PORTABLE_TARGET_ID_MAX_LENGTH} characters.",
                    f"{location}.items[{item_index}]",
                )
            )
    if not cluster.cached_result.strip():
        diagnostics.append(
            SemanticDiagnostic(
                "error",
                "semantic.citation.cached_result_empty",
                "Citation cached_result must not be empty.",
                location,
            )
        )
    return diagnostics


def derive_table_header_shape(table: SemanticTable) -> tuple[int, int]:
    """Return validated contiguous header row and column counts."""

    role_grid, diagnostics = _table_role_grid(table, location="table")
    if diagnostics or not role_grid:
        raise SemanticDocumentValidationError(tuple(diagnostics))
    return _derive_header_shape(role_grid)


def _validate_table(table: SemanticTable, *, location: str) -> list[SemanticDiagnostic]:
    role_grid, diagnostics = _table_role_grid(table, location=location)
    if diagnostics or not role_grid:
        return diagnostics

    header_rows, header_columns = _derive_header_shape(role_grid)
    for row in range(table.row_count):
        for column in range(table.column_count):
            expected = _role_for_position(row, column, header_rows=header_rows, header_columns=header_columns)
            actual = role_grid[row][column]
            if actual != expected:
                diagnostics.append(
                    SemanticDiagnostic(
                        "error",
                        "semantic.table.role_boundary_invalid",
                        "Table roles must form contiguous header row and column prefixes; spans may not cross them.",
                        f"{location}.cell[{row},{column}]",
                    )
                )
                return diagnostics

    if table.repeat_header not in _REPEAT_HEADER_VALUES:
        diagnostics.append(
            SemanticDiagnostic(
                "error",
                "semantic.table.repeat_header_invalid",
                "repeat_header must be inherit, always, or never.",
                location,
            )
        )
    elif table.repeat_header != "inherit" and header_rows == 0:
        diagnostics.append(
            SemanticDiagnostic(
                "error",
                "semantic.table.repeat_header_without_header",
                "Explicit repeat_header requires at least one header row.",
                location,
            )
        )
    return diagnostics


def _table_role_grid(
    table: SemanticTable,
    *,
    location: str,
) -> tuple[list[list[str]], list[SemanticDiagnostic]]:
    diagnostics: list[SemanticDiagnostic] = []
    if table.row_count <= 0 or table.column_count <= 0:
        diagnostics.append(
            SemanticDiagnostic(
                "error",
                "semantic.table.dimensions_invalid",
                "Table dimensions must both be positive.",
                location,
            )
        )
        return [], diagnostics

    coverage: dict[tuple[int, int], SemanticTableCell] = {}
    for cell_index, cell in enumerate(table.cells):
        cell_location = f"{location}.cells[{cell_index}]"
        if cell.role not in _TABLE_ROLES:
            diagnostics.append(
                SemanticDiagnostic(
                    "error",
                    "semantic.table.cell_role_invalid",
                    f"Unsupported table cell role: {cell.role}",
                    cell_location,
                )
            )
            continue
        if cell.row_span <= 0 or cell.column_span <= 0:
            diagnostics.append(
                SemanticDiagnostic(
                    "error",
                    "semantic.table.span_invalid",
                    "Table cell spans must both be positive.",
                    cell_location,
                )
            )
            continue
        if (
            cell.row < 0
            or cell.column < 0
            or cell.row + cell.row_span > table.row_count
            or cell.column + cell.column_span > table.column_count
        ):
            diagnostics.append(
                SemanticDiagnostic(
                    "error",
                    "semantic.table.cell_out_of_bounds",
                    "Table cell geometry exceeds the declared grid.",
                    cell_location,
                )
            )
            continue
        for row in range(cell.row, cell.row + cell.row_span):
            for column in range(cell.column, cell.column + cell.column_span):
                position = (row, column)
                if position in coverage:
                    diagnostics.append(
                        SemanticDiagnostic(
                            "error",
                            "semantic.table.cell_overlap",
                            "Anchor-only table cells must not overlap.",
                            f"{location}.cell[{row},{column}]",
                        )
                    )
                    continue
                coverage[position] = cell

    expected_positions = {(row, column) for row in range(table.row_count) for column in range(table.column_count)}
    missing = sorted(expected_positions - set(coverage))
    if missing:
        row, column = missing[0]
        diagnostics.append(
            SemanticDiagnostic(
                "error",
                "semantic.table.grid_incomplete",
                "Anchor-only table cells must cover the complete declared grid.",
                f"{location}.cell[{row},{column}]",
            )
        )
    if diagnostics:
        return [], diagnostics

    role_grid = [
        [coverage[(row, column)].role for column in range(table.column_count)] for row in range(table.row_count)
    ]
    return role_grid, diagnostics


def _derive_header_shape(role_grid: list[list[str]]) -> tuple[int, int]:
    header_rows = 0
    for row in role_grid:
        if all(role in {"corner_header", "column_header"} for role in row):
            header_rows += 1
        else:
            break

    header_columns = 0
    column_count = len(role_grid[0]) if role_grid else 0
    for column in range(column_count):
        if all(row[column] in {"corner_header", "row_header"} for row in role_grid):
            header_columns += 1
        else:
            break
    return header_rows, header_columns


def _role_for_position(row: int, column: int, *, header_rows: int, header_columns: int) -> str:
    if row < header_rows and column < header_columns:
        return "corner_header"
    if row < header_rows:
        return "column_header"
    if column < header_columns:
        return "row_header"
    return "data"


def is_portable_semantic_id(value: str) -> bool:
    """Return whether an in-process semantic ID fits all reversible DOCX names."""

    return len(value) <= PORTABLE_TARGET_ID_MAX_LENGTH and _TARGET_ID_RE.fullmatch(value) is not None


__all__ = [
    "PORTABLE_TARGET_ID_MAX_LENGTH",
    "SemanticBibliographyEntry",
    "SemanticBibliographyFragment",
    "SemanticBlock",
    "SemanticCaption",
    "SemanticCaptionKind",
    "SemanticCitationCluster",
    "SemanticCitationItem",
    "SemanticDiagnostic",
    "SemanticDocument",
    "SemanticDocumentValidationError",
    "SemanticImportResult",
    "SemanticInline",
    "SemanticLevel",
    "SemanticParagraph",
    "SemanticReference",
    "SemanticRepeatHeader",
    "SemanticTable",
    "SemanticTableCell",
    "SemanticTableCellRole",
    "SemanticText",
    "derive_table_header_shape",
    "is_portable_semantic_id",
    "validate_semantic_document",
]

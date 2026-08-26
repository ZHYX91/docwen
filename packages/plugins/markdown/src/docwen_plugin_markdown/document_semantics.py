"""Structured Markdown document semantics for the first DOCX interop slice.

The internal annotations in this module remain DocWen-owned.  The optional
``oracle_projection`` is a test/acceptance projection, not a wire DTO or a
required runtime model.
"""

from __future__ import annotations

import copy
import re
from dataclasses import dataclass
from typing import Any

_CAPTION_RE = re.compile(
    r"^(Figure|Table|Equation|Listing|):[ \t]+(.+?)(?:[ \t]+\{#([A-Za-z][A-Za-z0-9._-]*)\})?[ \t]*$",
    re.IGNORECASE,
)
_TABLE_ATTRIBUTE_RE = re.compile(r"^\{([^{}]+)\}$")
_SEMANTIC_TOKEN_RE = re.compile(
    r"\[@[A-Za-z0-9._:-]+(?:\s*;\s*@[A-Za-z0-9._:-]+)*\]"
    r"|(?<![A-Za-z0-9_@])@(?:fig|tbl|eq|lst)-[A-Za-z0-9](?:[A-Za-z0-9._-]*[A-Za-z0-9_-])?"
)
_TARGET_KIND_PREFIX = {
    "fig": "figure",
    "tbl": "table",
    "eq": "equation",
    "lst": "listing",
}
_KIND_PREFIX = {value: key for key, value in _TARGET_KIND_PREFIX.items()}
_DIAGNOSTIC_ORDER = {
    "interop.caption.kind_mismatch": 10,
    "interop.caption.binding_broken": 20,
    "interop.target.duplicate": 30,
    "interop.reference.missing": 40,
    "interop.reference.kind_mismatch": 50,
    "interop.table.merge_non_rectangular": 60,
    "interop.table.merge_role_crossing": 70,
    "interop.table.attribute_invalid": 80,
    "interop.citation.processor_unavailable": 90,
}


@dataclass(frozen=True, slots=True)
class SemanticDiagnostic:
    """Stable, provider-neutral diagnostic emitted by DocWen semantics."""

    level: str
    code: str
    message: str


@dataclass(slots=True)
class DocumentSemanticsAnalysis:
    """Transformed AST plus a brand-neutral acceptance projection."""

    ast: list[dict[str, Any]]
    oracle_projection: dict[str, Any]
    diagnostics: list[SemanticDiagnostic]

    @property
    def has_errors(self) -> bool:
        return any(item.level == "error" for item in self.diagnostics)


def analyze_document_semantics(
    ast_nodes: list[dict[str, Any]],
    *,
    current_v3: bool = False,
) -> DocumentSemanticsAnalysis:
    """Recognize legacy semantics or only shared table metadata for v3.

    ``current_v3`` is used only after the raw-source v3 adapter has already
    produced typed nodes. In that mode this compatibility module must not
    reinterpret ``Listing:``, ``{#id}``, or bare ``@id`` source as the
    superseded v1 grammar; it still computes the independent table grid used
    by the renderer.
    """

    ast = copy.deepcopy(ast_nodes)
    diagnostics: list[SemanticDiagnostic] = []
    removed: set[int] = set()
    target_occurrences: dict[str, list[dict[str, Any]]] = {}
    counters = dict.fromkeys(_KIND_PREFIX, 0)

    for index, node in enumerate(ast):
        if current_v3:
            continue
        caption = _parse_caption(node)
        if caption is None:
            continue
        next_index = _next_nonblank(ast, index + 1)
        if next_index is None:
            diagnostics.append(_binding_broken())
            continue
        object_node = ast[next_index]
        object_kind = _captionable_kind(object_node)
        if object_kind is None:
            diagnostics.append(_binding_broken())
            continue

        caption_kind = caption["kind"] or object_kind
        if caption_kind != object_kind:
            diagnostics.append(
                SemanticDiagnostic(
                    level="error",
                    code="interop.caption.kind_mismatch",
                    message=f"{caption_kind.title()} caption is followed by a {object_kind}.",
                )
            )
            continue

        target_id = caption.get("target_id")
        if target_id is not None and _kind_from_target_id(target_id) != object_kind:
            diagnostics.append(
                SemanticDiagnostic(
                    level="error",
                    code="interop.caption.kind_mismatch",
                    message=f"{caption_kind.title()} caption target ID has the wrong kind prefix.",
                )
            )
            continue

        semantic_caption = {
            "kind": object_kind,
            "content": caption["content"],
            "source_form": caption["source_form"],
            "target_id": target_id,
        }
        if object_kind == "figure":
            ast[next_index] = {
                "type": "semantic_figure",
                "object": object_node,
                "_document_semantics_caption": semantic_caption,
            }
            bound_node = ast[next_index]
        else:
            object_node["_document_semantics_caption"] = semantic_caption
            bound_node = object_node

        counters[object_kind] += 1
        bound_node["_document_semantics_number"] = counters[object_kind]

        removed.add(index)
        removed.update(
            item_index for item_index in range(index + 1, next_index) if ast[item_index].get("type") == "blank_line"
        )
        if target_id is not None:
            target_occurrences.setdefault(target_id, []).append(
                {
                    "kind": object_kind,
                    "node": bound_node,
                    "index": next_index,
                    "number": counters[object_kind],
                }
            )

    for index, node in enumerate(ast):
        if node.get("type") != "table":
            continue
        attributes = None
        attribute_index = index + 1
        if attribute_index < len(ast) and ast[attribute_index].get("type") == "paragraph":
            attributes = _parse_table_attributes(ast[attribute_index])
            if attributes is not None:
                removed.add(attribute_index)
        elif (
            attribute_index + 1 < len(ast)
            and ast[attribute_index].get("type") == "blank_line"
            and _looks_like_table_attribute(ast[attribute_index + 1])
        ):
            diagnostics.append(_attribute_invalid("Table attributes must immediately follow the table."))

        metadata, table_diagnostics = _analyze_table(node, attributes)
        node["_document_semantics_table"] = metadata
        diagnostics.extend(table_diagnostics)

    targets: dict[str, dict[str, Any]] = {}
    for target_id, occurrences in target_occurrences.items():
        if len(occurrences) > 1:
            diagnostics.append(
                SemanticDiagnostic(
                    level="error",
                    code="interop.target.duplicate",
                    message=f"Target {target_id} occurs more than once in the document.",
                )
            )
        first = occurrences[0]
        targets[target_id] = {
            "kind": first["kind"],
            "duplicate": len(occurrences) > 1,
            "number": first["number"],
        }

    citation_found = False
    for node in ast:
        if current_v3:
            continue
        if node.get("type") != "paragraph":
            continue
        children, inline_diagnostics, has_citation = _annotate_inline_children(node.get("children", []), targets)
        node["children"] = children
        diagnostics.extend(inline_diagnostics)
        citation_found = citation_found or has_citation

    if citation_found:
        diagnostics.append(
            SemanticDiagnostic(
                level="warning",
                code="interop.citation.processor_unavailable",
                message="Citation keys remain literal because DocWen does not run a CSL citation processor.",
            )
        )

    transformed = [node for index, node in enumerate(ast) if index not in removed]
    projection = {
        "schema": "docwen.document_semantics.v1",
        "blocks": _project_blocks(transformed),
    }
    diagnostics = _deduplicate_and_sort_diagnostics(diagnostics)
    return DocumentSemanticsAnalysis(
        ast=transformed,
        oracle_projection=projection,
        diagnostics=diagnostics,
    )


def _parse_caption(node: dict[str, Any]) -> dict[str, Any] | None:
    if node.get("type") != "paragraph":
        return None
    raw = _plain_inline_text(node.get("children", [])).strip()
    match = _CAPTION_RE.fullmatch(raw)
    if match is None:
        return None
    declaration, content, target_id = match.groups()
    if not declaration:
        kind = None
        source_form = "shorthand"
    else:
        kind = declaration.lower()
        source_form = "canonical"
    return {
        "kind": kind,
        "content": content.strip(),
        "target_id": target_id,
        "source_form": source_form,
    }


def _captionable_kind(node: dict[str, Any]) -> str | None:
    node_type = node.get("type")
    if node_type == "table":
        return "table"
    if node_type in {"block_math", "block_latex"}:
        return "equation"
    if node_type == "block_code":
        return "listing"
    if node_type == "paragraph":
        meaningful = [
            child
            for child in node.get("children", [])
            if child.get("type") not in {"softbreak", "linebreak"}
            and (child.get("type") != "text" or _plain_inline_text([child]).strip())
        ]
        if len(meaningful) == 1 and meaningful[0].get("type") == "image":
            return "figure"
    return None


def _binding_broken() -> SemanticDiagnostic:
    return SemanticDiagnostic(
        level="warning",
        code="interop.caption.binding_broken",
        message="An intervening paragraph breaks shorthand caption binding.",
    )


def _next_nonblank(ast: list[dict[str, Any]], start: int) -> int | None:
    for index in range(start, len(ast)):
        if ast[index].get("type") != "blank_line":
            return index
    return None


def _looks_like_table_attribute(node: dict[str, Any]) -> bool:
    if node.get("type") != "paragraph":
        return False
    raw = _plain_inline_text(node.get("children", [])).strip()
    return raw.startswith("{") and any(key in raw for key in ("header-rows", "header-cols", "repeat-header"))


def _parse_table_attributes(node: dict[str, Any]) -> dict[str, str] | None:
    raw = _plain_inline_text(node.get("children", [])).strip()
    match = _TABLE_ATTRIBUTE_RE.fullmatch(raw)
    if match is None or not any(key in match.group(1) for key in ("header-rows", "header-cols", "repeat-header")):
        return None
    attributes: dict[str, str] = {}
    for part in match.group(1).split():
        if "=" not in part:
            attributes["__invalid__"] = part
            continue
        key, value = part.split("=", 1)
        if key in attributes:
            attributes["__duplicate__"] = key
        attributes[key] = value
    return attributes


def _analyze_table(
    node: dict[str, Any],
    attributes: dict[str, str] | None,
) -> tuple[dict[str, Any], list[SemanticDiagnostic]]:
    diagnostics: list[SemanticDiagnostic] = []
    rows = _table_rows(node)
    row_count = len(rows)
    column_count = max((len(row) for row in rows), default=0)
    attributes = attributes or {}

    attribute_error = bool(set(attributes) - {"header-rows", "header-cols", "repeat-header"})
    try:
        header_rows = int(attributes.get("header-rows", "1"))
        header_columns = int(attributes.get("header-cols", "0"))
    except ValueError:
        header_rows = 1
        header_columns = 0
        attribute_error = True
    repeat_value = attributes.get("repeat-header")
    if repeat_value not in {None, "true", "false"}:
        attribute_error = True
    repeat_header = {None: "inherit", "true": "always", "false": "never"}.get(repeat_value, "inherit")

    if header_rows < 0 or header_rows > row_count:
        diagnostics.append(_attribute_invalid("header-rows exceeds the table row count."))
        attribute_error = False
    if header_columns < 0 or header_columns > column_count:
        diagnostics.append(_attribute_invalid("header-cols exceeds the table column count."))
        attribute_error = False
    if attribute_error:
        diagnostics.append(_attribute_invalid("Table attributes are invalid."))

    effective_header_rows = max(0, min(header_rows, row_count))
    effective_header_columns = max(0, min(header_columns, column_count))
    anchors, merge_non_rectangular, merge_role_crossing = _table_anchor_cells(
        rows,
        row_count=row_count,
        column_count=column_count,
        header_rows=effective_header_rows,
        header_columns=effective_header_columns,
    )
    if merge_non_rectangular:
        diagnostics.append(
            SemanticDiagnostic(
                level="error",
                code="interop.table.merge_non_rectangular",
                message="Merge markers do not cover a complete rectangle.",
            )
        )
    if merge_role_crossing:
        diagnostics.append(
            SemanticDiagnostic(
                level="error",
                code="interop.table.merge_role_crossing",
                message="Merge markers cross table header/data role boundaries.",
            )
        )

    return (
        {
            "row_count": row_count,
            "column_count": column_count,
            "header_rows": header_rows,
            "header_columns": header_columns,
            "repeat_header": repeat_header,
            "anchors": anchors,
        },
        diagnostics,
    )


def _attribute_invalid(message: str) -> SemanticDiagnostic:
    return SemanticDiagnostic(
        level="error",
        code="interop.table.attribute_invalid",
        message=message,
    )


def _table_rows(node: dict[str, Any]) -> list[list[dict[str, Any]]]:
    rows: list[list[dict[str, Any]]] = []
    for child in node.get("children", []):
        child_type = child.get("type")
        if child_type not in {"table_head", "table_body"}:
            continue
        nested = child.get("children", [])
        if nested and nested[0].get("type") == "table_cell":
            rows.append(nested)
            continue
        rows.extend(row.get("children", []) for row in nested if row.get("type") == "table_row")
    return rows


def _table_anchor_cells(
    rows: list[list[dict[str, Any]]],
    *,
    row_count: int,
    column_count: int,
    header_rows: int,
    header_columns: int,
) -> tuple[list[dict[str, Any]], bool, bool]:
    markers: dict[tuple[int, int], str] = {}
    cell_nodes: dict[tuple[int, int], dict[str, Any]] = {}
    for row_index in range(row_count):
        row = rows[row_index]
        for column_index in range(column_count):
            cell = row[column_index] if column_index < len(row) else {"type": "table_cell", "children": []}
            cell_nodes[(row_index, column_index)] = cell
            raw = _plain_inline_text(cell.get("children", [])).strip()
            if raw in {"<", "^"}:
                markers[(row_index, column_index)] = raw

    resolved: dict[tuple[int, int], tuple[int, int] | None] = {}

    def resolve(position: tuple[int, int]) -> tuple[int, int] | None:
        if position in resolved:
            return resolved[position]
        marker = markers.get(position)
        if marker is None:
            resolved[position] = position
            return position
        row, column = position
        target = (row, column - 1) if marker == "<" else (row - 1, column)
        if target not in cell_nodes:
            resolved[position] = None
            return None
        resolved[position] = resolve(target)
        return resolved[position]

    groups: dict[tuple[int, int], set[tuple[int, int]]] = {}
    invalid = False
    for position in cell_nodes:
        anchor = resolve(position)
        if anchor is None:
            invalid = True
            continue
        groups.setdefault(anchor, set()).add(position)

    anchors: list[dict[str, Any]] = []
    role_crossing = False
    for anchor, covered in sorted(groups.items()):
        min_row = min(position[0] for position in covered)
        max_row = max(position[0] for position in covered)
        min_column = min(position[1] for position in covered)
        max_column = max(position[1] for position in covered)
        rectangle = {
            (row, column) for row in range(min_row, max_row + 1) for column in range(min_column, max_column + 1)
        }
        if covered != rectangle or anchor != (min_row, min_column):
            invalid = True
        roles = {
            _cell_role(row, column, header_rows=header_rows, header_columns=header_columns) for row, column in covered
        }
        if len(roles) != 1:
            role_crossing = True
        cell = cell_nodes[anchor]
        anchors.append(
            {
                "row": anchor[0],
                "column": anchor[1],
                "row_span": max_row - min_row + 1,
                "column_span": max_column - min_column + 1,
                "role": _cell_role(
                    anchor[0],
                    anchor[1],
                    header_rows=header_rows,
                    header_columns=header_columns,
                ),
                "children": cell.get("children", []),
            }
        )
    return anchors, invalid, role_crossing


def _cell_role(row: int, column: int, *, header_rows: int, header_columns: int) -> str:
    if row < header_rows and column < header_columns:
        return "corner_header"
    if row < header_rows:
        return "column_header"
    if column < header_columns:
        return "row_header"
    return "data"


def _annotate_inline_children(
    children: list[dict[str, Any]],
    targets: dict[str, dict[str, Any]],
) -> tuple[list[dict[str, Any]], list[SemanticDiagnostic], bool]:
    coalesced: list[dict[str, Any]] = []
    for child in children:
        if child.get("type") == "text" and coalesced and coalesced[-1].get("type") == "text":
            coalesced[-1]["raw"] = (coalesced[-1].get("raw", "") or "") + (child.get("raw", "") or "")
        else:
            coalesced.append(child)

    output: list[dict[str, Any]] = []
    diagnostics: list[SemanticDiagnostic] = []
    citation_found = False
    for child in coalesced:
        child_type = child.get("type")
        if child_type in {"strong", "emphasis", "strikethrough"}:
            nested, nested_diagnostics, nested_citation = _annotate_inline_children(child.get("children", []), targets)
            child["children"] = nested
            diagnostics.extend(nested_diagnostics)
            citation_found = citation_found or nested_citation
            output.append(child)
            continue
        if child_type != "text":
            output.append(child)
            continue

        raw = child.get("raw", "") or ""
        cursor = 0
        for match in _SEMANTIC_TOKEN_RE.finditer(raw):
            if match.start() > cursor:
                output.append({"type": "text", "raw": raw[cursor : match.start()]})
            token = match.group(0)
            if token.startswith("["):
                items = [part.strip()[1:] for part in token[1:-1].split(";")]
                output.append(
                    {
                        "type": "semantic_citation",
                        "raw": token,
                        "items": items,
                    }
                )
                citation_found = True
            else:
                target_id = token[1:]
                expected_kind = _kind_from_target_id(target_id) or "figure"
                target = targets.get(target_id)
                if target is None:
                    status = "missing"
                    target_kind = expected_kind
                    diagnostics.append(
                        SemanticDiagnostic(
                            level="error",
                            code="interop.reference.missing",
                            message=f"Target {target_id} does not exist.",
                        )
                    )
                    cached_result = "?"
                elif target["kind"] != expected_kind:
                    status = "kind_mismatch"
                    target_kind = target["kind"]
                    diagnostics.append(
                        SemanticDiagnostic(
                            level="error",
                            code="interop.reference.kind_mismatch",
                            message=f"Target {target_id} has a different document-object kind.",
                        )
                    )
                    cached_result = str(target["number"])
                else:
                    status = "duplicate" if target["duplicate"] else "resolved"
                    target_kind = target["kind"]
                    cached_result = str(target["number"])
                output.append(
                    {
                        "type": "semantic_cross_reference",
                        "target_id": target_id,
                        "target_kind": target_kind,
                        "status": status,
                        "display": "auto",
                        "cached_result": cached_result,
                    }
                )
            cursor = match.end()
        if cursor < len(raw):
            output.append({"type": "text", "raw": raw[cursor:]})
    return output, diagnostics, citation_found


def _kind_from_target_id(target_id: str) -> str | None:
    prefix, separator, _rest = target_id.partition("-")
    return _TARGET_KIND_PREFIX.get(prefix.lower()) if separator else None


def _project_blocks(ast: list[dict[str, Any]]) -> list[dict[str, Any]]:
    blocks: list[dict[str, Any]] = []
    for node in ast:
        node_type = node.get("type")
        if node_type == "semantic_figure":
            caption = node["_document_semantics_caption"]
            image = next(child for child in node["object"].get("children", []) if child.get("type") == "image")
            block: dict[str, Any] = {
                "type": "figure",
                "caption": {
                    "kind": "figure",
                    "content": [{"type": "text", "value": caption["content"]}],
                    "source_form": caption["source_form"],
                },
                "image": {
                    "source": image.get("attrs", {}).get("url", ""),
                    "alt": _plain_inline_text(image.get("children", [])),
                },
            }
            if caption.get("target_id"):
                block["target_id"] = caption["target_id"]
            blocks.append(block)
        elif node_type == "table" and node.get("_document_semantics_caption"):
            caption = node["_document_semantics_caption"]
            metadata = node["_document_semantics_table"]
            block = {
                "type": "table",
                "caption": {
                    "kind": "table",
                    "content": [{"type": "text", "value": caption["content"]}],
                    "source_form": caption["source_form"],
                },
                "row_count": metadata["row_count"],
                "column_count": metadata["column_count"],
                "header_rows": metadata["header_rows"],
                "header_columns": metadata["header_columns"],
                "repeat_header": metadata["repeat_header"],
                "cells": [
                    {
                        "row": cell["row"],
                        "column": cell["column"],
                        "row_span": cell["row_span"],
                        "column_span": cell["column_span"],
                        "role": cell["role"],
                        "content": _project_inline_content(cell["children"]),
                    }
                    for cell in metadata["anchors"]
                ],
            }
            if caption.get("target_id"):
                block["target_id"] = caption["target_id"]
            blocks.append(block)
        elif node_type in {"block_math", "block_latex", "block_code"} and node.get("_document_semantics_caption"):
            caption = node["_document_semantics_caption"]
            block = {
                "type": caption["kind"],
                "caption": {
                    "kind": caption["kind"],
                    "content": [{"type": "text", "value": caption["content"]}],
                    "source_form": caption["source_form"],
                },
            }
            if caption["kind"] == "equation":
                block["latex"] = str(node.get("raw", "") or node.get("text", "")).strip()
            else:
                block["code"] = str(node.get("raw", "") or node.get("text", "")).rstrip("\n")
                language = str(node.get("attrs", {}).get("info", "")).strip()
                if language:
                    block["language"] = language.split(maxsplit=1)[0]
            if caption.get("target_id"):
                block["target_id"] = caption["target_id"]
            blocks.append(block)
        elif node_type == "paragraph" and _plain_inline_text(node.get("children", [])).strip():
            blocks.append(
                {
                    "type": "paragraph",
                    "content": _project_inline_content(node.get("children", [])),
                }
            )
    return blocks


def _project_inline_content(children: list[dict[str, Any]]) -> list[dict[str, Any]]:
    projected: list[dict[str, Any]] = []
    for child in children:
        child_type = child.get("type")
        if child.get("schema") == "docwen.markdown_semantics.v3":
            # The v1 projection is a test oracle for the superseded grammar,
            # not a second authority for current Markdown semantics.  Keep
            # already-typed v3 constructs literal here; their authoritative
            # projection was produced from raw source before Mistune parsing.
            value = str(child.get("raw", ""))
            if value:
                if projected and projected[-1].get("type") == "text":
                    projected[-1]["value"] += value
                else:
                    projected.append({"type": "text", "value": value})
            continue
        if child_type == "semantic_cross_reference":
            projected.append(
                {
                    "type": "cross_reference",
                    "target_id": child["target_id"],
                    "target_kind": child["target_kind"],
                    "status": child["status"],
                    "display": "auto",
                }
            )
        elif child_type == "semantic_citation":
            projected.append(
                {
                    "type": "citation",
                    "items": [{"key": key} for key in child["items"]],
                    "raw": child["raw"],
                    "status": "unresolved",
                }
            )
        else:
            value = _plain_inline_text([child])
            if value:
                if projected and projected[-1].get("type") == "text":
                    projected[-1]["value"] += value
                else:
                    projected.append({"type": "text", "value": value})
    return projected


def _plain_inline_text(children: list[dict[str, Any]]) -> str:
    parts: list[str] = []
    for child in children:
        child_type = child.get("type")
        if child_type == "text":
            parts.append(child.get("raw", "") or child.get("text", "") or "")
        elif child_type in {"semantic_cross_reference", "semantic_citation"}:
            parts.append(child.get("raw") or f"@{child.get('target_id', '')}")
        elif child_type in {"image", "strong", "emphasis", "strikethrough", "link"}:
            parts.append(_plain_inline_text(child.get("children", [])))
        else:
            parts.append(child.get("raw", "") or child.get("text", "") or "")
    return "".join(parts)


def _deduplicate_and_sort_diagnostics(
    diagnostics: list[SemanticDiagnostic],
) -> list[SemanticDiagnostic]:
    unique: dict[tuple[str, str, str], SemanticDiagnostic] = {}
    for diagnostic in diagnostics:
        unique[(diagnostic.level, diagnostic.code, diagnostic.message)] = diagnostic
    return sorted(
        unique.values(),
        key=lambda item: (_DIAGNOSTIC_ORDER.get(item.code, 999), item.message),
    )

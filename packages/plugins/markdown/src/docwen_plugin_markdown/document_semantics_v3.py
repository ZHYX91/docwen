"""Authenticated Markdown source semantics for ``docwen.markdown_semantics.v3``.

This module intentionally parses authored source before Mistune or the generic
link preprocessor.  The v3 oracle needs lossless tokens and Unicode code-point
ranges; an AST produced by a renderer is too late to recover either fact.

The API is source-only.  It does not scan a Workspace, open linked pages, look
up citations, or project Word fields.  An adapter that owns those domains may
supply closed, consumer-neutral resolution records.
"""

from __future__ import annotations

import base64
import hashlib
import re
from collections.abc import Iterable, Mapping, Sequence
from dataclasses import dataclass
from typing import Any, Literal, cast

from docwen_plugin_markdown.document_semantics_v3_fenced_source import (
    project_fenced_source_v3,
)

SEMANTICS_SCHEMA = "docwen.markdown_semantics.v3"
SEMANTICS_SCHEMA_ID = "urn:docwen:schema:markdown-semantics:v3"
DIAGNOSTICS_SCHEMA = "docwen.markdown_diagnostics.v3"
DIAGNOSTICS_SCHEMA_ID = "urn:docwen:schema:markdown-diagnostics:v3"
DIAGNOSTIC_EVIDENCE_SCHEMA = "docwen.machine.diagnostic_evidence.v1"

_ID_RE = re.compile(r"^[A-Za-z0-9-]{1,128}$")
_CITATION_KEY_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9_-]{0,127}$")
_CAPTION_RE = re.compile(r"^(Figure|Table|Equation|Code):(.*)$", re.IGNORECASE)
_HEADING_RE = re.compile(r"^( {0,3})(#{1,9})(?!#)[ \t]+(.+?)\s*$")
_FENCE_RE = re.compile(r"^( {0,3})(?P<fence>`{3,}|~{3,})(?P<info>[^\r\n]*)$")
_QUOTE_PREFIX_RE = re.compile(r"^ {0,3}>[ \t]?")
_UNSUPPORTED_LINE_SEPARATOR_RE = re.compile(r"\r(?!\n)|[\v\f\x1c-\x1e\x85\u2028\u2029]")


def _match_fence_opener_v3(text: str) -> re.Match[str] | None:
    """Match one CommonMark-valid fenced-code opening line."""

    match = _FENCE_RE.fullmatch(text)
    if match is None:
        return None
    if match.group("fence").startswith("`") and "`" in match.group("info"):
        return None
    return match


_MATH_SINGLE_RE = re.compile(r"^ {0,3}\$\$(?P<body>.+?)\$\$[ \t]*$")
_RESOURCELESS_IMAGE_CARRIER = "![image omitted]()"
# Ordinary Markdown images still require a non-empty destination.  The one
# empty-destination exception is a fixed DOCX-recovery carrier, not a lookup.
_IMAGE_RE = re.compile(
    rf"^\s*(?:!\[[^\]]*\]\([^\n)]+\)|!\[\[[^\]\n]+\]\]|{re.escape(_RESOURCELESS_IMAGE_CARRIER)})\s*$"
)
_LIST_ITEM_RE = re.compile(r"^(?P<indent>[ \t]*)(?:[-+*]|\d+[.)])[ \t]+")
_TABLE_DELIMITER_CELL_RE = re.compile(r"^:?-{3,}:?$")
_ANCHOR_ONLY_CANDIDATE_RE = re.compile(r"^(?P<indent>[ \t]*)(?P<token>\^[^\s]*)[ \t]*$")
_INLINE_ANCHOR_CANDIDATE_RE = re.compile(r"(?P<space>[ \t]+)(?P<token>\^[^\s]*)[ \t]*$")
_SEMANTIC_REFERENCE_RE = re.compile(r"@\[\[(?P<body>[^\]\r\n]+)\]\]")
_WIKILINK_RE = re.compile(r"(?P<embed>!)?\[\[(?P<body>[^\]\r\n]+)\]\]")
_PARENTHETICAL_CITATION_RE = re.compile(
    r"\[@(?P<first>[A-Za-z0-9][A-Za-z0-9_-]{0,127})(?P<rest>(?:\s*;\s*@[A-Za-z0-9][A-Za-z0-9_-]{0,127})+)\]"
)
_NARRATIVE_CITATION_RE = re.compile(r"(?<![A-Za-z0-9._%+@-])@(?P<key>[A-Za-z0-9][A-Za-z0-9_-]{0,127})")
_YAML_FRONT_OPEN_RE = re.compile(r"^(?:\ufeff)?---[ \t]*(?:\r?\n)")
_YAML_FRONT_CLOSE_RE = re.compile(r"^---[ \t]*(?:\r?\n|$)", re.MULTILINE)

type TargetKind = Literal["heading", "figure", "table", "equation", "code_block"]
type ResolutionStatus = Literal[
    "resolved",
    "missing",
    "ambiguous",
    "non_semantic",
    "unnumbered",
    "external_unresolved",
]

_CAPTION_KIND_BY_KEYWORD: dict[str, TargetKind] = {
    "figure": "figure",
    "table": "table",
    "equation": "equation",
    "code": "code_block",
}


@dataclass(frozen=True, slots=True)
class SourceRange:
    """A half-open Unicode code-point range in authenticated source."""

    start: int
    end: int

    def as_dict(self) -> dict[str, int]:
        return {"start": self.start, "end": self.end}


@dataclass(frozen=True, slots=True)
class ExternalReferenceResolution:
    """One neutral cross-document resolution supplied by an external owner."""

    page_locator: str
    selector_kind: Literal["stable_id", "heading_path"]
    resolved_document_id: str
    resolved_document_sha256: str
    resolved_kind: TargetKind
    cached_number: str | None
    current_title: str
    target_id: str | None = None
    heading_path: tuple[str, ...] = ()


@dataclass(frozen=True, slots=True)
class ExternalCitationResolution:
    """One neutral citation-key result; the key is not record identity."""

    key: str
    record_id: str
    record_sha256: str
    presentation: str


@dataclass(frozen=True, slots=True)
class MarkdownSemanticsV3Analysis:
    """Closed source projection plus exact source-backed diagnostics."""

    projection: dict[str, Any]
    diagnostics: tuple[dict[str, Any], ...]

    @property
    def has_errors(self) -> bool:
        return any(item["severity"] == "error" for item in self.diagnostics)

    def authored_tokens(self) -> tuple[str, ...]:
        """Return every semantic/link/citation token exactly as authored."""

        records = [*self.projection["links"], *self.projection["references"], *self.projection["citations"]]
        records.sort(key=lambda item: (item["range"]["start"], item["range"]["end"]))
        return tuple(str(item["raw"]) for item in records)


def is_resource_less_image_carrier_v3(source: str) -> bool:
    """Classify only the fixed DOCX-recovery image token as resource-less."""

    return source.strip() == _RESOURCELESS_IMAGE_CARRIER


@dataclass(frozen=True, slots=True)
class _Line:
    text: str
    start: int
    content_end: int
    end: int
    number: int
    source_start: int


@dataclass(slots=True)
class _Block:
    kind: str
    start: int
    end: int
    first_line: int
    last_line: int
    text: str
    data: dict[str, Any]


@dataclass(frozen=True, slots=True)
class _Owner:
    owner_kind: str
    semantic_kind: TargetKind | None
    id_range: SourceRange
    record_index: int


def analyze_markdown_semantics_v3(
    source: str,
    *,
    input_id: str,
    external_references: Sequence[ExternalReferenceResolution] = (),
    external_citations: Sequence[ExternalCitationResolution] = (),
    rename_replacements: Mapping[int, str] | None = None,
    semantic_id_replacements: Mapping[int, str] | None = None,
) -> MarkdownSemanticsV3Analysis:
    """Parse one authenticated Markdown source into the v3 source oracle.

    The replacement maps are optional caller intent keyed by diagnostic start
    offset.  They only enable a fix when the replacement is already valid and
    conflict-free; the parser never invents an identifier.
    """

    if not input_id or len(input_id) > 256:
        raise ValueError("input_id must contain 1..256 characters")
    unsupported_separator = _UNSUPPORTED_LINE_SEPARATOR_RE.search(source)
    if unsupported_separator is not None:
        raise ValueError(
            "Markdown semantics v3 accepts only LF and CRLF line endings "
            f"(unsupported separator at offset {unsupported_separator.start()})"
        )
    replacements = dict(rename_replacements or {})
    semantic_replacements = dict(semantic_id_replacements or {})
    source_sha256 = hashlib.sha256(source.encode("utf-8")).hexdigest()
    source_identity = _source_identity(input_id, source_sha256)
    semantic_source = _mask_yaml_front_matter(source)
    lines = _split_lines(semantic_source)
    blocks = _scan_blocks(lines)

    diagnostics: list[dict[str, Any]] = []
    targets: list[dict[str, Any]] = []
    anchors: list[dict[str, Any]] = []
    links: list[dict[str, Any]] = []
    owners: dict[str, list[_Owner]] = {}
    excluded_inline_ranges: list[SourceRange] = []

    heading_counters = [0] * 9
    caption_counters: dict[TargetKind, int] = {
        "figure": 0,
        "table": 0,
        "equation": 0,
        "code_block": 0,
        "heading": 0,
    }
    heading_records: list[dict[str, Any]] = []
    active_heading_titles: list[str] = []

    # Headings and declarations own semantic IDs.  Raw objects never do.
    for index, block in enumerate(blocks):
        if block.kind == "heading":
            level = int(block.data["level"])
            title, anchor = _extract_inline_anchor(
                block.data["content"],
                absolute_start=int(block.data["content_start"]),
            )
            title = title.strip()
            heading_counters[level - 1] += 1
            for position in range(level, 9):
                heading_counters[position] = 0
            number = ".".join(str(value) for value in heading_counters[:level] if value)
            active_heading_titles = active_heading_titles[: level - 1]
            active_heading_titles.append(title)
            record: dict[str, Any] = {
                "kind": "heading",
                "title": title,
                "heading_level": level,
                "heading_path": list(active_heading_titles),
                "number": number,
                "source_form": "heading",
                "range": SourceRange(block.start, block.end).as_dict(),
            }
            if anchor is not None:
                token, token_range = anchor
                excluded_inline_ranges.append(token_range)
                anchor_id = token[1:]
                if _valid_id(anchor_id):
                    record["id"] = anchor_id
                    record["id_range"] = token_range.as_dict()
                    _register_owner(
                        owners,
                        anchor_id,
                        _Owner("semantic_target", "heading", token_range, len(targets)),
                    )
                else:
                    diagnostics.append(
                        _invalid_anchor_diagnostic(
                            source_identity,
                            token_range,
                            anchor_id,
                            replacements,
                            owners,
                        )
                    )
            targets.append(record)
            heading_records.append(record)
            continue

        if block.kind != "caption_declaration":
            continue
        declaration_kind: TargetKind = block.data["kind"]
        content, anchor = _extract_inline_anchor(
            block.data["content"],
            absolute_start=int(block.data["content_start"]),
        )
        content = content.strip()
        anchor_id: str | None = None
        anchor_range: SourceRange | None = None
        if anchor is not None:
            token, anchor_range = anchor
            excluded_inline_ranges.append(anchor_range)
            candidate = token[1:]
            if _valid_id(candidate):
                anchor_id = candidate
            else:
                diagnostics.append(
                    _invalid_anchor_diagnostic(
                        source_identity,
                        anchor_range,
                        candidate,
                        replacements,
                        owners,
                    )
                )

        keyword_range = SourceRange(int(block.data["keyword_start"]), int(block.data["keyword_end"]))
        if declaration_kind in {"figure", "table", "code_block"} and not content:
            diagnostics.append(
                _diagnostic(
                    source_identity,
                    "error",
                    "docwen.markdown.caption.content_required",
                    f"{block.data['canonical_keyword']} requires non-empty caption text.",
                    keyword_range,
                )
            )
            continue
        if declaration_kind == "equation" and not content and anchor_id is None:
            insertion_offset = int(block.data["line_content_end"])
            insertion = SourceRange(insertion_offset, insertion_offset)
            fixes: tuple[dict[str, Any], ...] = ()
            replacement_id = semantic_replacements.get(keyword_range.start)
            if replacement_id is not None and _valid_id(replacement_id) and replacement_id not in owners:
                fixes = (
                    {
                        "fix_id": "docwen.markdown.fix.add_semantic_id",
                        "edits": [{"range": insertion.as_dict(), "replacement": f" ^{replacement_id}"}],
                    },
                )
            diagnostics.append(
                _diagnostic(
                    source_identity,
                    "error",
                    "docwen.markdown.caption.empty_equation_target_required",
                    "An empty Equation declaration requires an explicit semantic ID.",
                    keyword_range,
                    fixes=fixes,
                )
            )
            continue

        object_index = _next_non_marker_block(
            blocks,
            index + 1,
            container_path=tuple(block.data["container_path"]),
        )
        if object_index is None:
            diagnostics.append(
                _diagnostic(
                    source_identity,
                    "error",
                    "docwen.markdown.caption.object_mismatch",
                    f"{block.data['canonical_keyword']} is not followed by its matching object.",
                    keyword_range,
                )
            )
            continue
        object_block = blocks[object_index]
        expected_object = {
            "figure": "image",
            "table": "table",
            "equation": "equation",
            "code_block": "code_block",
        }[declaration_kind]
        if object_block.kind != expected_object:
            diagnostics.append(
                _diagnostic(
                    source_identity,
                    "error",
                    "docwen.markdown.caption.object_mismatch",
                    f"{block.data['canonical_keyword']} is followed by {object_block.kind}, not {expected_object}.",
                    SourceRange(object_block.start, object_block.end),
                )
            )
            continue

        caption_counters[declaration_kind] += 1
        record = {
            "kind": declaration_kind,
            "title": content,
            "number": str(caption_counters[declaration_kind]),
            "source_form": "declaration",
            "source_keyword": block.data["source_keyword"],
            "range": SourceRange(block.start, object_block.end).as_dict(),
            "declaration_range": SourceRange(block.start, block.end).as_dict(),
            "object_range": SourceRange(object_block.start, object_block.end).as_dict(),
        }
        if anchor_id is not None and anchor_range is not None:
            record["id"] = anchor_id
            record["id_range"] = anchor_range.as_dict()
            _register_owner(
                owners,
                anchor_id,
                _Owner("semantic_target", declaration_kind, anchor_range, len(targets)),
            )
        targets.append(record)

    # Ordinary inline and post-block anchors share the same ID namespace.
    for block in blocks:
        if block.kind in {
            "heading",
            "caption_declaration",
            "anchor_marker",
            "blank",
            "list",
        }:
            continue
        inline = _block_inline_anchor(block)
        if inline is None:
            continue
        token, token_range = inline
        if _overlaps_any(token_range, excluded_inline_ranges) or _has_nested_inline_owner(
            blocks,
            block,
            token_range,
        ):
            continue
        excluded_inline_ranges.append(token_range)
        anchor_id = token[1:]
        if not _valid_id(anchor_id):
            diagnostics.append(
                _invalid_anchor_diagnostic(source_identity, token_range, anchor_id, replacements, owners)
            )
            continue
        inline_kind_supported = block.kind in {"paragraph", "container_text", "image", "list_item"}
        if block.kind == "container_text" and block.first_line != block.last_line:
            inline_kind_supported = False
        if not inline_kind_supported:
            diagnostics.append(
                _diagnostic(
                    source_identity,
                    "error",
                    "docwen.markdown.anchor.invalid_id",
                    "This block kind requires an anchor-only line after the complete block.",
                    token_range,
                )
            )
            continue
        anchors.append(
            {
                "id": anchor_id,
                "block_kind": "paragraph" if block.kind == "container_text" else block.kind,
                "placement": "inline",
                "range": token_range.as_dict(),
                "block_range": _ordinary_anchor_owner_range(block).as_dict(),
                "container_path": _container_path_projection(block),
            }
        )
        _register_owner(owners, anchor_id, _Owner("ordinary_anchor", None, token_range, len(anchors) - 1))

    for index, block in enumerate(blocks):
        if block.kind != "anchor_marker":
            continue
        token = str(block.data["token"])
        token_range = SourceRange(int(block.data["token_start"]), int(block.data["token_end"]))
        excluded_inline_ranges.append(token_range)
        anchor_id = token[1:]
        if not _valid_id(anchor_id):
            diagnostics.append(
                _invalid_anchor_diagnostic(source_identity, token_range, anchor_id, replacements, owners)
            )
            continue
        previous_index = _previous_non_marker_block(
            blocks,
            index - 1,
            container_path=tuple(block.data["container_path"]),
        )
        attachable = (
            previous_index is not None
            and blocks[previous_index].kind
            in {"list", "block_quote", "callout", "table", "equation", "code_block", "fenced_block"}
            and int(block.data["container_indent"]) == int(blocks[previous_index].data.get("container_indent", 0))
        )
        if not attachable:
            diagnostics.append(
                _diagnostic(
                    source_identity,
                    "error",
                    "docwen.markdown.anchor.dangling",
                    "An anchor-only line has no attachable preceding structured block.",
                    token_range,
                )
            )
            continue
        assert previous_index is not None
        owner_block = blocks[previous_index]
        anchors.append(
            {
                "id": anchor_id,
                "block_kind": owner_block.kind,
                "placement": "post_block",
                "range": token_range.as_dict(),
                "block_range": SourceRange(owner_block.start, owner_block.end).as_dict(),
                "container_path": _container_path_projection(owner_block),
            }
        )
        _register_owner(owners, anchor_id, _Owner("ordinary_anchor", None, token_range, len(anchors) - 1))

    # Duplicate ownership is document-wide and reported on every later owner.
    for anchor_id, occurrences in owners.items():
        first = occurrences[0]
        for later in occurrences[1:]:
            fixes: tuple[dict[str, Any], ...] = ()
            replacement = replacements.get(later.id_range.start)
            if replacement is not None and _valid_id(replacement) and replacement not in owners:
                fixes = (
                    {
                        "fix_id": "docwen.markdown.fix.rename_anchor",
                        "edits": [{"range": later.id_range.as_dict(), "replacement": f"^{replacement}"}],
                    },
                )
            diagnostics.append(
                _diagnostic(
                    source_identity,
                    "error",
                    "docwen.markdown.anchor.duplicate",
                    f"Anchor ID {anchor_id} already has an owner in this document.",
                    later.id_range,
                    related_ranges=(first.id_range,),
                    fixes=fixes,
                )
            )

    external_reference_map = _validate_external_references(external_references)
    external_citation_map = _validate_external_citations(external_citations)
    references: list[dict[str, Any]] = []
    citations: list[dict[str, Any]] = []

    excluded = _merge_ranges(excluded_inline_ranges + _literal_shield_ranges(semantic_source, blocks))
    semantic_matches = list(_SEMANTIC_REFERENCE_RE.finditer(semantic_source))
    occupied: list[SourceRange] = []
    for match in semantic_matches:
        token_range = SourceRange(match.start(), match.end())
        if _overlaps_any(token_range, excluded):
            continue
        occupied.append(token_range)
        reference, reference_diagnostics = _resolve_reference(
            source_identity,
            match,
            token_range,
            targets,
            heading_records,
            owners,
            external_reference_map,
        )
        references.append(reference)
        diagnostics.extend(reference_diagnostics)

    for match in _WIKILINK_RE.finditer(semantic_source):
        token_range = SourceRange(match.start(), match.end())
        if _overlaps_any(token_range, [*excluded, *occupied]):
            continue
        links.append(_project_wikilink(match, token_range))
        occupied.append(token_range)

    for match in _PARENTHETICAL_CITATION_RE.finditer(semantic_source):
        token_range = SourceRange(match.start(), match.end())
        if _overlaps_any(token_range, [*excluded, *occupied]):
            continue
        raw = match.group(0)
        item_matches = list(re.finditer(r"@([A-Za-z0-9][A-Za-z0-9_-]{0,127})", raw))
        items = [
            _citation_item(
                item.group(1),
                SourceRange(match.start() + item.start(1), match.start() + item.end(1)),
                external_citation_map,
            )
            for item in item_matches
        ]
        citations.append(
            {
                "form": "parenthetical",
                "raw": raw,
                "range": token_range.as_dict(),
                "items": items,
            }
        )
        occupied.append(token_range)

    for match in _NARRATIVE_CITATION_RE.finditer(semantic_source):
        token_range = SourceRange(match.start(), match.end())
        if _overlaps_any(token_range, [*excluded, *occupied]):
            continue
        key = match.group("key")
        if not _CITATION_KEY_RE.fullmatch(key):
            continue
        citations.append(
            {
                "form": "narrative",
                "raw": match.group(0),
                "range": token_range.as_dict(),
                "items": [
                    _citation_item(
                        key,
                        SourceRange(match.start("key"), match.end("key")),
                        external_citation_map,
                    )
                ],
            }
        )
        occupied.append(token_range)

    projection = {
        "$schema": SEMANTICS_SCHEMA_ID,
        "schema": SEMANTICS_SCHEMA,
        "source": source_identity,
        "targets": sorted(targets, key=lambda item: (item["range"]["start"], item["kind"])),
        "anchors": sorted(anchors, key=lambda item: item["range"]["start"]),
        "fenced_sources": _project_fenced_sources(source, blocks, source_sha256),
        "links": sorted(links, key=lambda item: item["range"]["start"]),
        "references": sorted(references, key=lambda item: item["range"]["start"]),
        "citations": sorted(citations, key=lambda item: item["range"]["start"]),
    }
    ordered_diagnostics = tuple(
        sorted(
            diagnostics,
            key=lambda item: (item["range"]["start"], item["range"]["end"], item["code"]),
        )
    )
    return MarkdownSemanticsV3Analysis(projection=projection, diagnostics=ordered_diagnostics)


def _project_fenced_sources(source: str, blocks: Sequence[_Block], source_sha256: str) -> list[dict[str, Any]]:
    """Project exact authored framing for every unique fenced occurrence."""

    by_range: dict[tuple[int, int], _Block] = {}
    for block in blocks:
        if block.kind not in {"code_block", "fenced_block"}:
            continue
        if block.data["invalid_pseudo_closer"]:
            # A fence-like line with an anchor suffix terminates dialect
            # scanning only so the stable token diagnostic can be emitted.
            # It is never an authenticated CommonMark closer or omitted-EOF
            # boundary, including when that line is physically at EOF.
            continue
        key = (block.start, block.end)
        current = by_range.get(key)
        if current is None or len(tuple(block.data["container_path"])) > len(tuple(current.data["container_path"])):
            by_range[key] = block
    records = [_project_fenced_source(source, block, source_sha256) for block in by_range.values()]
    return sorted(records, key=lambda item: (item["source_start"], item["source_end"], item["identity_sha256"]))


def _project_fenced_source(source: str, block: _Block, source_sha256: str) -> dict[str, Any]:
    lines = _source_lines(source[block.start : block.end], absolute_start=block.start)
    if not lines:
        raise RuntimeError("fenced source range is empty")
    opener = lines[0]
    closer = lines[-1]
    return project_fenced_source_v3(
        source,
        source_sha256=source_sha256,
        source_start=block.start,
        source_end=block.end,
        fence=str(block.data["fence"]),
        fence_start=int(block.data["fence_start"]),
        opening_eol=str(opener["eol"]),
        body_line_coordinates=tuple(block.data["body_line_coordinates"]),
        closing_present=bool(block.data["closing_present"]),
        closing_fence_start=int(block.data["closing_fence_start"]),
        closing_fence_end=int(block.data["closing_fence_end"]),
        closing_line_start=int(closer["start"]),
        closing_content_end=int(closer["content_end"]),
        closing_eol=str(closer["eol"]),
    )


def _source_lines(value: str, *, absolute_start: int) -> list[dict[str, Any]]:
    output: list[dict[str, Any]] = []
    offset = 0
    for raw in value.splitlines(keepends=True):
        text = raw.rstrip("\r\n")
        eol = raw[len(text) :]
        output.append(
            {
                "text": text,
                "eol": eol,
                "start": absolute_start + offset,
                "content_end": absolute_start + offset + len(text),
            }
        )
        offset += len(raw)
    if offset < len(value):
        output.append(
            {
                "text": value[offset:],
                "eol": "",
                "start": absolute_start + offset,
                "content_end": absolute_start + len(value),
            }
        )
    return output


def _b64(value: str) -> str:
    return base64.b64encode(value.encode("utf-8")).decode("ascii")


def render_cross_reference_token(
    *,
    selector_kind: Literal["stable_id", "heading_path"],
    page_locator: str | None = None,
    target_id: str | None = None,
    heading_path: Sequence[str] = (),
    alias: str | None = None,
) -> str:
    """Serialize one newly created cross-reference without guessing identity."""

    if selector_kind == "stable_id":
        if target_id is None or not _valid_id(target_id):
            raise ValueError("stable_id references require a valid target_id")
        fragment = f"^{target_id}"
    else:
        if target_id is not None or not heading_path or any(not item or "#" in item for item in heading_path):
            raise ValueError("heading_path references require non-empty path segments and no target_id")
        fragment = "#".join(heading_path)
    page = page_locator or ""
    if any(token in page for token in ("#", "|", "[", "]")):
        raise ValueError("page_locator contains cross-reference delimiters")
    alias_text = "" if alias is None else f"|{alias}"
    if alias is not None and (not alias or any(token in alias for token in ("[", "]"))):
        raise ValueError("alias must be non-empty and cannot contain brackets")
    return f"@[[{page}#{fragment}{alias_text}]]"


def render_citation_token(keys: Sequence[str], *, parenthetical: bool) -> str:
    """Serialize a citation while keeping mutable keys separate from identity."""

    if not keys or any(_CITATION_KEY_RE.fullmatch(key) is None for key in keys):
        raise ValueError("citation keys must match the v3 citation-key grammar")
    if parenthetical:
        return "[" + "; ".join(f"@{key}" for key in keys) + "]"
    if len(keys) != 1:
        raise ValueError("a narrative citation contains exactly one key")
    return f"@{keys[0]}"


def render_caption_declaration(kind: TargetKind, content: str, *, target_id: str | None = None) -> str:
    """Write a newly created declaration with canonical keyword casing."""

    keyword = {
        "figure": "Figure",
        "table": "Table",
        "equation": "Equation",
        "code_block": "Code",
    }.get(kind)
    if keyword is None:
        raise ValueError("Heading is not a caption declaration")
    trimmed = content.strip()
    if kind in {"figure", "table", "code_block"} and not trimmed:
        raise ValueError(f"{keyword} requires non-empty caption text")
    if kind == "equation" and not trimmed and target_id is None:
        raise ValueError("an empty Equation declaration requires a target_id")
    if target_id is not None and not _valid_id(target_id):
        raise ValueError("target_id must match [A-Za-z0-9-]{1,128}")
    suffix = "" if target_id is None else f"^{target_id}"
    payload = " ".join(item for item in (trimmed, suffix) if item)
    return f"{keyword}: {payload}".rstrip()


def select_safe_fence(body: str) -> str:
    """Choose the shortest safe new CommonMark fence with backtick tie-break."""

    def required(character: str) -> int:
        closing_line = re.compile(rf"^ {{0,3}}(?P<run>{re.escape(character)}{{3,}})[ \t]*$")
        longest = max(
            (
                len(match.group("run"))
                for line in body.splitlines()
                if (match := closing_line.fullmatch(line)) is not None
            ),
            default=0,
        )
        return max(3, longest + 1)

    backticks = required("`")
    tildes = required("~")
    if backticks <= tildes:
        return "`" * backticks
    return "~" * tildes


def markdown_semantics_body_start_v3(source: str) -> int:
    """Return the exact body offset while treating closed YAML as non-Markdown."""

    opening = _YAML_FRONT_OPEN_RE.match(source)
    if opening is None:
        return 0
    closing = _YAML_FRONT_CLOSE_RE.search(source, opening.end())
    return closing.end() if closing is not None else 0


def _mask_yaml_front_matter(source: str) -> str:
    end = markdown_semantics_body_start_v3(source)
    if end == 0:
        return source
    return "".join(character if character in "\r\n" else " " for character in source[:end]) + source[end:]


def _split_lines(source: str) -> list[_Line]:
    lines: list[_Line] = []
    offset = 0
    for number, raw in enumerate(source.splitlines(keepends=True), start=1):
        text = raw.rstrip("\r\n")
        content_end = offset + len(text)
        lines.append(
            _Line(
                text=text,
                start=offset,
                content_end=content_end,
                end=offset + len(raw),
                number=number,
                source_start=offset,
            )
        )
        offset += len(raw)
    if not lines or offset < len(source) or (source and source[-1] not in "\r\n"):
        if offset < len(source):
            text = source[offset:]
            lines.append(
                _Line(
                    text=text,
                    start=offset,
                    content_end=len(source),
                    end=len(source),
                    number=len(lines) + 1,
                    source_start=offset,
                )
            )
        elif not lines:
            lines.append(_Line(text="", start=0, content_end=0, end=0, number=1, source_start=0))
    return lines


def _scan_blocks(lines: Sequence[_Line]) -> list[_Block]:
    blocks = _scan_container_blocks(
        lines,
        container_path=(),
        container_segments=(),
        paragraph_kind="paragraph",
    )
    return sorted(
        blocks,
        key=lambda block: (
            block.start,
            0 if block.kind in {"heading", "caption_declaration"} else 1,
            block.end,
            len(tuple(block.data["container_path"])),
        ),
    )


def _scan_container_blocks(
    lines: Sequence[_Line],
    *,
    container_path: tuple[tuple[str, int], ...],
    container_segments: tuple[tuple[str, int, int], ...],
    paragraph_kind: str,
) -> list[_Block]:
    blocks: list[_Block] = []
    index = 0
    while index < len(lines):
        line = lines[index]
        if not line.text.strip():
            blocks.append(_Block("blank", line.source_start, line.end, index, index, line.text, {}))
            index += 1
            continue
        anchor_match = _ANCHOR_ONLY_CANDIDATE_RE.fullmatch(line.text)
        if anchor_match is not None:
            token_start = line.start + anchor_match.start("token")
            token_end = line.start + anchor_match.end("token")
            blocks.append(
                _Block(
                    "anchor_marker",
                    line.source_start,
                    line.end,
                    index,
                    index,
                    line.text,
                    {
                        "token": anchor_match.group("token"),
                        "token_start": token_start,
                        "token_end": token_end,
                        "container_indent": len(anchor_match.group("indent")),
                    },
                )
            )
            index += 1
            continue
        fence_match = _match_fence_opener_v3(line.text)
        if fence_match is not None:
            character = fence_match.group("fence")[0]
            length = len(fence_match.group("fence"))
            end_index = index
            close_re = re.compile(
                rf"^(?P<prefix> {{0,3}})(?P<fence>{re.escape(character)}{{{length},}})(?P<suffix>[ \t]*)$"
            )
            close_with_anchor_re = re.compile(rf"^ {{0,3}}{re.escape(character)}{{{length},}}[ \t]+\^[^\s]*[ \t]*$")
            closing_match: re.Match[str] | None = None
            invalid_pseudo_closer = False
            for candidate in range(index + 1, len(lines)):
                end_index = candidate
                candidate_match = close_re.fullmatch(lines[candidate].text)
                pseudo_closer_match = close_with_anchor_re.fullmatch(lines[candidate].text)
                if candidate_match or pseudo_closer_match:
                    closing_match = candidate_match
                    invalid_pseudo_closer = pseudo_closer_match is not None
                    break
            info = fence_match.group("info").strip()
            language = info.split(maxsplit=1)[0].casefold() if info else ""
            fenced_block_kind = "code_block" if language not in {"mermaid", "query", "view"} else "fenced_block"
            body_end_index = end_index if closing_match is None else end_index - 1
            body_coordinates: list[tuple[int, int, int, int]] = []
            opening_indent = len(fence_match.group(1))
            for body_line in lines[index + 1 : body_end_index + 1]:
                removable_indent = min(opening_indent, len(body_line.text) - len(body_line.text.lstrip(" ")))
                body_coordinates.append(
                    (
                        body_line.source_start,
                        body_line.start + removable_indent,
                        body_line.content_end,
                        body_line.end,
                    )
                )
            blocks.append(
                _Block(
                    fenced_block_kind,
                    line.source_start,
                    lines[end_index].end,
                    index,
                    end_index,
                    "\n".join(item.text for item in lines[index : end_index + 1]),
                    {
                        "fence": fence_match.group("fence"),
                        "fence_start": line.start + fence_match.start("fence"),
                        "info": info,
                        "fenced_kind": language or "code",
                        "last_line_start": lines[end_index].start,
                        "container_indent": len(fence_match.group(1)),
                        "body_line_coordinates": tuple(body_coordinates),
                        "closing_present": closing_match is not None,
                        "invalid_pseudo_closer": invalid_pseudo_closer,
                        "closing_fence_start": (
                            lines[end_index].start + closing_match.start("fence") if closing_match is not None else 0
                        ),
                        "closing_fence_end": (
                            lines[end_index].start + closing_match.end("fence") if closing_match is not None else 0
                        ),
                    },
                )
            )
            index = end_index + 1
            continue
        if line.text.lstrip().startswith("$$"):
            if _MATH_SINGLE_RE.fullmatch(line.text):
                blocks.append(
                    _Block(
                        "equation",
                        line.source_start,
                        line.end,
                        index,
                        index,
                        line.text,
                        {"last_line_start": line.start, "container_indent": len(line.text) - len(line.text.lstrip())},
                    )
                )
                index += 1
                continue
            if line.text.strip() == "$$":
                end_index = index
                for candidate in range(index + 1, len(lines)):
                    end_index = candidate
                    if lines[candidate].text.strip() == "$$" or re.fullmatch(
                        r" {0,3}\$\$[ \t]+\^[^\s]*[ \t]*", lines[candidate].text
                    ):
                        break
                blocks.append(
                    _Block(
                        "equation",
                        line.source_start,
                        lines[end_index].end,
                        index,
                        end_index,
                        "\n".join(item.text for item in lines[index : end_index + 1]),
                        {
                            "last_line_start": lines[end_index].start,
                            "container_indent": len(line.text) - len(line.text.lstrip()),
                        },
                    )
                )
                index = end_index + 1
                continue
        heading = _HEADING_RE.fullmatch(line.text)
        if heading is not None:
            content_start = line.start + heading.start(3)
            blocks.append(
                _Block(
                    "heading",
                    line.source_start,
                    line.end,
                    index,
                    index,
                    line.text,
                    {"level": len(heading.group(2)), "content": heading.group(3), "content_start": content_start},
                )
            )
            index += 1
            continue
        caption = _CAPTION_RE.fullmatch(line.text)
        if caption is not None and re.search(r"\{#[^{}\s]+\}[ \t]*$", caption.group(2)):
            # Historical Pandoc-style attributes are ordinary current source.
            # Only the explicit migration module may reinterpret them.
            caption = None
        if caption is not None:
            source_keyword = caption.group(1)
            caption_kind = _CAPTION_KIND_BY_KEYWORD[source_keyword.casefold()]
            blocks.append(
                _Block(
                    "caption_declaration",
                    line.source_start,
                    line.end,
                    index,
                    index,
                    line.text,
                    {
                        "kind": caption_kind,
                        "source_keyword": source_keyword,
                        "canonical_keyword": {"code_block": "Code"}.get(caption_kind, caption_kind.title()),
                        "content": caption.group(2),
                        "content_start": line.start + caption.start(2),
                        "keyword_start": line.start + caption.start(1),
                        "keyword_end": line.start + caption.end(1),
                        "line_content_end": line.content_end,
                    },
                )
            )
            index += 1
            continue
        if index + 1 < len(lines) and "|" in line.text and _is_table_delimiter(lines[index + 1].text):
            end_index = index + 1
            while end_index + 1 < len(lines) and lines[end_index + 1].text.strip() and "|" in lines[end_index + 1].text:
                end_index += 1
            blocks.append(
                _Block(
                    "table",
                    line.source_start,
                    lines[end_index].end,
                    index,
                    end_index,
                    "\n".join(item.text for item in lines[index : end_index + 1]),
                    {
                        "last_line_start": lines[end_index].start,
                        "container_indent": len(line.text) - len(line.text.lstrip()),
                    },
                )
            )
            index = end_index + 1
            continue
        if _QUOTE_PREFIX_RE.match(line.text) is not None:
            end_index = index
            while end_index + 1 < len(lines) and _QUOTE_PREFIX_RE.match(lines[end_index + 1].text) is not None:
                end_index += 1
            quote_kind = "callout" if re.match(r"^\s*>\s*\[!", line.text, re.IGNORECASE) else "block_quote"
            blocks.append(
                _Block(
                    quote_kind,
                    line.source_start,
                    lines[end_index].end,
                    index,
                    end_index,
                    "\n".join(item.text for item in lines[index : end_index + 1]),
                    {
                        "last_line_start": lines[end_index].start,
                        "container_indent": len(line.text) - len(line.text.lstrip()),
                    },
                )
            )
            index = end_index + 1
            continue
        if _LIST_ITEM_RE.match(line.text):
            end_index = index
            while end_index + 1 < len(lines):
                following = lines[end_index + 1].text
                if not following.strip() or _LIST_ITEM_RE.match(following) or following.startswith(("  ", "\t")):
                    end_index += 1
                    continue
                break
            blocks.append(
                _Block(
                    "list",
                    line.source_start,
                    lines[end_index].end,
                    index,
                    end_index,
                    "\n".join(item.text for item in lines[index : end_index + 1]),
                    {
                        "last_line_start": lines[end_index].start,
                        "container_indent": len(line.text) - len(line.text.lstrip()),
                    },
                )
            )
            index = end_index + 1
            continue
        image_source = _strip_valid_inline_anchor_text(line.text)[0]
        if _IMAGE_RE.fullmatch(image_source):
            blocks.append(
                _Block(
                    "image",
                    line.source_start,
                    line.end,
                    index,
                    index,
                    line.text,
                    {
                        "last_line_start": line.start,
                        "container_indent": len(line.text) - len(line.text.lstrip()),
                        "resource_less_carrier": is_resource_less_image_carrier_v3(image_source),
                    },
                )
            )
            index += 1
            continue

        end_index = index
        while end_index + 1 < len(lines) and lines[end_index + 1].text.strip():
            candidate = lines[end_index + 1].text
            if (
                _HEADING_RE.fullmatch(candidate)
                or _CAPTION_RE.fullmatch(candidate)
                or _match_fence_opener_v3(candidate)
                or _ANCHOR_ONLY_CANDIDATE_RE.fullmatch(candidate)
                or _QUOTE_PREFIX_RE.match(candidate) is not None
                or candidate.lstrip().startswith("$$")
                or _LIST_ITEM_RE.match(candidate)
            ):
                break
            end_index += 1
        blocks.append(
            _Block(
                paragraph_kind,
                line.source_start,
                lines[end_index].end,
                index,
                end_index,
                "\n".join(item.text for item in lines[index : end_index + 1]),
                {
                    "last_line_start": lines[end_index].start,
                    "container_indent": len(line.text) - len(line.text.lstrip()),
                },
            )
        )
        index = end_index + 1
    for block in blocks:
        block.data["container_path"] = container_path
        block.data["container_segments"] = container_segments

    nested: list[_Block] = []
    for block in blocks:
        if block.kind in {"block_quote", "callout"}:
            child_path = (*container_path, (block.kind, block.start))
            child_segments = (*container_segments, (block.kind, block.start, block.end))
            quote_lines = [_strip_quote_prefix(item) for item in lines[block.first_line : block.last_line + 1]]
            nested.extend(
                _scan_container_blocks(
                    quote_lines,
                    container_path=child_path,
                    container_segments=child_segments,
                    paragraph_kind="container_text",
                )
            )
        elif block.kind == "list":
            list_path = (*container_path, ("list", block.start))
            list_segments = (*container_segments, ("list", block.start, block.end))
            for item_lines in _split_list_items(lines[block.first_line : block.last_line + 1]):
                if not item_lines:
                    continue
                item_start = item_lines[0].source_start
                item_end = item_lines[-1].end
                item_path = (*list_path, ("list_item", item_start))
                item_segments = (*list_segments, ("list_item", item_start, item_end))
                nested.extend(
                    _scan_container_blocks(
                        item_lines,
                        container_path=item_path,
                        container_segments=item_segments,
                        paragraph_kind="list_item",
                    )
                )
    return [*blocks, *nested]


def _strip_quote_prefix(line: _Line) -> _Line:
    match = _QUOTE_PREFIX_RE.match(line.text)
    if match is None:
        return line
    consumed = match.end()
    return _Line(
        text=line.text[consumed:],
        start=line.start + consumed,
        content_end=line.content_end,
        end=line.end,
        number=line.number,
        source_start=line.source_start,
    )


def _split_list_items(lines: Sequence[_Line]) -> list[list[_Line]]:
    if not lines:
        return []
    first_match = _LIST_ITEM_RE.match(lines[0].text)
    if first_match is None:
        return []
    base_indent = len(first_match.group("indent"))
    groups: list[list[_Line]] = []
    current: list[_Line] = []
    marker_width = first_match.end()
    for line in lines:
        match = _LIST_ITEM_RE.match(line.text)
        is_sibling = match is not None and len(match.group("indent")) == base_indent
        if is_sibling:
            if current:
                groups.append(current)
            current = []
            assert match is not None
            marker_width = match.end()
            consumed = marker_width
        elif not line.text.strip():
            consumed = len(line.text)
        else:
            consumed = _continuation_prefix_length(line.text, marker_width)
        current.append(
            _Line(
                text=line.text[consumed:],
                start=line.start + consumed,
                content_end=line.content_end,
                end=line.end,
                number=line.number,
                source_start=line.source_start,
            )
        )
    if current:
        groups.append(current)
    return groups


def _continuation_prefix_length(text: str, width: int) -> int:
    consumed = 0
    columns = 0
    while consumed < len(text) and columns < width and text[consumed] in " \t":
        if text[consumed] == "\t":
            columns += 4 - (columns % 4)
        else:
            columns += 1
        consumed += 1
    return consumed if columns >= width else 0


def _extract_inline_anchor(text: str, *, absolute_start: int) -> tuple[str, tuple[str, SourceRange] | None]:
    match = _INLINE_ANCHOR_CANDIDATE_RE.search(text)
    if match is None:
        return text, None
    token_range = SourceRange(absolute_start + match.start("token"), absolute_start + match.end("token"))
    return text[: match.start("space")], (match.group("token"), token_range)


def _is_table_delimiter(text: str) -> bool:
    stripped = text.strip()
    if "|" not in stripped:
        return False
    cells = stripped.strip("|").split("|")
    return bool(cells) and all(_TABLE_DELIMITER_CELL_RE.fullmatch(cell.strip()) is not None for cell in cells)


def _strip_valid_inline_anchor_text(text: str) -> tuple[str, str | None]:
    match = _INLINE_ANCHOR_CANDIDATE_RE.search(text)
    if match is None or not _valid_id(match.group("token")[1:]):
        return text, None
    return text[: match.start("space")], match.group("token")


def _block_inline_anchor(block: _Block) -> tuple[str, SourceRange] | None:
    last_line = block.text.split("\n")[-1]
    match = _INLINE_ANCHOR_CANDIDATE_RE.search(last_line)
    if match is None:
        return None
    absolute_start = int(block.data.get("last_line_start", block.start))
    return match.group("token"), SourceRange(absolute_start + match.start("token"), absolute_start + match.end("token"))


def _ordinary_anchor_owner_range(block: _Block) -> SourceRange:
    """Return the complete structural range owned by one inline anchor."""

    segments = cast(
        tuple[tuple[str, int, int], ...],
        tuple(block.data.get("container_segments", ())),
    )
    if block.kind == "list_item" and segments and segments[-1][0] == "list_item":
        _kind, start, end = segments[-1]
        return SourceRange(start, end)
    return SourceRange(block.start, block.end)


def _container_path_projection(block: _Block) -> list[dict[str, Any]]:
    """Project strict CommonMark ancestors in frozen outer-to-inner order."""

    segments = cast(
        tuple[tuple[str, int, int], ...],
        tuple(block.data.get("container_segments", ())),
    )
    if block.kind == "list_item" and segments and segments[-1][0] == "list_item":
        segments = segments[:-1]
    return [
        {
            "block_kind": kind,
            "block_range": SourceRange(start, end).as_dict(),
        }
        for kind, start, end in segments
    ]


def _has_nested_inline_owner(blocks: Sequence[_Block], structural: _Block, token_range: SourceRange) -> bool:
    structural_path = tuple(structural.data["container_path"])
    for candidate in blocks:
        candidate_path = tuple(candidate.data["container_path"])
        if len(candidate_path) <= len(structural_path) or candidate_path[: len(structural_path)] != structural_path:
            continue
        if candidate.kind == "anchor_marker":
            candidate_range = SourceRange(
                int(candidate.data["token_start"]),
                int(candidate.data["token_end"]),
            )
            if candidate_range == token_range:
                return True
        inline = _block_inline_anchor(candidate)
        if inline is not None and inline[1] == token_range:
            return True
    return False


def _next_non_marker_block(
    blocks: Sequence[_Block],
    start: int,
    *,
    container_path: tuple[tuple[str, int], ...],
) -> int | None:
    for index in range(start, len(blocks)):
        if tuple(blocks[index].data["container_path"]) != container_path:
            continue
        if blocks[index].kind == "blank":
            continue
        if blocks[index].kind == "anchor_marker":
            return index
        return index
    return None


def _previous_non_marker_block(
    blocks: Sequence[_Block],
    start: int,
    *,
    container_path: tuple[tuple[str, int], ...],
) -> int | None:
    for index in range(start, -1, -1):
        if tuple(blocks[index].data["container_path"]) != container_path:
            continue
        if blocks[index].kind == "blank":
            continue
        return index
    return None


def _register_owner(owners: dict[str, list[_Owner]], anchor_id: str, owner: _Owner) -> None:
    owners.setdefault(anchor_id, []).append(owner)


def _source_identity(input_id: str, source_sha256: str) -> dict[str, Any]:
    return {
        "input_id": input_id,
        "sha256": source_sha256,
        "encoding": "utf-8",
        "coordinate_system": "unicode_code_point",
        "offset_base": 0,
        "range_end": "exclusive",
    }


def _diagnostic(
    source_identity: Mapping[str, Any],
    severity: Literal["warning", "error"],
    code: str,
    message: str,
    source_range: SourceRange,
    *,
    related_ranges: Sequence[SourceRange] = (),
    fixes: Sequence[dict[str, Any]] = (),
) -> dict[str, Any]:
    diagnostic: dict[str, Any] = {
        "severity": severity,
        "code": code,
        "message": message,
        "evidence_schema": DIAGNOSTIC_EVIDENCE_SCHEMA,
        "source": dict(source_identity),
        "range": source_range.as_dict(),
    }
    if related_ranges:
        diagnostic["related_ranges"] = [item.as_dict() for item in related_ranges]
    if fixes:
        diagnostic["fixes"] = list(fixes)
    return diagnostic


def _invalid_anchor_diagnostic(
    source_identity: Mapping[str, Any],
    token_range: SourceRange,
    anchor_id: str,
    replacements: Mapping[int, str],
    owners: Mapping[str, Sequence[_Owner]],
) -> dict[str, Any]:
    fixes: tuple[dict[str, Any], ...] = ()
    replacement = replacements.get(token_range.start)
    if replacement is not None and _valid_id(replacement) and replacement not in owners:
        fixes = (
            {
                "fix_id": "docwen.markdown.fix.rename_anchor",
                "edits": [{"range": token_range.as_dict(), "replacement": f"^{replacement}"}],
            },
        )
    return _diagnostic(
        source_identity,
        "error",
        "docwen.markdown.anchor.invalid_id",
        f"Anchor ID {anchor_id!r} does not match [A-Za-z0-9-]{{1,128}}.",
        token_range,
        fixes=fixes,
    )


def _valid_id(value: str) -> bool:
    return _ID_RE.fullmatch(value) is not None


def _validate_external_references(
    records: Sequence[ExternalReferenceResolution],
) -> dict[tuple[str, str, str], ExternalReferenceResolution]:
    output: dict[tuple[str, str, str], ExternalReferenceResolution] = {}
    for record in records:
        if not record.page_locator or not record.resolved_document_id:
            raise ValueError("external reference records require non-empty locator and document identity")
        if re.fullmatch(r"[0-9a-f]{64}", record.resolved_document_sha256) is None:
            raise ValueError("external reference document SHA-256 must be lowercase 64-hex")
        if record.selector_kind == "stable_id":
            if record.target_id is None or not _valid_id(record.target_id) or record.heading_path:
                raise ValueError("external stable resolution requires only a valid target_id")
            selector = record.target_id
        else:
            if record.target_id is not None or not record.heading_path or any(not item for item in record.heading_path):
                raise ValueError("external soft resolution requires only a non-empty heading_path")
            if record.resolved_kind != "heading":
                raise ValueError("external soft resolution must resolve to Heading")
            selector = "#".join(record.heading_path)
        key = (record.page_locator, record.selector_kind, selector)
        if key in output:
            raise ValueError("external reference resolution keys must be unique")
        output[key] = record
    return output


def _validate_external_citations(
    records: Sequence[ExternalCitationResolution],
) -> dict[str, ExternalCitationResolution]:
    output: dict[str, ExternalCitationResolution] = {}
    for record in records:
        if _CITATION_KEY_RE.fullmatch(record.key) is None or not record.record_id or not record.presentation:
            raise ValueError("external citation records require a valid key, record identity, and presentation")
        if re.fullmatch(r"[0-9a-f]{64}", record.record_sha256) is None:
            raise ValueError("external citation record SHA-256 must be lowercase 64-hex")
        if record.key in output:
            raise ValueError("external citation keys must be unique")
        output[record.key] = record
    return output


def _parse_reference_body(body: str) -> tuple[str | None, str, str | None]:
    selector_text, separator, alias = body.partition("|")
    if separator and (not alias or "|" in alias):
        raise ValueError("semantic reference Alias must be one non-empty suffix")
    if "#" not in selector_text:
        raise ValueError("semantic reference requires an explicit # fragment")
    page, fragment = selector_text.split("#", 1)
    if not fragment:
        raise ValueError("semantic reference fragment must not be empty")
    return (page or None), fragment, (alias if separator else None)


def _resolve_reference(
    source_identity: Mapping[str, Any],
    match: re.Match[str],
    token_range: SourceRange,
    targets: Sequence[dict[str, Any]],
    headings: Sequence[dict[str, Any]],
    owners: Mapping[str, Sequence[_Owner]],
    external: Mapping[tuple[str, str, str], ExternalReferenceResolution],
) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    diagnostics: list[dict[str, Any]] = []
    body = match.group("body")
    try:
        page_locator, fragment, alias = _parse_reference_body(body)
    except ValueError:
        record = {
            "selector_kind": "heading_path",
            "heading_path": [body],
            "raw": match.group(0),
            "range": token_range.as_dict(),
            "resolution_status": "missing",
        }
        diagnostics.append(
            _diagnostic(
                source_identity,
                "error",
                "docwen.markdown.cross_reference.missing",
                "The semantic cross-reference has no valid target selector.",
                token_range,
            )
        )
        return record, diagnostics

    record: dict[str, Any] = {
        "selector_kind": "stable_id" if fragment.startswith("^") else "heading_path",
        "raw": match.group(0),
        "range": token_range.as_dict(),
    }
    if page_locator is not None:
        record["page_locator"] = page_locator
    if alias is not None:
        record["alias"] = alias
        alias_start = match.start() + match.group(0).rfind(alias)
        record["alias_range"] = SourceRange(alias_start, alias_start + len(alias)).as_dict()

    current_title: str | None = None
    if fragment.startswith("^"):
        target_id = fragment[1:]
        record["target_id"] = target_id
        selector_key = target_id
        if not _valid_id(target_id):
            status: ResolutionStatus = "missing"
        elif page_locator is not None:
            status = "external_unresolved"
        else:
            matching_targets = [item for item in targets if item.get("id") == target_id]
            owner_records = owners.get(target_id, ())
            if len(owner_records) > 1:
                status = "ambiguous"
            elif matching_targets:
                target = matching_targets[0]
                status = "resolved" if target.get("number") else "unnumbered"
                record["resolved_kind"] = target["kind"]
                current_title = str(target["title"])
                if target.get("number"):
                    record["cached_number"] = target["number"]
            elif owner_records and owner_records[0].owner_kind == "ordinary_anchor":
                status = "non_semantic"
            else:
                status = "missing"
    else:
        heading_path = tuple(fragment.split("#"))
        record["heading_path"] = list(heading_path)
        selector_key = "#".join(heading_path)
        if any(not item for item in heading_path):
            status = "missing"
        elif page_locator is not None:
            status = "external_unresolved"
        else:
            matching_headings = [
                heading for heading in headings if tuple(heading["heading_path"][-len(heading_path) :]) == heading_path
            ]
            if len(matching_headings) == 1:
                target = matching_headings[0]
                status = "resolved" if target.get("number") else "unnumbered"
                record["resolved_kind"] = "heading"
                current_title = str(target["title"])
                if target.get("number"):
                    record["cached_number"] = target["number"]
                if target.get("id"):
                    record["resolved_target_id"] = target["id"]
            elif not matching_headings:
                status = "missing"
            else:
                status = "ambiguous"
                diagnostics.append(
                    _diagnostic(
                        source_identity,
                        "error",
                        "docwen.markdown.cross_reference.ambiguous",
                        "The same-document Heading path selects more than one Heading.",
                        token_range,
                        related_ranges=tuple(
                            SourceRange(item["range"]["start"], item["range"]["end"]) for item in matching_headings[:16]
                        ),
                    )
                )

    if page_locator is not None:
        external_record = external.get((page_locator, record["selector_kind"], selector_key))
        if external_record is not None:
            status = "resolved" if external_record.cached_number else "unnumbered"
            record["resolved_document_id"] = external_record.resolved_document_id
            record["resolved_document_sha256"] = external_record.resolved_document_sha256
            record["resolved_kind"] = external_record.resolved_kind
            current_title = external_record.current_title
            if external_record.cached_number:
                record["cached_number"] = external_record.cached_number

    record["resolution_status"] = status
    if status == "missing":
        diagnostics.append(
            _diagnostic(
                source_identity,
                "error",
                "docwen.markdown.cross_reference.missing",
                "The semantic cross-reference target does not exist in the supplied resolution boundary.",
                token_range,
            )
        )
    elif status == "non_semantic":
        diagnostics.append(
            _diagnostic(
                source_identity,
                "error",
                "docwen.markdown.cross_reference.non_semantic_target",
                "The referenced ID belongs only to an ordinary navigation anchor.",
                token_range,
            )
        )
    elif status == "unnumbered":
        diagnostics.append(
            _diagnostic(
                source_identity,
                "error",
                "docwen.markdown.cross_reference.unnumbered_target",
                "The resolved semantic target has no materializable number.",
                token_range,
            )
        )

    if alias is not None and current_title is not None and alias != current_title:
        alias_range_dict = record["alias_range"]
        diagnostics.append(
            _diagnostic(
                source_identity,
                "warning",
                "docwen.markdown.cross_reference.alias_stale",
                "The authored Alias differs from the current target title or caption.",
                SourceRange(alias_range_dict["start"], alias_range_dict["end"]),
            )
        )
    return record, diagnostics


def _project_wikilink(match: re.Match[str], token_range: SourceRange) -> dict[str, Any]:
    body = match.group("body")
    destination, separator, alias = body.partition("|")
    page_locator = destination
    fragment: str | None = None
    if "#" in destination:
        page_locator, fragment = destination.split("#", 1)
    record: dict[str, Any] = {
        "kind": "embed" if match.group("embed") else "link",
        "raw": match.group(0),
        "range": token_range.as_dict(),
        "page_locator": page_locator,
    }
    if fragment is not None:
        if fragment.startswith("^"):
            record["fragment_kind"] = "stable_id"
            record["target_id"] = fragment[1:]
        else:
            record["fragment_kind"] = "heading_path"
            record["heading_path"] = fragment.split("#")
    if separator:
        record["alias"] = alias
    return record


def _citation_item(
    key: str,
    key_range: SourceRange,
    external: Mapping[str, ExternalCitationResolution],
) -> dict[str, Any]:
    record: dict[str, Any] = {"key": key, "key_range": key_range.as_dict(), "resolution_status": "unresolved"}
    resolved = external.get(key)
    if resolved is not None:
        record.update(
            {
                "resolution_status": "resolved",
                "resolved_record_id": resolved.record_id,
                "resolved_record_sha256": resolved.record_sha256,
                "presentation": resolved.presentation,
            }
        )
    return record


def _literal_shield_ranges(source: str, blocks: Sequence[_Block]) -> list[SourceRange]:
    ranges = [SourceRange(block.start, block.end) for block in blocks if block.kind in {"code_block", "fenced_block"}]
    for match in re.finditer(r"(`+)(?:(?!\1).)*\1", source):
        ranges.append(SourceRange(match.start(), match.end()))
    for match in re.finditer(r"https?://[^\s<]+|<[^>\r\n]*>", source):
        ranges.append(SourceRange(match.start(), match.end()))
    return _merge_ranges(ranges)


def _merge_ranges(ranges: Iterable[SourceRange]) -> list[SourceRange]:
    ordered = sorted(ranges, key=lambda item: (item.start, item.end))
    merged: list[SourceRange] = []
    for item in ordered:
        if merged and item.start <= merged[-1].end:
            merged[-1] = SourceRange(merged[-1].start, max(merged[-1].end, item.end))
        else:
            merged.append(item)
    return merged


def _overlaps_any(candidate: SourceRange, ranges: Sequence[SourceRange]) -> bool:
    return any(candidate.start < item.end and item.start < candidate.end for item in ranges)


__all__ = [
    "DIAGNOSTICS_SCHEMA",
    "DIAGNOSTICS_SCHEMA_ID",
    "DIAGNOSTIC_EVIDENCE_SCHEMA",
    "SEMANTICS_SCHEMA",
    "SEMANTICS_SCHEMA_ID",
    "ExternalCitationResolution",
    "ExternalReferenceResolution",
    "MarkdownSemanticsV3Analysis",
    "SourceRange",
    "analyze_markdown_semantics_v3",
    "is_resource_less_image_carrier_v3",
    "markdown_semantics_body_start_v3",
    "render_caption_declaration",
    "render_citation_token",
    "render_cross_reference_token",
    "select_safe_fence",
]

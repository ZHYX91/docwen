"""Range-authenticated AST projection for the v4 resolved-document port.

This adapter deliberately does not parse numbering or resolve Markdown
identities.  Core has already authenticated the authored source and every
resolved occurrence.  The adapter replaces only those exact occurrences (or
adds a zero-width target marker at a closed structural location), lets the
ordinary Mistune parser run, and then proves which AST node owns each marker.

Resource replacements are intentionally outside this module.  A later
request-private composer must merge Core-authenticated resource edits with the
``marker_edits`` exposed by :class:`ResolvedRuntimeV4Plan` before parsing.
"""

from __future__ import annotations

import hashlib
import re
from dataclasses import dataclass, replace
from typing import Any, Literal

from docwen_core.models.resolved_numbering import (
    RESOLVED_DOCUMENT_SCHEMA,
    ResolvedCitation,
    ResolvedDocumentTarget,
    ResolvedNumberingPort,
    ResolvedReference,
)
from docwen_plugin_markdown._resolved_runtime_v4_evidence import (
    authored_inline_text,
    prove_resolved_occurrence_order,
)

type RuntimeMarkerRoleV4 = Literal["target", "reference", "citation"]
type RuntimeMarkerPayloadV4 = ResolvedDocumentTarget | ResolvedReference | ResolvedCitation

_TARGET_NODE = "_docwen_resolved_v4_target_marker"
_TARGET_KEY = "_docwen_resolved_v4_target"
_CAPTION_CHILDREN_KEY = "_docwen_resolved_v4_caption_children"
_INLINE_IMAGE_CHILDREN_KEY = "_docwen_resolved_v4_inline_image_children"
_INLINE_IMAGE_BEFORE_KEY = "_docwen_resolved_v4_inline_image_before"
_CAPTION_DIRECTIONS_KEY = "_docwen_resolved_v4_caption_directions"
_REFERENCE_KEY = "_docwen_resolved_v4_reference"
_CITATION_KEY = "_docwen_resolved_v4_citation"
_RESERVED_KEYS = frozenset(
    {
        _TARGET_KEY,
        _CAPTION_CHILDREN_KEY,
        _INLINE_IMAGE_CHILDREN_KEY,
        _INLINE_IMAGE_BEFORE_KEY,
        _CAPTION_DIRECTIONS_KEY,
        _REFERENCE_KEY,
        _CITATION_KEY,
    }
)

_ATX_HEADING_RE = re.compile(r"^(?P<indent> {0,3})(?P<marks>#{1,9})(?!#)[ \t]+(?P<body>.*?)[ \t]*$")
_CAPTION_RE = re.compile(r"^(?P<keyword>Figure|Table|Equation|Code):(?P<body>.*)$", re.IGNORECASE)
_TRAILING_ID_RE = re.compile(r"(?P<space>[ \t]+)(?P<token>\^(?P<id>[^\s]+))[ \t]*$")
_CLOSING_ATX_RE = re.compile(r"[ \t]+#+$")
_HISTORICAL_ATTRIBUTE_RE = re.compile(r"\{#[^{}\s]+\}[ \t]*$")
_QUOTE_PREFIX_RE = re.compile(r" {0,3}>[ \t]?")
_LIST_PREFIX_RE = re.compile(r"[ \t]*(?:[-+*]|\d{1,9}[.)])[ \t]+")
_ORDINARY_ID_TEXT_RE = re.compile(r"\^[A-Za-z0-9-]{1,128}")
_STANDALONE_ID_RE = re.compile(r"(?P<token>\^(?P<id>[^\s]+))[ \t]*$")
_KIND_BY_KEYWORD = {
    "figure": "figure",
    "table": "table",
    "equation": "equation",
    "code": "code_block",
}


class ResolvedRuntimeV4Unsupported(ValueError):
    """The authenticated snapshot cannot be bound to the ordinary AST."""


@dataclass(frozen=True, slots=True)
class RuntimeMarkerEditV4:
    """One original-source edit that projects a typed occurrence to a marker."""

    source_start: int
    source_end: int
    original: str
    replacement: str
    role: RuntimeMarkerRoleV4


@dataclass(frozen=True, slots=True)
class RuntimeMarkerV4:
    """One inert source marker and its already-validated typed payload."""

    marker: str
    role: RuntimeMarkerRoleV4
    payload: RuntimeMarkerPayloadV4
    edit: RuntimeMarkerEditV4
    trim_preceding_space: bool = False
    standalone_target_id: bool = False
    caption_carrier_directions: tuple[Literal[-1, 1], ...] = ()


@dataclass(frozen=True, slots=True)
class ResolvedRuntimeV4Plan:
    """Immutable marker projection for one exact resolved-numbering port."""

    port: ResolvedNumberingPort
    marker_prefix: str
    markers: tuple[RuntimeMarkerV4, ...]
    marker_edits: tuple[RuntimeMarkerEditV4, ...]
    shielded_source: str

    @property
    def source_sha256(self) -> str:
        return self.port.source_sha256


@dataclass(frozen=True, slots=True)
class _MarkerSpec:
    source_start: int
    source_end: int
    original: str
    role: RuntimeMarkerRoleV4
    payload: RuntimeMarkerPayloadV4
    trim_preceding_space: bool = False
    standalone_target_id: bool = False
    caption_carrier_directions: tuple[Literal[-1, 1], ...] = ()


def prepare_resolved_runtime_v4(port: ResolvedNumberingPort) -> ResolvedRuntimeV4Plan:
    """Create inert marker edits without invoking any semantic/numbering parser."""

    source = port.document.authored_markdown
    _validate_port_snapshot(port)
    specs: list[_MarkerSpec] = [_target_marker_spec(source, target) for target in port.document.targets]
    specs.extend(_inline_marker_spec(source, item, "reference") for item in port.document.references)
    specs.extend(_inline_marker_spec(source, item, "citation") for item in port.document.citations)
    specs.sort(key=_spec_sort_key)
    _validate_marker_specs(specs)

    marker_prefix = _unique_marker_prefix(source, port.source_sha256)
    markers: list[RuntimeMarkerV4] = []
    for index, spec in enumerate(specs):
        token = f"{marker_prefix}{index:08d}X"
        edit = RuntimeMarkerEditV4(
            source_start=spec.source_start,
            source_end=spec.source_end,
            original=spec.original,
            replacement=token,
            role=spec.role,
        )
        markers.append(
            RuntimeMarkerV4(
                marker=token,
                role=spec.role,
                payload=spec.payload,
                edit=edit,
                trim_preceding_space=spec.trim_preceding_space,
                standalone_target_id=spec.standalone_target_id,
                caption_carrier_directions=spec.caption_carrier_directions,
            )
        )
    marker_tuple = tuple(markers)
    edits = tuple(item.edit for item in marker_tuple)
    return ResolvedRuntimeV4Plan(
        port=port,
        marker_prefix=marker_prefix,
        markers=marker_tuple,
        marker_edits=edits,
        shielded_source=_apply_marker_edits(source, marker_tuple),
    )


def apply_resolved_runtime_v4(
    ast_nodes: list[dict[str, Any]],
    plan: ResolvedRuntimeV4Plan,
) -> list[dict[str, Any]]:
    """Restore typed occurrences and prove every target's structural AST owner."""

    marker_map = {item.marker: item for item in plan.markers}
    found: set[str] = set()
    restored = [
        _restore_node(node, marker_map=marker_map, marker_prefix=plan.marker_prefix, found=found) for node in ast_nodes
    ]
    restored = _attach_standalone_target_markers(restored)
    bound_targets: set[tuple[int, int, str]] = set()
    for node in restored:
        _bind_target_markers(node, bound_targets)
    restored = _bind_caption_targets(restored)

    expected_markers = set(marker_map)
    if found != expected_markers:
        raise ResolvedRuntimeV4Unsupported("not every v4 source marker was restored exactly once")
    expected_targets = {item.occurrence_key for item in plan.port.document.targets}
    if bound_targets != expected_targets:
        raise ResolvedRuntimeV4Unsupported("not every v4 target marker has one structural AST owner")
    if any(True for _item in _iter_unbound_target_markers(restored)):
        raise ResolvedRuntimeV4Unsupported("a v4 target marker remained outside a supported structural owner")
    try:
        prove_resolved_occurrence_order(restored, plan.port)
    except ValueError as exc:
        raise ResolvedRuntimeV4Unsupported(str(exc)) from exc
    return restored


def _validate_port_snapshot(port: ResolvedNumberingPort) -> None:
    source = port.document.authored_markdown
    source_sha256 = hashlib.sha256(source.encode("utf-8")).hexdigest()
    if source_sha256 != port.source_sha256:
        raise ResolvedRuntimeV4Unsupported("authored Markdown no longer matches the resolved source hash")
    if (
        port.document_envelope.input_id != port.plan_envelope.input_id
        or port.document_envelope.source_sha256 != port.plan_envelope.source_sha256
        or port.document_envelope.plan_sha256 != port.plan_envelope.plan_sha256
    ):
        raise ResolvedRuntimeV4Unsupported("resolved document and numbering plan pointers differ")

    document_targets = {item.occurrence_key: item for item in port.document.targets}
    plan_targets = {item.occurrence_key: item for item in port.plan.targets}
    if document_targets.keys() != plan_targets.keys():
        raise ResolvedRuntimeV4Unsupported("resolved document and numbering plan target inventories differ")
    for key, target in document_targets.items():
        plan_target = plan_targets[key]
        if target.target_id != plan_target.target_id:
            raise ResolvedRuntimeV4Unsupported("resolved target identity differs from the numbering plan")
        _require_authenticated_slice(
            source,
            target.source_start,
            target.source_end,
            target.source_slice_sha256,
        )
    for item in (*port.document.references, *port.document.citations):
        source_slice = _require_authenticated_slice(
            source,
            item.source_start,
            item.source_end,
            item.source_slice_sha256,
        )
        if source_slice != item.authored_token:
            raise ResolvedRuntimeV4Unsupported("resolved inline token differs from its authenticated source slice")


def _require_authenticated_slice(
    source: str,
    source_start: int,
    source_end: int,
    expected_sha256: str,
) -> str:
    if source_start < 0 or source_end <= source_start or source_end > len(source):
        raise ResolvedRuntimeV4Unsupported("resolved source range is empty or outside authored Markdown")
    source_slice = source[source_start:source_end]
    if hashlib.sha256(source_slice.encode("utf-8")).hexdigest() != expected_sha256:
        raise ResolvedRuntimeV4Unsupported("resolved source range hash changed after admission")
    return source_slice


def _target_marker_spec(source: str, target: ResolvedDocumentTarget) -> _MarkerSpec:
    if target.kind == "heading":
        target_lines = _target_lines(source, target)
        line, line_start = target_lines[0]
        content_offset = _commonmark_container_prefix_length(line)
        content = line[content_offset:]
        standalone_id = _standalone_target_id(source, line_start, target)
        return _heading_marker_spec(content, line_start + content_offset, target, standalone_id=standalone_id)
    target_lines = _target_lines(source, target)
    declarations: list[tuple[int, str, int]] = []
    for index, (line, line_start) in enumerate(target_lines):
        content_offset = _commonmark_container_prefix_length(line)
        content = line[content_offset:]
        match = _CAPTION_RE.fullmatch(content)
        if match is None or _KIND_BY_KEYWORD[match.group("keyword").casefold()] != target.kind:
            continue
        declarations.append((index, content, line_start + content_offset))
    if len(declarations) != 1:
        raise ResolvedRuntimeV4Unsupported("resolved caption range does not contain one exact declaration kind")
    declaration_index, content, absolute_start = declarations[0]
    standalone_id = _standalone_target_id(source, target_lines[declaration_index][1], target)
    spec = _caption_marker_spec(content, absolute_start, target, standalone_id=standalone_id)
    directions = _caption_source_directions(source, target, spec.source_start, standalone_id=standalone_id)
    if not directions:
        raise ResolvedRuntimeV4Unsupported("resolved caption has no object within one blank source line")
    return replace(spec, caption_carrier_directions=directions)


def _first_target_line(source: str, target: ResolvedDocumentTarget) -> tuple[str, int]:
    if target.source_start > 0 and source[target.source_start - 1] != "\n":
        raise ResolvedRuntimeV4Unsupported("resolved target does not start at a source line boundary")
    line_end = source.find("\n", target.source_start, target.source_end)
    if line_end < 0:
        line_end = target.source_end
    if line_end > target.source_start and source[line_end - 1] == "\r":
        line_end -= 1
    return source[target.source_start : line_end], target.source_start


def _target_lines(source: str, target: ResolvedDocumentTarget) -> tuple[tuple[str, int], ...]:
    if target.source_start > 0 and source[target.source_start - 1] != "\n":
        raise ResolvedRuntimeV4Unsupported("resolved target does not start at a source line boundary")
    output: list[tuple[str, int]] = []
    cursor = target.source_start
    while cursor < target.source_end:
        line_end = source.find("\n", cursor, target.source_end)
        physical_end = target.source_end if line_end < 0 else line_end
        content_end = physical_end - 1 if physical_end > cursor and source[physical_end - 1] == "\r" else physical_end
        output.append((source[cursor:content_end], cursor))
        if line_end < 0:
            break
        cursor = line_end + 1
    return tuple(output)


def _standalone_target_id(
    source: str,
    owner_line_start: int,
    target: ResolvedDocumentTarget,
) -> tuple[int, int, str] | None:
    owner_line_end = source.find("\n", owner_line_start)
    if owner_line_end < 0:
        return None
    line_start = owner_line_end + 1
    line_end = source.find("\n", line_start)
    if line_end < 0:
        line_end = len(source)
    content_end = line_end - 1 if line_end > line_start and source[line_end - 1] == "\r" else line_end
    line = source[line_start:content_end]
    content_offset = _commonmark_container_prefix_length(line)
    content = line[content_offset:]
    match = _STANDALONE_ID_RE.fullmatch(content)
    if match is None:
        return None
    if target.target_id is None or match.group("id") != target.target_id:
        raise ResolvedRuntimeV4Unsupported("resolved target standalone ID differs from its typed identity")
    token_start = line_start + content_offset + match.start("token")
    return token_start, token_start + len(match.group("token")), match.group("token")


def _caption_source_directions(
    source: str,
    target: ResolvedDocumentTarget,
    declaration_marker_start: int,
    *,
    standalone_id: tuple[int, int, str] | None,
) -> tuple[Literal[-1, 1], ...]:
    """Return source sides reachable through at most one blank line."""

    lines: list[tuple[str, int]] = []
    cursor = 0
    while cursor < len(source):
        line_end = source.find("\n", cursor)
        physical_end = len(source) if line_end < 0 else line_end
        content_end = physical_end - 1 if physical_end > cursor and source[physical_end - 1] == "\r" else physical_end
        lines.append((source[cursor:content_end], cursor))
        if line_end < 0:
            break
        cursor = line_end + 1
    declaration_index = next(
        (
            index
            for index, (_line, start) in enumerate(lines)
            if start <= declaration_marker_start < start + len(_line) + 1
        ),
        None,
    )
    if declaration_index is None:
        raise ResolvedRuntimeV4Unsupported("resolved caption declaration is outside its target range")

    def adjacent(direction: Literal[-1, 1]) -> bool:
        blanks = 0
        skipped_target_id = False
        stop = len(lines) if direction > 0 else -1
        for index in range(declaration_index + direction, stop, direction):
            line = lines[index][0]
            content = line[_commonmark_container_prefix_length(line) :]
            line_start = lines[index][1]
            if (
                not skipped_target_id
                and standalone_id is not None
                and line_start <= standalone_id[0] < line_start + len(line) + 1
            ):
                skipped_target_id = True
                continue
            if not content.strip():
                blanks += 1
                if blanks > 1:
                    return False
                continue
            return True
        return False

    return tuple(direction for direction in (-1, 1) if adjacent(direction))


def _commonmark_container_prefix_length(line: str) -> int:
    cursor = 0
    while cursor < len(line):
        quote = _QUOTE_PREFIX_RE.match(line, cursor)
        if quote is not None:
            cursor = quote.end()
            continue
        list_item = _LIST_PREFIX_RE.match(line, cursor)
        if list_item is not None:
            cursor = list_item.end()
            continue
        break
    return cursor


def _heading_marker_spec(
    content: str,
    absolute_start: int,
    target: ResolvedDocumentTarget,
    *,
    standalone_id: tuple[int, int, str] | None,
) -> _MarkerSpec:
    match = _ATX_HEADING_RE.fullmatch(content)
    if match is None or len(match.group("marks")) != target.heading_level:
        raise ResolvedRuntimeV4Unsupported("resolved Heading range is not its exact ATX level")
    body = match.group("body")
    body_start = absolute_start + match.start("body")
    trailing_id = _TRAILING_ID_RE.search(body)
    if target.target_id is not None:
        if trailing_id is None and standalone_id is not None:
            source_start, source_end, original = standalone_id
            return _MarkerSpec(
                source_start=source_start,
                source_end=source_end,
                original=original,
                role="target",
                payload=target,
                standalone_target_id=True,
            )
        if trailing_id is None or trailing_id.group("id") != target.target_id or standalone_id is not None:
            raise ResolvedRuntimeV4Unsupported("resolved Heading target ID is not its exact trailing source token")
        return _MarkerSpec(
            source_start=body_start + trailing_id.start("token"),
            source_end=body_start + trailing_id.end("token"),
            original=trailing_id.group("token"),
            role="target",
            payload=target,
            trim_preceding_space=True,
        )
    if trailing_id is not None:
        raise ResolvedRuntimeV4Unsupported("ID-less Heading range contains a trailing source ID token")
    closing = _CLOSING_ATX_RE.search(body)
    insertion = closing.start() if closing is not None else len(body)
    return _MarkerSpec(
        source_start=body_start + insertion,
        source_end=body_start + insertion,
        original="",
        role="target",
        payload=target,
    )


def _caption_marker_spec(
    content: str,
    absolute_start: int,
    target: ResolvedDocumentTarget,
    *,
    standalone_id: tuple[int, int, str] | None,
) -> _MarkerSpec:
    match = _CAPTION_RE.fullmatch(content)
    if match is None or _KIND_BY_KEYWORD[match.group("keyword").casefold()] != target.kind:
        raise ResolvedRuntimeV4Unsupported("resolved caption range is not its exact declaration kind")
    body = match.group("body")
    body_start = absolute_start + match.start("body")
    trailing_id = _TRAILING_ID_RE.search(body)
    if _HISTORICAL_ATTRIBUTE_RE.search(body) is not None:
        raise ResolvedRuntimeV4Unsupported("historical caption attributes are not current v4 declarations")
    if target.target_id is not None:
        if trailing_id is None and standalone_id is not None:
            authored_source = body.strip(" \t")
            _validate_caption_content(target, authored_source)
            source_start, source_end, original = standalone_id
            return _MarkerSpec(
                source_start=source_start,
                source_end=source_end,
                original=original,
                role="target",
                payload=target,
                standalone_target_id=True,
            )
        if trailing_id is None or trailing_id.group("id") != target.target_id or standalone_id is not None:
            raise ResolvedRuntimeV4Unsupported("resolved caption target ID is not its exact trailing source token")
        authored_source = body[: trailing_id.start("space")].strip(" \t")
        _validate_caption_content(target, authored_source)
        return _MarkerSpec(
            source_start=body_start + trailing_id.start("token"),
            source_end=body_start + trailing_id.end("token"),
            original=trailing_id.group("token"),
            role="target",
            payload=target,
            trim_preceding_space=True,
        )
    if trailing_id is not None:
        raise ResolvedRuntimeV4Unsupported("ID-less caption range contains a trailing source ID token")
    _validate_caption_content(target, body.strip(" \t"))
    colon_end = absolute_start + match.start("body")
    return _MarkerSpec(
        source_start=colon_end,
        source_end=colon_end,
        original="",
        role="target",
        payload=target,
    )


def _validate_caption_content(target: ResolvedDocumentTarget, authored_source: str) -> None:
    if target.kind in {"figure", "table"} and not authored_source:
        raise ResolvedRuntimeV4Unsupported("Figure and Table captions require authored content")
    if target.kind in {"equation", "code_block"} and not authored_source and target.target_id is None:
        raise ResolvedRuntimeV4Unsupported("an empty Equation or Code caption requires an explicit source ID")


def _inline_marker_spec(
    source: str,
    item: ResolvedReference | ResolvedCitation,
    role: Literal["reference", "citation"],
) -> _MarkerSpec:
    original = source[item.source_start : item.source_end]
    if original != item.authored_token:
        raise ResolvedRuntimeV4Unsupported("resolved inline marker no longer matches authored Markdown")
    return _MarkerSpec(
        source_start=item.source_start,
        source_end=item.source_end,
        original=original,
        role=role,
        payload=item,
    )


def _spec_sort_key(spec: _MarkerSpec) -> tuple[int, int, str, str]:
    payload_kind = getattr(spec.payload, "kind", type(spec.payload).__name__)
    return (spec.source_start, spec.source_end, spec.role, str(payload_kind))


def _validate_marker_specs(specs: list[_MarkerSpec]) -> None:
    insertion_points: set[int] = set()
    active_end = -1
    for spec in specs:
        if spec.source_start < 0 or spec.source_end < spec.source_start:
            raise ResolvedRuntimeV4Unsupported("v4 marker edit has an invalid source range")
        if spec.source_start < active_end:
            raise ResolvedRuntimeV4Unsupported("v4 marker edits overlap")
        if spec.source_start == spec.source_end:
            if spec.source_start in insertion_points:
                raise ResolvedRuntimeV4Unsupported("v4 marker insertion points collide")
            insertion_points.add(spec.source_start)
        else:
            active_end = spec.source_end


def _unique_marker_prefix(source: str, source_sha256: str) -> str:
    for salt in range(256):
        prefix = f"DOCWENRESOLVEDV4X{source_sha256[:16].upper()}X{salt:02X}X"
        if prefix not in source:
            return prefix
    raise ResolvedRuntimeV4Unsupported("authored source exhausts the v4 marker namespace")


def _apply_marker_edits(source: str, markers: tuple[RuntimeMarkerV4, ...]) -> str:
    for marker in markers:
        edit = marker.edit
        if source[edit.source_start : edit.source_end] != edit.original:
            raise ResolvedRuntimeV4Unsupported("v4 marker edit original changed after preparation")
    projected = source
    for marker in reversed(markers):
        edit = marker.edit
        projected = projected[: edit.source_start] + edit.replacement + projected[edit.source_end :]
    return projected


def _restore_node(
    node: dict[str, Any],
    *,
    marker_map: dict[str, RuntimeMarkerV4],
    marker_prefix: str,
    found: set[str],
) -> dict[str, Any]:
    if _RESERVED_KEYS.intersection(node):
        raise ResolvedRuntimeV4Unsupported("input AST already contains reserved v4 binding metadata")
    output = dict(node)
    children = node.get("children")
    if isinstance(children, list):
        output["children"] = _restore_children(
            children,
            marker_map=marker_map,
            marker_prefix=marker_prefix,
            found=found,
        )
    return output


def _restore_children(
    children: list[dict[str, Any]],
    *,
    marker_map: dict[str, RuntimeMarkerV4],
    marker_prefix: str,
    found: set[str],
) -> list[dict[str, Any]]:
    output: list[dict[str, Any]] = []
    for child in children:
        restored = _restore_node(
            child,
            marker_map=marker_map,
            marker_prefix=marker_prefix,
            found=found,
        )
        if restored.get("type") != "text":
            output.append(restored)
            continue
        key = "raw" if "raw" in restored else "text"
        text = str(restored.get(key, ""))
        output.extend(
            _restore_marker_text(
                text,
                key=key,
                marker_map=marker_map,
                marker_prefix=marker_prefix,
                found=found,
            )
        )
    return output


def _restore_marker_text(
    text: str,
    *,
    key: str,
    marker_map: dict[str, RuntimeMarkerV4],
    marker_prefix: str,
    found: set[str],
) -> list[dict[str, Any]]:
    output: list[dict[str, Any]] = []
    cursor = 0
    marker_length = len(marker_prefix) + 9
    while cursor < len(text):
        marker_start = text.find(marker_prefix, cursor)
        if marker_start < 0:
            if cursor < len(text):
                output.append({"type": "text", key: text[cursor:]})
            break
        if marker_start > cursor:
            output.append({"type": "text", key: text[cursor:marker_start]})
        token = text[marker_start : marker_start + marker_length]
        marker = marker_map.get(token)
        if marker is None:
            raise ResolvedRuntimeV4Unsupported("AST contains an unknown or truncated v4 source marker")
        if token in found:
            raise ResolvedRuntimeV4Unsupported("AST contains a duplicated v4 source marker")
        found.add(token)
        output.append(_marker_node(marker))
        cursor = marker_start + marker_length
    return output


def _marker_node(marker: RuntimeMarkerV4) -> dict[str, Any]:
    payload = marker.payload
    if marker.role == "target":
        return {"type": _TARGET_NODE, "marker": marker}
    if marker.role == "reference":
        assert isinstance(payload, ResolvedReference)
        return {
            "type": "semantic_cross_reference",
            "schema": RESOLVED_DOCUMENT_SCHEMA,
            "raw": payload.authored_token,
            "source_start": payload.source_start,
            "source_end": payload.source_end,
            _REFERENCE_KEY: payload,
        }
    assert isinstance(payload, ResolvedCitation)
    return {
        "type": "semantic_citation",
        "schema": RESOLVED_DOCUMENT_SCHEMA,
        "raw": payload.authored_token,
        "source_start": payload.source_start,
        "source_end": payload.source_end,
        _CITATION_KEY: payload,
    }


def _attach_standalone_target_markers(nodes: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Attach an authenticated next-line target ID to its preceding owner."""

    output: list[dict[str, Any]] = []
    for node in nodes:
        children = node.get("children")
        if isinstance(children, list):
            node["children"] = _attach_standalone_target_markers(children)
            children = node["children"]

        marker_indices = [
            index
            for index, child in enumerate(children if isinstance(children, list) else [])
            if _is_standalone_target_marker(child)
        ]
        if marker_indices and node.get("type") in {"paragraph", "block_text"}:
            if len(marker_indices) != 1:
                raise ResolvedRuntimeV4Unsupported("one source block contains multiple standalone target IDs")
            marker_index = marker_indices[0]
            assert isinstance(children, list)
            before = children[:marker_index]
            after = children[marker_index + 1 :]
            meaningful_before = [item for item in before if not _is_inline_separator_or_blank(item)]
            meaningful_after = [item for item in after if not _is_inline_separator_or_blank(item)]
            if not meaningful_before and not meaningful_after:
                owner_index = len(output) - 1
                while owner_index >= 0 and output[owner_index].get("type") == "blank_line":
                    owner_index -= 1
                if owner_index < 0 or output[owner_index].get("type") not in {"heading", "paragraph", "block_text"}:
                    raise ResolvedRuntimeV4Unsupported("standalone target ID has no immediately preceding target owner")
                owner_children = output[owner_index].get("children")
                if not isinstance(owner_children, list):
                    raise ResolvedRuntimeV4Unsupported("standalone target ID owner has no inline children")
                owner_children.append(children[marker_index])
                continue
            if meaningful_after:
                raise ResolvedRuntimeV4Unsupported("standalone target ID is not an isolated next-line token")
            while marker_index > 0 and _is_inline_separator_or_blank(children[marker_index - 1]):
                del children[marker_index - 1]
                marker_index -= 1
        output.append(node)
    return output


def _is_standalone_target_marker(node: dict[str, Any]) -> bool:
    if node.get("type") != _TARGET_NODE:
        return False
    marker = node.get("marker")
    return isinstance(marker, RuntimeMarkerV4) and marker.standalone_target_id


def _is_inline_separator_or_blank(node: dict[str, Any]) -> bool:
    if node.get("type") in {"softbreak", "linebreak"}:
        return True
    if node.get("type") == "text":
        return not str(node.get("raw") or node.get("text") or "").strip()
    return False


def _bind_target_markers(
    node: dict[str, Any],
    bound_targets: set[tuple[int, int, str]],
) -> None:
    children = node.get("children")
    if isinstance(children, list):
        for child in children:
            _bind_target_markers(child, bound_targets)
    if not isinstance(children, list):
        return
    indices = [index for index, child in enumerate(children) if child.get("type") == _TARGET_NODE]
    if not indices:
        return
    if len(indices) != 1:
        raise ResolvedRuntimeV4Unsupported("one Markdown block owns multiple v4 target markers")
    index = indices[0]
    marker = children[index].get("marker")
    if not isinstance(marker, RuntimeMarkerV4) or not isinstance(marker.payload, ResolvedDocumentTarget):
        raise ResolvedRuntimeV4Unsupported("v4 target marker payload is not typed")
    target = marker.payload
    if target.occurrence_key in bound_targets:
        raise ResolvedRuntimeV4Unsupported("v4 target marker has duplicate structural owners")
    if target.kind == "heading":
        level = node.get("attrs", {}).get("level") if isinstance(node.get("attrs"), dict) else None
        if node.get("type") != "heading" or level != target.heading_level:
            return
        if index != len(children) - 1:
            raise ResolvedRuntimeV4Unsupported("Heading target marker is not at the authenticated title boundary")
        del children[index]
        if marker.trim_preceding_space:
            _trim_text_before(children, index)
        if authored_inline_text(children) != target.authored_text:
            raise ResolvedRuntimeV4Unsupported("Heading authored_text differs from its parsed ATX title")
        node[_TARGET_KEY] = target
    else:
        if node.get("type") not in {"paragraph", "block_text"}:
            return
        inline_image_before = _split_inline_image_before_caption_children(children, index, target)
        if marker.standalone_target_id:
            if index != len(children) - 1:
                raise ResolvedRuntimeV4Unsupported("standalone caption target ID is not at the declaration boundary")
        elif marker.trim_preceding_space:
            if index != len(children) - 1:
                raise ResolvedRuntimeV4Unsupported("caption target ID marker is not at the declaration boundary")
        elif (
            inline_image_before is None
            and authored_inline_text(children[:index]).casefold()
            != ({"code_block": "Code"}.get(target.kind, target.kind.title()) + ":").casefold()
        ):
            raise ResolvedRuntimeV4Unsupported("ID-less caption target marker moved away from its kind colon")
        node["type"] = "_docwen_resolved_v4_caption_declaration"
        node[_TARGET_KEY] = target
        node[_CAPTION_DIRECTIONS_KEY] = marker.caption_carrier_directions
        if inline_image_before is not None:
            image_children, caption_children = inline_image_before
            node[_INLINE_IMAGE_CHILDREN_KEY] = image_children
            node[_INLINE_IMAGE_BEFORE_KEY] = True
        else:
            caption_children = _caption_content_children(children, index, target)
            inline_image = _split_inline_caption_image_children(caption_children, target)
            if inline_image is not None:
                caption_children, image_children = inline_image
                node[_INLINE_IMAGE_CHILDREN_KEY] = image_children
        if authored_inline_text(caption_children) != target.authored_text:
            raise ResolvedRuntimeV4Unsupported("caption authored_text differs from its parsed declaration content")
        node[_CAPTION_CHILDREN_KEY] = caption_children
        node["children"] = []
    bound_targets.add(target.occurrence_key)


def _caption_content_children(
    children: list[dict[str, Any]],
    marker_index: int,
    target: ResolvedDocumentTarget,
) -> list[dict[str, Any]]:
    output = [child for index, child in enumerate(children) if index != marker_index]
    expected = {"code_block": "Code"}.get(target.kind, target.kind.title()) + ":"
    first_text = next((index for index, child in enumerate(output) if child.get("type") == "text"), None)
    if first_text is None:
        raise ResolvedRuntimeV4Unsupported("caption declaration lost its authored kind prefix")
    key = "raw" if "raw" in output[first_text] else "text"
    value = str(output[first_text].get(key, ""))
    if not value.casefold().startswith(expected.casefold()):
        raise ResolvedRuntimeV4Unsupported("caption AST kind prefix differs from its typed target")
    remainder = value[len(expected) :]
    if remainder:
        output[first_text] = {"type": "text", key: remainder}
    else:
        del output[first_text]
    _trim_edge_text(output)
    return output


def _trim_text_before(children: list[dict[str, Any]], marker_index: int) -> None:
    index = min(marker_index - 1, len(children) - 1)
    if index < 0 or children[index].get("type") != "text":
        return
    key = "raw" if "raw" in children[index] else "text"
    value = str(children[index].get(key, "")).rstrip(" \t")
    if value:
        children[index] = {"type": "text", key: value}
    else:
        del children[index]


def _trim_edge_text(children: list[dict[str, Any]]) -> None:
    while children and children[0].get("type") == "text":
        key = "raw" if "raw" in children[0] else "text"
        value = str(children[0].get(key, "")).lstrip(" \t")
        if value:
            children[0] = {"type": "text", key: value}
            break
        del children[0]
    while children and children[-1].get("type") == "text":
        key = "raw" if "raw" in children[-1] else "text"
        value = str(children[-1].get(key, "")).rstrip(" \t")
        if value:
            children[-1] = {"type": "text", key: value}
            break
        children.pop()


def _split_inline_caption_image_children(
    children: list[dict[str, Any]],
    target: ResolvedDocumentTarget,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]] | None:
    """Split the zero-blank caption+image paragraph produced by Mistune."""

    for index, child in enumerate(children):
        if child.get("type") not in {"softbreak", "linebreak"}:
            continue
        caption_children = children[:index]
        image_children = children[index + 1 :]
        if authored_inline_text(caption_children) != target.authored_text:
            continue
        if _is_image_paragraph({"type": "paragraph", "children": image_children}):
            return caption_children, image_children
    return None


def _split_inline_image_before_caption_children(
    children: list[dict[str, Any]],
    marker_index: int,
    target: ResolvedDocumentTarget,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]] | None:
    """Split an image followed immediately by a caption in one paragraph."""

    separator = next(
        (
            index
            for index in range(marker_index - 1, -1, -1)
            if children[index].get("type") in {"softbreak", "linebreak"}
        ),
        None,
    )
    if separator is None:
        return None
    image_children = children[:separator]
    if not _is_image_paragraph({"type": "paragraph", "children": image_children}):
        return None
    caption_line = list(children[separator + 1 :])
    caption_children = _caption_content_children(caption_line, marker_index - separator - 1, target)
    return image_children, caption_children


def _bind_caption_targets(nodes: list[dict[str, Any]]) -> list[dict[str, Any]]:
    for node in nodes:
        children = node.get("children")
        if isinstance(children, list):
            node["children"] = _bind_caption_targets(children)
    nodes = _expand_inline_caption_images(nodes)
    declarations = [
        index for index, node in enumerate(nodes) if node.get("type") == "_docwen_resolved_v4_caption_declaration"
    ]
    candidate_sets: dict[int, set[int]] = {}
    for declaration_index in declarations:
        declaration = nodes[declaration_index]
        target = declaration.get(_TARGET_KEY)
        if not isinstance(target, ResolvedDocumentTarget) or target.kind == "heading":
            raise ResolvedRuntimeV4Unsupported("caption declaration has no typed caption target")
        directions = declaration.get(_CAPTION_DIRECTIONS_KEY)
        if not isinstance(directions, tuple) or any(direction not in {-1, 1} for direction in directions):
            raise ResolvedRuntimeV4Unsupported("caption declaration lost its authenticated source adjacency")
        candidates = {
            candidate
            for direction in directions
            for candidate in (_adjacent_caption_node(nodes, declaration_index + direction, direction),)
            if candidate is not None and _captionable_node_kind(nodes[candidate]) is not None
        }
        candidate_sets[declaration_index] = candidates
    bindings = _require_unique_caption_bindings(candidate_sets)

    removed = set(declarations)
    for declaration_index, object_index in bindings.items():
        declaration = nodes[declaration_index]
        object_node = nodes[object_index]
        if _TARGET_KEY in object_node or _CAPTION_CHILDREN_KEY in object_node:
            raise ResolvedRuntimeV4Unsupported("one Markdown object participates in multiple caption claims")
        object_node[_TARGET_KEY] = declaration[_TARGET_KEY]
        object_node[_CAPTION_CHILDREN_KEY] = declaration[_CAPTION_CHILDREN_KEY]
        for between in range(min(declaration_index, object_index) + 1, max(declaration_index, object_index)):
            if nodes[between].get("type") == "blank_line":
                removed.add(between)
    return [node for index, node in enumerate(nodes) if index not in removed]


def _expand_inline_caption_images(nodes: list[dict[str, Any]]) -> list[dict[str, Any]]:
    output: list[dict[str, Any]] = []
    for node in nodes:
        image_children = node.get(_INLINE_IMAGE_CHILDREN_KEY)
        if node.get("type") != "_docwen_resolved_v4_caption_declaration" or not isinstance(image_children, list):
            output.append(node)
            continue
        declaration = dict(node)
        declaration.pop(_INLINE_IMAGE_CHILDREN_KEY, None)
        image_before = bool(declaration.pop(_INLINE_IMAGE_BEFORE_KEY, False))
        image_node: dict[str, Any] = {"type": "paragraph", "children": image_children}
        for key in ("_docwen_v3_ordinary_anchor", "_docwen_v3_ordinary_anchor_parent_source_id"):
            if key in declaration:
                image_node[key] = declaration.pop(key)
        output.extend((image_node, declaration) if image_before else (declaration, image_node))
    return output


def _adjacent_caption_node(nodes: list[dict[str, Any]], start: int, direction: int) -> int | None:
    blanks = 0
    stop = len(nodes) if direction > 0 else -1
    for index in range(start, stop, direction):
        if nodes[index].get("type") == "blank_line":
            blanks += 1
            if blanks > 1:
                return None
            continue
        return index
    return None


def _captionable_node_kind(node: dict[str, Any]) -> str | None:
    node_type = node.get("type")
    if node_type == "table":
        return "table"
    if node_type in {"block_math", "block_latex"}:
        return "equation"
    if node_type == "block_code":
        return "code_block"
    if node_type == "paragraph" and _is_image_paragraph(node):
        return "figure"
    return None


def _require_unique_caption_bindings(candidate_sets: dict[int, set[int]]) -> dict[int, int]:
    """Require one local carrier per declaration and one claimant per carrier."""

    if any(len(candidates) != 1 for candidates in candidate_sets.values()):
        raise ResolvedRuntimeV4Unsupported("caption declarations do not have one unique object matching")
    bindings = {caption: next(iter(candidates)) for caption, candidates in candidate_sets.items()}
    if len(set(bindings.values())) != len(bindings):
        raise ResolvedRuntimeV4Unsupported("caption declarations do not have one unique object matching")
    return bindings


def _is_image_paragraph(node: dict[str, Any]) -> bool:
    children = node.get("children")
    if not isinstance(children, list):
        return False
    images = [child for child in children if child.get("type") == "image"]
    if len(images) != 1:
        return False
    text = "".join(
        str(child.get("raw", child.get("text", ""))) for child in children if child.get("type") == "text"
    ).strip()
    if text and _ORDINARY_ID_TEXT_RE.fullmatch(text) is None:
        return False
    return all(child.get("type") in {"image", "text"} for child in children)


def _iter_unbound_target_markers(nodes: list[dict[str, Any]]):
    for node in nodes:
        if node.get("type") == _TARGET_NODE:
            yield node
        children = node.get("children")
        if isinstance(children, list):
            yield from _iter_unbound_target_markers(children)


__all__ = [
    "ResolvedRuntimeV4Plan",
    "ResolvedRuntimeV4Unsupported",
    "RuntimeMarkerEditV4",
    "RuntimeMarkerPayloadV4",
    "RuntimeMarkerRoleV4",
    "RuntimeMarkerV4",
    "apply_resolved_runtime_v4",
    "prepare_resolved_runtime_v4",
]

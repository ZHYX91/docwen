"""Production adapter for the ``docwen.markdown_semantics.v3`` source oracle.

The generic Markdown link pipeline predates semantic ``@[[...]]`` references
and rewrites WikiLinks before Mistune sees them.  This adapter therefore owns
one deliberately small bridge:

1. analyze the exact authenticated input while masking closed YAML front matter;
2. replace every lossless semantic/link token with an inert marker;
3. let the existing link and Markdown pipelines run; and
4. restore typed AST nodes without reparsing or guessing source text.

Only the production slice implemented by the DOCX v3 adapter is admitted.
Unsupported valid-v3 shapes fail closed rather than falling through to the
superseded document-semantics v1 parser.
"""

from __future__ import annotations

import base64
import hashlib
from dataclasses import dataclass
from typing import Any, Literal

from docwen_core.docx_semantics_v3 import fenced_source_identity_from_mapping_v3
from docwen_plugin_markdown.document_semantics_v3 import (
    MarkdownSemanticsV3Analysis,
    analyze_markdown_semantics_v3,
    markdown_semantics_body_start_v3,
)
from docwen_plugin_markdown.document_semantics_v3_fenced_source import (
    fenced_source_info_insertion_offset_v3,
    recover_fenced_logical_body_v3,
)

type _MarkerRole = Literal[
    "heading_target",
    "caption_declaration",
    "semantic_target_id",
    "ordinary_anchor",
    "cross_reference",
    "citation",
    "fenced_source",
]


class RuntimeSemanticsV3Unsupported(ValueError):
    """Raised when valid source is outside the current atomic runtime slice."""


@dataclass(frozen=True, slots=True)
class RuntimeMarkerV3:
    marker: str
    role: _MarkerRole
    payload: dict[str, Any]


@dataclass(frozen=True, slots=True)
class RuntimeSemanticsV3Plan:
    """One immutable source analysis and its inert preprocessor projection."""

    analysis: MarkdownSemanticsV3Analysis
    shielded_source: str
    markers: tuple[RuntimeMarkerV3, ...]
    body_start: int
    ordinary_anchor_parents: tuple[tuple[str, str | None], ...]

    @property
    def source_sha256(self) -> str:
        return str(self.analysis.projection["source"]["sha256"])

    @property
    def shielded_body(self) -> str:
        """Return the original shielded body without a closed YAML prefix.

        A caller that subsequently rewrites ``shielded_source`` must extract
        the body again from that rewritten full source; this property cannot
        represent bytes changed after the immutable plan was created.
        """

        return self.shielded_source[self.body_start :]


def prepare_runtime_semantics_v3(source: str, *, input_id: str) -> RuntimeSemanticsV3Plan:
    """Analyze and shield one exact accepted input before generic processing."""

    try:
        analysis = analyze_markdown_semantics_v3(source, input_id=input_id)
    except ValueError as exc:
        raise RuntimeSemanticsV3Unsupported(str(exc)) from exc
    body_start = markdown_semantics_body_start_v3(source)
    if analysis.has_errors:
        return RuntimeSemanticsV3Plan(
            analysis=analysis,
            shielded_source=source,
            markers=(),
            body_start=body_start,
            ordinary_anchor_parents=(),
        )
    _validate_supported_projection(analysis.projection)
    ordinary_anchor_parents = _ordinary_anchor_parent_ids(analysis.projection)

    source_hash = hashlib.sha256(source.encode("utf-8")).hexdigest()
    marker_prefix = _unique_marker_prefix(source, source_hash)
    marker_specs: list[tuple[int, int, _MarkerRole, dict[str, Any]]] = []

    for target in analysis.projection["targets"]:
        if target["kind"] == "heading":
            id_range = target.get("id_range")
            if id_range is None:
                continue
            if _is_standalone_target_id(source, target):
                insertion = _first_line_content_end(source, int(target["range"]["start"]))
                marker_specs.append((insertion, insertion, "heading_target", dict(target)))
                marker_specs.append((id_range["start"], id_range["end"], "semantic_target_id", dict(target)))
            else:
                marker_specs.append((id_range["start"], id_range["end"], "heading_target", dict(target)))
            continue
        declaration_range = _caption_marker_range(source, target)
        marker_specs.append(
            (
                declaration_range["start"],
                declaration_range["end"],
                "caption_declaration",
                dict(target),
            )
        )
        if _is_standalone_target_id(source, target):
            id_range = target["id_range"]
            marker_specs.append((id_range["start"], id_range["end"], "semantic_target_id", dict(target)))
    for anchor in analysis.projection["anchors"]:
        source_range = anchor["range"]
        marker_specs.append((source_range["start"], source_range["end"], "ordinary_anchor", dict(anchor)))
    for reference in analysis.projection["references"]:
        source_range = reference["range"]
        marker_specs.append((source_range["start"], source_range["end"], "cross_reference", dict(reference)))
    for citation in analysis.projection["citations"]:
        source_range = citation["range"]
        marker_specs.append((source_range["start"], source_range["end"], "citation", dict(citation)))

    marker_specs.sort(key=lambda item: (item[0], item[1], item[2]))
    for fenced_source in analysis.projection["fenced_sources"]:
        record = dict(fenced_source)
        insertion_offset = fenced_source_info_insertion_offset_v3(source, record)
        marker_specs.append(
            (
                insertion_offset,
                insertion_offset,
                "fenced_source",
                {
                    "record": record,
                    "logical_body": recover_fenced_logical_body_v3(source, record),
                },
            )
        )

    marker_specs.sort(key=lambda item: (item[0], item[1], item[2]))
    _require_non_overlapping(marker_specs)
    markers = tuple(
        RuntimeMarkerV3(
            marker=f"{marker_prefix}{index:06d}X",
            role=role,
            payload=payload,
        )
        for index, (_start, _end, role, payload) in enumerate(marker_specs)
    )
    shielded = source
    for (start, end, _role, _payload), marker in reversed(list(zip(marker_specs, markers, strict=True))):
        replacement = marker.marker if end > start else f" {marker.marker}"
        shielded = shielded[:start] + replacement + shielded[end:]
    return RuntimeSemanticsV3Plan(
        analysis=analysis,
        shielded_source=shielded,
        markers=markers,
        body_start=body_start,
        ordinary_anchor_parents=ordinary_anchor_parents,
    )


def _caption_marker_range(source: str, target: dict[str, Any]) -> dict[str, int]:
    declaration_range = target["declaration_range"]
    start = int(declaration_range["start"])
    end = int(declaration_range["end"])
    token = f"{target['source_keyword']}:"
    relative = source[start:end].find(token)
    if relative < 0:
        raise RuntimeSemanticsV3Unsupported("caption declaration prefix is not source-bound")
    marker_start = start + relative
    return {"start": marker_start, "end": marker_start + len(token)}


def _is_standalone_target_id(source: str, target: dict[str, Any]) -> bool:
    id_range = target.get("id_range")
    if not isinstance(id_range, dict):
        return False
    first_line_end = source.find("\n", int(target["range"]["start"]))
    if first_line_end < 0:
        first_line_end = len(source)
    return int(id_range["start"]) > first_line_end


def _first_line_content_end(source: str, start: int) -> int:
    end = source.find("\n", start)
    if end < 0:
        end = len(source)
    if end > start and source[end - 1] == "\r":
        end -= 1
    return end


def apply_runtime_semantics_v3(
    ast_nodes: list[dict[str, Any]],
    plan: RuntimeSemanticsV3Plan,
) -> list[dict[str, Any]]:
    """Restore shielded source constructs as typed runtime AST nodes."""

    if plan.analysis.has_errors:
        raise ValueError("an invalid v3 analysis cannot be applied to an AST")
    markers = {marker.marker: marker for marker in plan.markers}
    restored = [_restore_node(node, markers) for node in ast_nodes]
    _bind_fenced_source_markers(restored, plan)
    for node in restored:
        _bind_block_markers(node)
    restored = _remove_standalone_target_id_nodes(restored)
    restored = _bind_post_block_anchors(restored)
    restored = _bind_caption_targets(restored)
    _lift_container_anchors(restored)
    _attach_ordinary_anchor_parents(restored, plan.ordinary_anchor_parents)
    remaining = tuple(_iter_marker_nodes(restored))
    if remaining:
        roles = ", ".join(sorted({str(item.get("role")) for item in remaining}))
        raise RuntimeSemanticsV3Unsupported(f"v3 source markers were not owned by a supported Markdown block: {roles}")
    return restored


def _validate_supported_projection(projection: dict[str, Any]) -> None:
    previous_end = -1
    tags: set[str] = set()
    for record in projection["fenced_sources"]:
        identity = fenced_source_identity_from_mapping_v3(record)
        if identity.tag in tags:
            raise RuntimeSemanticsV3Unsupported("fenced source tags must be unique")
        if identity.source_start < previous_end:
            raise RuntimeSemanticsV3Unsupported("fenced source ranges must not overlap")
        tags.add(identity.tag)
        previous_end = identity.source_end
    unsupported_anchors = [
        item["block_kind"]
        for item in projection["anchors"]
        if item["block_kind"]
        not in {
            "paragraph",
            "image",
            "table",
            "equation",
            "code_block",
            "fenced_block",
            "list",
            "list_item",
            "block_quote",
            "callout",
        }
    ]
    if unsupported_anchors:
        kinds = ", ".join(sorted(set(unsupported_anchors)))
        raise RuntimeSemanticsV3Unsupported(
            f"the current production slice does not yet project these structured anchors: {kinds}"
        )
    for reference in projection["references"]:
        if reference["resolution_status"] != "resolved" or reference.get("page_locator") is not None:
            raise RuntimeSemanticsV3Unsupported(
                "cross-document or unresolved semantic references require the external neutral resolver boundary"
            )
        if reference["selector_kind"] == "stable_id":
            continue
        if reference.get("resolved_kind") != "heading":
            raise RuntimeSemanticsV3Unsupported("soft references must resolve to Heading")


def _ordinary_anchor_parent_ids(projection: dict[str, Any]) -> tuple[tuple[str, str | None], ...]:
    """Derive the longest-proper-prefix ordinary owner forest from source proof."""

    anchors = list(projection["anchors"])
    owner_paths = {str(anchor["id"]): _ordinary_anchor_owner_path(anchor) for anchor in anchors}
    output: list[tuple[str, str | None]] = []
    for anchor in anchors:
        source_id = str(anchor["id"])
        owner_path = owner_paths[source_id]
        candidates = [
            other_id
            for other_id, other_path in owner_paths.items()
            if other_id != source_id
            and len(other_path) < len(owner_path)
            and owner_path[: len(other_path)] == other_path
        ]
        if candidates:
            longest = max(len(owner_paths[item]) for item in candidates)
            parents = [item for item in candidates if len(owner_paths[item]) == longest]
            if len(parents) != 1:
                raise RuntimeSemanticsV3Unsupported("ordinary-anchor source topology has ambiguous direct parents")
            parent_id: str | None = parents[0]
        else:
            parent_id = None
        output.append((source_id, parent_id))
    return tuple(output)


def _ordinary_anchor_owner_path(anchor: dict[str, Any]) -> tuple[tuple[str, int, int], ...]:
    segments = [
        (
            str(segment["block_kind"]),
            int(segment["block_range"]["start"]),
            int(segment["block_range"]["end"]),
        )
        for segment in anchor["container_path"]
    ]
    block_range = anchor["block_range"]
    segments.append((str(anchor["block_kind"]), int(block_range["start"]), int(block_range["end"])))
    previous: tuple[str, int, int] | None = None
    for segment in segments:
        _kind, start, end = segment
        if start < 0 or end <= start:
            raise RuntimeSemanticsV3Unsupported("ordinary-anchor source topology has an invalid range")
        if previous is not None:
            _previous_kind, previous_start, previous_end = previous
            if start < previous_start or end > previous_end:
                raise RuntimeSemanticsV3Unsupported("ordinary-anchor source topology is not laminar")
            if segment == previous:
                raise RuntimeSemanticsV3Unsupported("ordinary-anchor container path contains its own block")
        previous = segment
    return tuple(segments)


def _unique_marker_prefix(source: str, source_hash: str) -> str:
    for salt in range(256):
        prefix = f"DOCWENV3X{source_hash[:16].upper()}X{salt:02X}X"
        if prefix not in source:
            return prefix
    raise RuntimeSemanticsV3Unsupported("source exhausts the inert marker namespace")


def _require_non_overlapping(
    marker_specs: list[tuple[int, int, _MarkerRole, dict[str, Any]]],
) -> None:
    previous_end = -1
    insertion_points: set[int] = set()
    for start, end, _role, _payload in marker_specs:
        if start < 0 or end < start or start < previous_end:
            raise RuntimeSemanticsV3Unsupported("v3 runtime marker ranges overlap")
        if end == start:
            if start in insertion_points:
                raise RuntimeSemanticsV3Unsupported("v3 runtime marker insertion points collide")
            insertion_points.add(start)
        else:
            previous_end = end


def _bind_fenced_source_markers(
    nodes: list[dict[str, Any]],
    plan: RuntimeSemanticsV3Plan,
) -> None:
    expected = {marker.marker: marker for marker in plan.markers if marker.role == "fenced_source"}
    bound: set[str] = set()

    def visit(node: dict[str, Any]) -> None:
        attrs = node.get("attrs")
        info = str(attrs.get("info", "")) if isinstance(attrs, dict) else ""
        matches = [marker for token, marker in expected.items() if token in info]
        if matches:
            if node.get("type") != "block_code" or len(matches) != 1:
                raise RuntimeSemanticsV3Unsupported("a fenced source marker must have exactly one block_code owner")
            marker = matches[0]
            if info.count(marker.marker) != 1 or not info.endswith(marker.marker):
                raise RuntimeSemanticsV3Unsupported("a fenced source marker was altered in the info string")
            if marker.marker in bound:
                raise RuntimeSemanticsV3Unsupported("a fenced source marker has duplicate AST owners")
            record = marker.payload["record"]
            exact_info = base64.b64decode(record["info_b64"], validate=True).decode("utf-8").strip()
            restored_attrs = dict(attrs or {})
            if exact_info:
                restored_attrs["info"] = exact_info
            else:
                restored_attrs.pop("info", None)
            node["attrs"] = restored_attrs
            node["_docwen_v3_fenced_source"] = record
            node["_docwen_v3_fenced_body"] = marker.payload["logical_body"]
            bound.add(marker.marker)
        children = node.get("children")
        if isinstance(children, list):
            for child in children:
                visit(child)

    for node in nodes:
        visit(node)
    missing = set(expected) - bound
    if missing:
        raise RuntimeSemanticsV3Unsupported("fenced source markers lost their block_code owners")


def _restore_node(
    node: dict[str, Any],
    markers: dict[str, RuntimeMarkerV3],
) -> dict[str, Any]:
    output = dict(node)
    children = node.get("children")
    if isinstance(children, list):
        output["children"] = _restore_children(children, markers)
    return output


def _restore_children(
    children: list[dict[str, Any]],
    markers: dict[str, RuntimeMarkerV3],
) -> list[dict[str, Any]]:
    output: list[dict[str, Any]] = []
    for child in children:
        restored = _restore_node(child, markers)
        if restored.get("type") != "text":
            output.append(restored)
            continue
        key = "raw" if "raw" in restored else "text"
        text = str(restored.get(key, ""))
        cursor = 0
        while cursor < len(text):
            match = _next_marker(text, cursor, markers)
            if match is None:
                if cursor < len(text):
                    output.append({"type": "text", key: text[cursor:]})
                break
            marker_start, marker = match
            if marker_start > cursor:
                output.append({"type": "text", key: text[cursor:marker_start]})
            output.append(_marker_ast_node(marker))
            cursor = marker_start + len(marker.marker)
    return output


def _next_marker(
    text: str,
    start: int,
    markers: dict[str, RuntimeMarkerV3],
) -> tuple[int, RuntimeMarkerV3] | None:
    matches = ((text.find(token, start), marker) for token, marker in markers.items())
    present = [(position, marker) for position, marker in matches if position >= 0]
    return min(present, key=lambda item: item[0]) if present else None


def _marker_ast_node(marker: RuntimeMarkerV3) -> dict[str, Any]:
    if marker.role == "cross_reference":
        return {
            "type": "semantic_cross_reference",
            "schema": "docwen.markdown_semantics.v3",
            **marker.payload,
        }
    if marker.role == "citation":
        return {
            "type": "semantic_citation",
            "schema": "docwen.markdown_semantics.v3",
            **marker.payload,
        }
    return {
        "type": "_docwen_v3_source_marker",
        "role": marker.role,
        "payload": marker.payload,
    }


def _bind_block_markers(node: dict[str, Any]) -> None:
    children = node.get("children")
    if isinstance(children, list):
        for child in children:
            _bind_block_markers(child)
    if node.get("type") not in {"heading", "paragraph", "block_text"} or not isinstance(children, list):
        return
    marker_indices = [index for index, child in enumerate(children) if child.get("type") == "_docwen_v3_source_marker"]
    if not marker_indices:
        return
    bound_indices: list[int] = []
    bound_roles: set[str] = set()
    original_type = str(node["type"])
    for marker_index in marker_indices:
        marker = children[marker_index]
        role = str(marker["role"])
        if role in bound_roles:
            raise RuntimeSemanticsV3Unsupported(f"a supported block cannot own duplicate {role} markers")
        if original_type == "heading" and role == "heading_target":
            node["_docwen_v3_heading_target"] = marker["payload"]
        elif original_type in {"paragraph", "block_text"} and role == "ordinary_anchor":
            node["_docwen_v3_ordinary_anchor"] = marker["payload"]
        elif original_type in {"paragraph", "block_text"} and role == "caption_declaration":
            node["type"] = "_docwen_v3_caption_declaration"
            node["_docwen_v3_caption_target"] = marker["payload"]
        elif original_type in {"paragraph", "block_text"} and role == "semantic_target_id":
            node["_docwen_v3_standalone_target_id"] = marker["payload"]
        else:
            continue
        bound_roles.add(role)
        bound_indices.append(marker_index)
    for marker_index in reversed(bound_indices):
        _remove_bound_marker(children, marker_index)


def _remove_standalone_target_id_nodes(nodes: list[dict[str, Any]]) -> list[dict[str, Any]]:
    output: list[dict[str, Any]] = []
    for node in nodes:
        children = node.get("children")
        if isinstance(children, list):
            node["children"] = _remove_standalone_target_id_nodes(children)
        if node.get("_docwen_v3_standalone_target_id") is not None:
            node.pop("_docwen_v3_standalone_target_id", None)
            if node.get("type") in {"paragraph", "block_text"} and not _plain_children(node.get("children", [])):
                continue
        output.append(node)
    return output


def _remove_bound_marker(children: list[dict[str, Any]], marker_index: int) -> None:
    del children[marker_index]
    if marker_index > 0 and children[marker_index - 1].get("type") == "text":
        key = "raw" if "raw" in children[marker_index - 1] else "text"
        children[marker_index - 1][key] = str(children[marker_index - 1].get(key, "")).rstrip(" \t")
        if not children[marker_index - 1][key]:
            del children[marker_index - 1]


def _bind_caption_targets(nodes: list[dict[str, Any]]) -> list[dict[str, Any]]:
    for node in nodes:
        children = node.get("children")
        if isinstance(children, list):
            node["children"] = _bind_caption_targets(children)
    nodes = _split_inline_caption_images(nodes)
    declarations = [index for index, node in enumerate(nodes) if node.get("type") == "_docwen_v3_caption_declaration"]
    candidate_sets: dict[int, set[int]] = {}
    for declaration_index in declarations:
        target = nodes[declaration_index]["_docwen_v3_caption_target"]
        declaration_start = int(target["declaration_range"]["start"])
        object_start = int(target["object_range"]["start"])
        direction = -1 if object_start < declaration_start else 1
        candidates = {
            candidate
            for candidate in (_adjacent_caption_node(nodes, declaration_index + direction, direction),)
            if candidate is not None and _captionable_node_kind(nodes[candidate]) is not None
        }
        candidate_sets[declaration_index] = candidates
    bindings = _require_unique_caption_bindings(candidate_sets)

    removed = set(declarations)
    for declaration_index, object_index in bindings.items():
        target = nodes[declaration_index]["_docwen_v3_caption_target"]
        object_node = nodes[object_index]
        if "_docwen_v3_caption_target" in object_node:
            raise RuntimeSemanticsV3Unsupported("one Markdown object participates in multiple caption claims")
        object_node["_docwen_v3_caption_target"] = target
        for between in range(min(declaration_index, object_index) + 1, max(declaration_index, object_index)):
            if nodes[between].get("type") == "blank_line":
                removed.add(between)
    return [node for index, node in enumerate(nodes) if index not in removed]


def _split_inline_caption_images(nodes: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Split either zero-blank declaration/image paragraph Mistune creates."""

    output: list[dict[str, Any]] = []
    for node in nodes:
        children = node.get("children")
        if node.get("type") != "_docwen_v3_caption_declaration" or not isinstance(children, list):
            output.append(node)
            continue
        separator = next(
            (index for index, child in enumerate(children) if child.get("type") in {"softbreak", "linebreak"}),
            None,
        )
        if separator is None:
            output.append(node)
            continue
        before = {"type": "paragraph", "children": children[:separator]}
        after = {"type": "paragraph", "children": children[separator + 1 :]}
        if _is_image_paragraph(after):
            declaration_children = children[:separator]
            image_node = after
            caption_first = True
        elif _is_image_paragraph(before):
            declaration_children = children[separator + 1 :]
            image_node = before
            caption_first = False
        else:
            output.append(node)
            continue
        declaration = dict(node)
        declaration["children"] = declaration_children
        for key in (
            "_docwen_v3_ordinary_anchor",
            "_docwen_v3_ordinary_anchor_parent_source_id",
        ):
            if key in node:
                image_node[key] = node[key]
                declaration.pop(key, None)
        output.extend((declaration, image_node) if caption_first else (image_node, declaration))
    return output


def _adjacent_caption_node(nodes: list[dict[str, Any]], start: int, direction: int) -> int | None:
    blank_nodes = 0
    stop = len(nodes) if direction > 0 else -1
    for index in range(start, stop, direction):
        if nodes[index].get("type") == "blank_line":
            blank_nodes += 1
            if blank_nodes > 1:
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
        raise RuntimeSemanticsV3Unsupported("caption declarations do not have one unique object matching")
    bindings = {caption: next(iter(candidates)) for caption, candidates in candidate_sets.items()}
    if len(set(bindings.values())) != len(bindings):
        raise RuntimeSemanticsV3Unsupported("caption declarations do not have one unique object matching")
    return bindings


def _bind_post_block_anchors(nodes: list[dict[str, Any]]) -> list[dict[str, Any]]:
    for node in nodes:
        children = node.get("children")
        if isinstance(children, list):
            node["children"] = _bind_post_block_anchors(children)
    output: list[dict[str, Any]] = []
    for node in nodes:
        anchor = node.get("_docwen_v3_ordinary_anchor")
        if (
            node.get("type") in {"paragraph", "_docwen_v3_caption_declaration"}
            and anchor is not None
            and anchor.get("placement") == "post_block"
        ):
            is_marker_only = node.get("type") == "paragraph" and not _plain_children(node.get("children", []))
            is_caption_boundary = node.get("type") == "_docwen_v3_caption_declaration"
            if not (is_marker_only or is_caption_boundary):
                output.append(node)
                continue
            owner_index = len(output) - 1
            while owner_index >= 0 and output[owner_index].get("type") == "blank_line":
                owner_index -= 1
            if owner_index < 0:
                raise RuntimeSemanticsV3Unsupported("post-block anchor lost its structured owner")
            output[owner_index]["_docwen_v3_ordinary_anchor"] = anchor
            if is_marker_only:
                continue
            node.pop("_docwen_v3_ordinary_anchor", None)
        output.append(node)
    return output


def _lift_container_anchors(nodes: list[dict[str, Any]]) -> None:
    """Move a marker restored inside a quote/list onto its complete owner."""

    for node in nodes:
        children = node.get("children")
        if not isinstance(children, list):
            continue
        _lift_container_anchors(children)
        if node.get("type") not in {"block_quote", "list", "list_item"}:
            continue
        owners = list(_direct_anchor_owners(children))
        if len(owners) != 1:
            continue
        anchor = owners[0]["_docwen_v3_ordinary_anchor"]
        expected = {
            "block_quote": "block_quote",
            "list": "list",
            "list_item": "list_item",
        }[str(node["type"])]
        if node.get("type") == "block_quote" and anchor.get("block_kind") == "callout":
            expected = "callout"
        if anchor.get("block_kind") != expected:
            continue
        del owners[0]["_docwen_v3_ordinary_anchor"]
        node["_docwen_v3_ordinary_anchor"] = anchor


def _attach_ordinary_anchor_parents(
    nodes: list[dict[str, Any]],
    parent_records: tuple[tuple[str, str | None], ...],
) -> None:
    expected = dict(parent_records)
    observed: set[str] = set()

    def visit(node: dict[str, Any]) -> None:
        anchor = node.get("_docwen_v3_ordinary_anchor")
        if isinstance(anchor, dict):
            source_id = str(anchor.get("id") or "")
            if source_id not in expected or source_id in observed:
                raise RuntimeSemanticsV3Unsupported("ordinary-anchor runtime topology does not match source proof")
            node["_docwen_v3_ordinary_anchor_parent_source_id"] = expected[source_id]
            observed.add(source_id)
        children = node.get("children")
        if isinstance(children, list):
            for child in children:
                visit(child)

    for node in nodes:
        visit(node)
    if observed != set(expected):
        raise RuntimeSemanticsV3Unsupported("ordinary-anchor runtime topology lost a source owner")


def _direct_anchor_owners(nodes: list[dict[str, Any]]):
    """Yield marker owners through transparent paragraph/block-text shells."""

    for node in nodes:
        if node.get("_docwen_v3_ordinary_anchor") is not None:
            yield node
        if node.get("type") in {"paragraph", "block_text"}:
            children = node.get("children")
            if isinstance(children, list):
                yield from _direct_anchor_owners(children)


def _plain_children(children: list[dict[str, Any]]) -> str:
    return "".join(str(item.get("raw") or item.get("text") or "") for item in children).strip()


def _is_image_paragraph(node: dict[str, Any]) -> bool:
    meaningful = [
        item
        for item in node.get("children", [])
        if item.get("type") not in {"softbreak", "linebreak"}
        and (item.get("type") != "text" or str(item.get("raw", "")).strip())
    ]
    return len(meaningful) == 1 and meaningful[0].get("type") == "image"


def _iter_marker_nodes(nodes: list[dict[str, Any]]):
    for node in nodes:
        if node.get("type") == "_docwen_v3_source_marker":
            yield node
        children = node.get("children")
        if isinstance(children, list):
            yield from _iter_marker_nodes(children)


__all__ = [
    "RuntimeMarkerV3",
    "RuntimeSemanticsV3Plan",
    "RuntimeSemanticsV3Unsupported",
    "apply_runtime_semantics_v3",
    "prepare_runtime_semantics_v3",
]

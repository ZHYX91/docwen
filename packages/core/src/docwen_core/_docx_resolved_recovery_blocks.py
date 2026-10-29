"""Authenticated block ownership admitted only by the resolved-v4 reader."""

from __future__ import annotations

from typing import Any

from docwen_core._docx_semantics_v3_model import (
    ANCHOR_TAG_PREFIX,
    TARGET_TAG_PREFIX,
    DocxSemanticsV3Error,
)
from docwen_core._docx_semantics_v3_ooxml import sdt_tag
from docwen_core._docx_semantics_v3_styles import caption_kind_for_paragraph_style
from docwen_core._docx_semantics_v3_topology import logical_group_elements, prove_block_sdt_envelope
from docwen_core.docx_standalone_caption_occurrence import STANDALONE_CAPTION_OCCURRENCE_TAG_PREFIX

_OCCURRENCE_PREFIX = "docwen-numbering-occurrence-v1:"
_BLOCK_PREFIXES = (ANCHOR_TAG_PREFIX, TARGET_TAG_PREFIX, _OCCURRENCE_PREFIX, STANDALONE_CAPTION_OCCURRENCE_TAG_PREFIX)
_CONTAINERS = frozenset({"list", "list_item", "block_quote", "callout"})
_OBJECT_KINDS = frozenset({"image", "table", "equation", "code_block"})


def prove_resolved_recovery_blocks_v4(
    body: Any,
    *,
    targets: list[Any],
    anchors: list[Any],
    occurrences: list[Any],
    standalone_occurrences: list[Any],
    inline_tags: frozenset[str],
    caption_styles: tuple[Any, ...],
) -> tuple[dict[Any, tuple[Any, ...]], frozenset[Any], frozenset[Any]]:
    """Admit exact mapped wrappers, their envelopes, and closed parent roles.

    The content proofs still run in the semantic reader before it returns.
    This admission merely supplies the finite set that anchor topology may
    flatten; an unknown SDT or an unmapped lookalike never gains ownership.
    """

    from docx.oxml.ns import qn

    records = [*targets, *anchors, *occurrences, *(item for item in standalone_occurrences if item.target_id is None)]
    by_tag = {item.tag: item for item in records}
    if len(by_tag) != len(records):
        raise DocxSemanticsV3Error("resolved recovery block maps duplicate an owner")
    anchor_by_tag = {item.tag: item for item in anchors}
    wrappers: dict[str, Any] = {}
    contents: dict[Any, tuple[Any, ...]] = {}
    for wrapper in body.iter(qn("w:sdt")):
        tag = sdt_tag(wrapper) or ""
        if not tag.startswith(_BLOCK_PREFIXES):
            continue
        if tag not in by_tag or tag in wrappers:
            raise DocxSemanticsV3Error("resolved recovery block is duplicated or unmapped")
        prove_block_sdt_envelope(wrapper, tag)
        content = wrapper.find(qn("w:sdtContent"))
        if content is None or not len(content):
            raise DocxSemanticsV3Error("resolved recovery block has no content")
        wrappers[tag] = wrapper
        contents[wrapper] = tuple(content)
    if wrappers.keys() != by_tag.keys():
        raise DocxSemanticsV3Error("resolved recovery block map has no exact physical carrier")
    for tag, wrapper in wrappers.items():
        parent = wrapper.getparent()
        if parent is body:
            continue
        parent_wrapper = None if parent is None else parent.getparent()
        parent_tag = sdt_tag(parent_wrapper) if parent_wrapper is not None else None
        if parent is None or parent.tag != qn("w:sdtContent") or parent_wrapper not in contents:
            raise DocxSemanticsV3Error("resolved recovery block has an unowned or non-block parent")
        parent_anchor = anchor_by_tag.get(parent_tag)
        child_anchor = anchor_by_tag.get(tag)
        if parent_anchor is not None:
            if child_anchor is None and parent_anchor.block_kind not in _CONTAINERS:
                raise DocxSemanticsV3Error("resolved semantic carrier is inside a non-container anchor")
        elif (
            child_anchor is None
            or child_anchor.block_kind not in _OBJECT_KINDS
            or by_tag[parent_tag].kind not in {"figure", "table", "equation", "code_block"}
        ):
            raise DocxSemanticsV3Error("resolved semantic block has an unsupported nested owner")

    # Exempt only the exact semantic paragraph / mapped inline slots whose
    # content proofs run before recovery returns. Native caption objects do
    # not gain permission for arbitrary bookmarks or fields merely because
    # they share a target wrapper with an authenticated caption.
    machine_owners: list[Any] = []
    reserved_semantic_elements: set[Any] = set()
    for tag, wrapper in wrappers.items():
        if tag in anchor_by_tag:
            continue
        record = by_tag[tag]
        blocks = contents[wrapper]
        reserved_semantic_elements.update(logical_group_elements(blocks, proven_nested_blocks=contents))
        if record.kind == "heading":
            machine_owners.extend(logical_group_elements(blocks, proven_nested_blocks=contents))
        else:
            paragraph = blocks[-1] if record.kind == "figure" else blocks[0]
            if paragraph.tag == qn("w:p"):
                machine_owners.append(paragraph)
    # Enabled ID-less captions use the existing native style/SEQ pairing.
    # Admit their exact main-story paragraph candidates here; _recover_captions
    # must then prove style, field skeleton, and adjacent native object before
    # returning. Table-cell paragraphs and arbitrary SDT content are excluded.
    if caption_styles:
        machine_owners.extend(
            block
            for block in logical_group_elements(tuple(body), proven_nested_blocks=contents)
            if block.tag == qn("w:p") and caption_kind_for_paragraph_style(block, caption_styles) is not None
        )
    machine_owners.extend(wrapper for wrapper in body.iter(qn("w:sdt")) if sdt_tag(wrapper) in inline_tags)
    machine_nodes = frozenset(
        node
        for owner in machine_owners
        for name in ("w:bookmarkStart", "w:bookmarkEnd", "w:instrText", "w:fldChar")
        for node in owner.iter(qn(name))
    )
    return contents, machine_nodes, frozenset(reserved_semantic_elements)

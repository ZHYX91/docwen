"""Append Word comments without rewriting template-owned identities or styles.

The document/element interfaces are duck typed; importing this module requires
neither python-docx nor an XML backend. Business callers own span selection.
"""

from __future__ import annotations

from collections.abc import Sequence
from copy import deepcopy
from itertools import pairwise
from typing import Any

from docwen_core.docx_xml import parse_word_xml_int

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_COMMENT_TAGS = {_W + name for name in ("comment", "commentRangeStart", "commentRangeEnd", "commentReference")}


def _child(parent: Any, name: str, **attributes: str) -> Any:
    node = parent.makeelement(_W + name, {_W + key: value for key, value in attributes.items()})
    parent.append(node)
    return node


def _comment_style(styles: Any, name: str, kind: str) -> str:
    """Only add new definitions; leave conflicting old IDs/names untouched."""
    definitions = list(styles)
    by_id: dict[str, list[Any]] = {}
    names: set[str] = set()
    for style in definitions:
        by_id.setdefault(style.get(_W + "styleId", ""), []).append(style)
        label = style.find(_W + "name")
        if label is not None:
            names.add(label.get(_W + "val", ""))

    def usable(candidate: str) -> bool:
        found = by_id.get(candidate, [])
        return len(found) == 1 and found[0].get(_W + "type") == kind

    if usable(name):
        return name
    candidate = name
    if candidate in by_id or candidate in names:
        base = f"DocWenComment{name}"
        candidate = base
        ordinal = 0
        while candidate in by_id or candidate in names:
            if usable(candidate):
                return candidate
            ordinal += 1
            candidate = f"{base}{ordinal}"
    style = _child(styles, "style", type=kind, customStyle="1", styleId=candidate)
    _child(style, "name", val=candidate)
    return candidate


def append_docx_comment(
    document: Any, runs: Sequence[Any], *, text: str, author: str = "DocWen", initials: str = "DW"
) -> int:
    """Add a comment over an already selected, ordered, nonempty run range.

    Reserve all numeric definitions and anchors, including orphan references
    in headers/notes and raw XML parts. Malformed legacy IDs remain untouched.
    """
    if not runs:
        raise ValueError("A comment requires a nonempty run range")
    first, last = runs[0]._r, runs[-1]._r
    # Comments written by these consumers belong to the main story. Reject
    # detached, foreign/header/note runs and reversed/repeated selections before
    # materializing any comments/styles parts.
    positions = {
        node: index
        for index, node in enumerate(document.element.iter(_W + "r"))
        if not any(ancestor.tag == _W + "txbxContent" for ancestor in node.iterancestors())
    }
    selected = [positions.get(run._r) for run in runs]
    if any(index is None for index in selected):
        raise ValueError("A comment range must belong to the document main story")
    ordered = [index for index in selected if index is not None]
    if any(left >= right for left, right in pairwise(ordered)):
        raise ValueError("A comment range must contain distinct runs in document order")
    from lxml import etree

    used: set[int] = set()
    for part in document.part.package.parts:
        root = getattr(part, "element", None)
        if root is None:
            if not part.content_type.endswith("xml"):
                continue
            try:
                root = etree.fromstring(part.blob, parser=etree.XMLParser(resolve_entities=False, no_network=True))
            except etree.XMLSyntaxError:
                continue
        for node in root.iter():
            if node.tag in _COMMENT_TAGS:
                identity = parse_word_xml_int(node.get(_W + "id"))
                if identity is not None:
                    used.add(identity)
    identity = 0
    while identity in used:
        identity += 1
    if identity >= 2**31:
        raise ValueError("Word comment identities are exhausted")

    original_relationships = set(document.part.rels)
    attached: list[Any] = []
    try:
        styles = document.styles.element
        style_count = len(styles)
        planned_styles = deepcopy(styles)
        reference_style = _comment_style(planned_styles, "CommentReference", "character")
        text_style = _comment_style(planned_styles, "CommentText", "paragraph")
        comment = document.element.makeelement(
            _W + "comment", {_W + "id": str(identity), _W + "author": author, _W + "initials": initials}
        )
        # Match Word's line/tab representation rather than putting control
        # characters into a single w:t, including localized multiline reviews.
        for line in text.split("\n"):
            paragraph = _child(comment, "p")
            _child(_child(paragraph, "pPr"), "pStyle", val=text_style)
            if paragraph is comment[0]:
                annotation = _child(paragraph, "r")
                _child(_child(annotation, "rPr"), "rStyle", val=reference_style)
                _child(annotation, "annotationRef")
            body = _child(paragraph, "r")
            for index, segment in enumerate(line.split("\t")):
                if index:
                    _child(body, "tab")
                node = _child(body, "t")
                node.text = segment
                if segment[:1].isspace() or segment[-1:].isspace():
                    node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        start = first.makeelement(_W + "commentRangeStart", {_W + "id": str(identity)})
        end = last.makeelement(_W + "commentRangeEnd", {_W + "id": str(identity)})
        reference = last.makeelement(_W + "r", {})
        _child(_child(reference, "rPr"), "rStyle", val=reference_style)
        _child(reference, "commentReference", id=str(identity))
        # All fallible content construction is complete. Track each owned
        # insertion before committing so even a later anchor failure rolls back.
        comments = document.comments._comments_elm
        for style in list(planned_styles)[style_count:]:
            attached.append(style)
            styles.append(style)
        attached.append(comment)
        comments.append(comment)
        attached.extend((start, end, reference))
        first.addprevious(start)
        last.addnext(reference)
        last.addnext(end)
    except BaseException:
        for node in reversed(attached):
            parent = node.getparent()
            if parent is not None:
                parent.remove(node)
        # Accessing document.comments/styles may lazily create a new part.
        # Removing only newly introduced relationships restores reachability;
        # existing template relationships are never replaced or removed.
        for relationship in set(document.part.rels) - original_relationships:
            # python-docx's drop_rel() also only deletes the dictionary entry;
            # its internal target index must be synchronized explicitly.
            document.part.related_parts.pop(relationship, None)
            del document.part.rels[relationship]
        raise
    return identity

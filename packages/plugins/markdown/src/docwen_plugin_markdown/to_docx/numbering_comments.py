"""Review comments anchored only to request-owned, newly inserted numbering."""

from __future__ import annotations

from typing import Any

from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml import etree

NUMBERING_REVIEW_TEXT = "此序号由软件自动添加，请核对标题级别及编号是否正确。"
_COMMENT_TAGS = {qn("w:comment"), qn("w:commentRangeStart"), qn("w:commentRangeEnd"), qn("w:commentReference")}


def add_numbering_review(document: Any, run: Any) -> None:
    """Allocate against definitions and anchors in every existing package part.

    python-docx's default allocator sees only comment definitions. Templates
    can also contain orphan references, including in headers or notes; never
    reuse those identities or replace existing comment parts/relationships.
    """
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
        for element in root.iter():
            if element.tag in _COMMENT_TAGS:
                value = element.get(qn("w:id"), "")
                try:
                    used.add(int(value))
                except ValueError:
                    continue
    comment_id = 0
    while comment_id in used:
        comment_id += 1
    reference_style = _comment_style(document, "CommentReference", WD_STYLE_TYPE.CHARACTER)
    text_style = _comment_style(document, "CommentText", WD_STYLE_TYPE.PARAGRAPH)
    # Append only our own node. The library allocator coerces every existing
    # definition ID with int(), so a malformed legacy ID would otherwise make
    # this optional feature fail or require rewriting the original comments.
    comment = OxmlElement("w:comment")
    for key, value in (("id", str(comment_id)), ("author", "DocWen"), ("initials", "DW")):
        comment.set(qn(f"w:{key}"), value)
    paragraph = OxmlElement("w:p")
    properties = OxmlElement("w:pPr")
    style = OxmlElement("w:pStyle")
    style.set(qn("w:val"), text_style)
    properties.append(style)
    paragraph.append(properties)
    annotation = OxmlElement("w:r")
    properties = OxmlElement("w:rPr")
    style = OxmlElement("w:rStyle")
    style.set(qn("w:val"), reference_style)
    properties.append(style)
    annotation.append(properties)
    annotation.append(OxmlElement("w:annotationRef"))
    paragraph.append(annotation)
    body = OxmlElement("w:r")
    text = OxmlElement("w:t")
    text.text = NUMBERING_REVIEW_TEXT
    body.append(text)
    paragraph.append(body)
    comment.append(paragraph)
    document.comments._comments_elm.append(comment)
    run.mark_comment_range(run, comment_id)
    # python-docx emits a fixed CommentReference style on this new reference.
    reference_run = run._r.getnext().getnext()
    for element in reference_run.iter(qn("w:rStyle")):
        element.set(qn("w:val"), reference_style)


def _comment_style(document: Any, name: str, kind: Any) -> str:
    """Preserve conflicting template definitions; bind only new comment nodes."""
    by_id: dict[str, list[Any]] = {}
    for style in document.styles:
        by_id.setdefault(style.style_id, []).append(style)
    existing = by_id.get(name, [])
    if len(existing) == 1 and existing[0].type == kind:
        return existing[0].style_id
    if not existing and name not in document.styles:
        return document.styles.add_style(name, kind).style_id
    base = f"DocWenNumberReview{name}"
    candidate = base
    suffix = 0
    while candidate in by_id or candidate in document.styles:
        existing = by_id.get(candidate, [])
        if len(existing) == 1 and existing[0].type == kind:
            return existing[0].style_id
        suffix += 1
        candidate = f"{base}{suffix}"
    return document.styles.add_style(candidate, kind).style_id


def annotate_numbering_prefixes(nodes: list[dict[str, Any]], prefixes: list[str]) -> int:
    """Bind the numbering transform's exact inventory to parsed headings.

    Do not infer ownership from visible numbers: authored handwritten numbers
    remain part of the title. Blank scheme levels still consume inventory slots.
    """
    headings: list[dict[str, Any]] = []

    def visit(items: list[dict[str, Any]]) -> None:
        for item in items:
            if item.get("type") == "heading":
                headings.append(item)
            else:
                visit(item.get("children", []))

    visit(nodes)
    if len(headings) != len(prefixes):
        return sum(bool(prefix.rstrip()) for prefix in prefixes)
    skipped = 0
    for node, prefix in zip(headings, prefixes, strict=True):
        # Whitespace separating a number from its title is not the number.
        number = prefix.rstrip()
        if not number:
            continue
        children = node.get("children", [])
        if not children or children[0].get("type") != "text" or not children[0].get("raw", "").startswith(number):
            skipped += 1
            continue
        node["_docwen_added_number"] = number
    return skipped

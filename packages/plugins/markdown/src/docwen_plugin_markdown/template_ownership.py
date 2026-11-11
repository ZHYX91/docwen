"""Read-only ownership checks and text spans for Word template placeholders.

Word fields and data-bound content controls own their displayed text. Editing
that cache would only last until Word recalculates it. Template substitution
therefore accepts plain text, including unbound controls, and rejects markers
inside or across those owned regions before any document changes are made.
"""

from __future__ import annotations

import re
from collections.abc import Iterator
from dataclasses import dataclass, field
from typing import Any

from docx.oxml.ns import qn

from docwen_plugin_markdown.template_errors import TemplatePackageError

_PLACEHOLDER = re.compile(r"\{\{\s*([^{}\r\n]+?)\s*\}\}")
_PARAGRAPH = qn("w:p")
_TEXT = qn("w:t")
_CONTROLS = {
    qn(f"w:{name}")
    for name in (
        "tab",
        "br",
        "cr",
        "noBreakHyphen",
        "softHyphen",
        "sym",
        "ptab",
        "drawing",
        "pict",
        "object",
        "footnoteReference",
        "endnoteReference",
        "annotationRef",
        "footnoteRef",
        "endnoteRef",
        "separator",
        "continuationSeparator",
        "pgNum",
    )
} | {qn("m:oMath"), qn("m:oMathPara")}
_INACTIVE_REVISIONS = {qn("w:del"), qn("w:moveFrom")}
_TEXTBOX = qn("w:txbxContent")


class TemplatePlaceholderOwnershipError(TemplatePackageError):
    """A marker would edit text owned by a field or a bound content control."""

    diagnostic_code = "MD2DOCX-TEMPLATE-PLACEHOLDER-OWNERSHIP"

    def __init__(
        self,
        *,
        part_name: str,
        paragraph_index: int,
        key: str,
        start: int,
        end: int,
        owner: str,
    ) -> None:
        self.part_name = part_name
        self.paragraph_index = paragraph_index
        self.key = key
        self.start = start
        self.end = end
        self.owner = owner
        super().__init__(
            f"模板占位符 {{{{{key}}}}} 位于{owner}中或跨越其边界"
            f"（部件 {part_name}，段落 {paragraph_index}，字符 {start + 1}–{end}）。"
            "请将占位符移到普通文本中；直接填写这里的显示值会被 Word 更新覆盖。"
        )


@dataclass(frozen=True, slots=True)
class TemplatePlaceholderSpan:
    """An original marker and the exact text-node positions it occupies."""

    key: str
    start: int
    end: int
    locations: tuple[tuple[Any, int], ...]


@dataclass(slots=True)
class _ParagraphText:
    text: str
    locations: list[tuple[Any, int] | None]
    boundaries: dict[Any, tuple[int, int]]

    def spans(self) -> Iterator[TemplatePlaceholderSpan]:
        for match in _PLACEHOLDER.finditer(self.text):
            positions = self.locations[match.start() : match.end()]
            if any(position is None for position in positions):
                continue
            yield TemplatePlaceholderSpan(
                key=match.group(1).strip(),
                start=match.start(),
                end=match.end(),
                locations=tuple(position for position in positions if position is not None),
            )


def _read_paragraph_text(element: Any) -> _ParagraphText:
    if any(ancestor.tag in _INACTIVE_REVISIONS for ancestor in element.iterancestors()):
        return _ParagraphText("", [], {})
    characters: list[str] = []
    locations: list[tuple[Any, int] | None] = []
    boundaries: dict[Any, tuple[int, int]] = {}

    def visit(node: Any) -> None:
        if node.tag in _INACTIVE_REVISIONS or (node is not element and node.tag == _PARAGRAPH):
            # A drawing's inner paragraph is a different substitution location.
            return
        start = len(characters)
        if node.tag == _TEXT:
            for offset, character in enumerate(node.text or ""):
                characters.append(character)
                locations.append((node, offset))
        elif node.tag in _CONTROLS:
            characters.append("\0")
            locations.append(None)
        else:
            for child in node:
                visit(child)
        boundaries[node] = (start, len(characters))

    visit(element)
    return _ParagraphText("".join(characters), locations, boundaries)


def iter_paragraph_placeholder_spans(paragraph: Any) -> Iterator[TemplatePlaceholderSpan]:
    """Read immutable original marker spans from a Paragraph or ``w:p`` node.

    Visible non-text run content cannot be part of a marker. Deleted content
    does not belong to the live text, and textbox paragraphs are scanned
    separately. Consumers must collect all spans before editing and replace
    them right-to-left so inserted values stay literal.
    """
    element = getattr(paragraph, "_element", paragraph)
    yield from _read_paragraph_text(element).spans()


@dataclass(slots=True)
class _InstructionText:
    characters: list[str] = field(default_factory=list)
    sources: list[Any] = field(default_factory=list)

    def append(self, node: Any) -> None:
        value = node.text or ""
        self.characters.extend(value)
        self.sources.extend([node] * len(value))

    def separate(self, node: Any) -> None:
        self.characters.append("\0")
        self.sources.append(node)


def validate_template_placeholder_ownership(doc: Any) -> None:
    """Reject field/SDT markers in the stories that template filling scans.

    Related headers and footers are inspected without accessing Section.header,
    which could otherwise create new package parts. No XML is changed.
    """
    parts = [doc.part]
    parts.extend(
        part
        for part in doc.part.related_parts.values()
        if hasattr(part, "element") and str(part.partname).startswith(("/word/header", "/word/footer"))
    )
    seen: set[str] = set()
    for part in parts:
        part_name = str(part.partname)
        if part_name not in seen:
            _validate_part(part.element, part_name)
            seen.add(part_name)


def _validate_part(root: Any, part_name: str) -> None:
    paragraphs = list(root.iter(_PARAGRAPH))
    paragraph_indices = {paragraph: index for index, paragraph in enumerate(paragraphs, 1)}
    paragraph_text = {paragraph: _read_paragraph_text(paragraph) for paragraph in paragraphs}
    owned_nodes: dict[Any, str] = {}
    boundary_nodes: list[tuple[Any, str]] = []

    def paragraph_of(node: Any) -> Any:
        current = node
        while current is not None and current.tag != _PARAGRAPH:
            current = current.getparent()
        if current is None:
            current = next(node.iter(_PARAGRAPH), None)
        return current

    def reject(key: str, start: int, end: int, node: Any, owner: str) -> None:
        raise TemplatePlaceholderOwnershipError(
            part_name=part_name,
            paragraph_index=paragraph_indices.get(paragraph_of(node), 0),
            key=key,
            start=start,
            end=end,
            owner=owner,
        )

    def check_instruction(instruction: _InstructionText) -> None:
        for match in _PLACEHOLDER.finditer("".join(instruction.characters)):
            if "\0" in match.group(0):
                continue
            reject(
                match.group(1).strip(),
                match.start(),
                match.end(),
                instruction.sources[match.start()],
                "Word 域指令",
            )

    def check_attribute(node: Any, value: str, owner: str) -> None:
        match = _PLACEHOLDER.search(value)
        if match is not None:
            reject(match.group(1).strip(), match.start(), match.end(), node, owner)

    def visit_story(story: Any, inherited_owners: tuple[str, ...] = ()) -> None:
        fields: list[_InstructionText] = []
        orphan_instructions: dict[Any, _InstructionText] = {}

        def visit(node: Any, owners: tuple[str, ...]) -> None:
            if node.tag in _INACTIVE_REVISIONS:
                # Deleted/moved-from field delimiters are history, not the
                # current story's live field stack. Keep that history intact.
                return
            instruction_boundary = node.tag in {
                _TEXT,
                _TEXTBOX,
                qn("w:fldChar"),
                qn("w:fldSimple"),
                qn("w:sdt"),
                *_CONTROLS,
            }
            if instruction_boundary:
                orphan = orphan_instructions.get(paragraph_of(node))
                if orphan is not None:
                    orphan.separate(node)
            if node is not story and node.tag == _TEXTBOX:
                # Fields can span paragraphs but never cross Word story borders.
                # A generated shape is still dynamic content even though its
                # internal field stack cannot consume the containing field's end.
                inherited = (*owners, "Word 复杂域") if fields else owners
                visit_story(node, inherited)
                return
            if node.tag == qn("w:fldSimple"):
                owners = (*owners, "Word 简单域")
                boundary_nodes.append((node, owners[-1]))
                check_attribute(node, node.get(qn("w:instr"), ""), "Word 域指令")
            elif node.tag == qn("w:sdt"):
                binding = node.find(qn("w:sdtPr") + "/" + qn("w:dataBinding"))
                if binding is not None:
                    owners = (*owners, "绑定内容控件")
                    boundary_nodes.append((node, owners[-1]))
                    for value in binding.attrib.values():
                        check_attribute(node, value, "内容控件数据绑定")
            elif node.tag == qn("w:fldChar"):
                boundary_nodes.append((node, "Word 复杂域"))
                kind = node.get(qn("w:fldCharType"))
                if kind == "begin":
                    if fields:
                        fields[-1].separate(node)
                    fields.append(_InstructionText())
                elif kind == "end" and fields:
                    check_instruction(fields.pop())
                    if fields:
                        fields[-1].separate(node)
                elif kind == "separate" and fields:
                    fields[-1].separate(node)
            elif node.tag == qn("w:instrText"):
                boundary_nodes.append((node, "Word 域指令"))
                if fields:
                    fields[-1].append(node)
                else:
                    orphan_instructions.setdefault(paragraph_of(node), _InstructionText()).append(node)
            elif node.tag == _TEXT:
                if owners:
                    owned_nodes[node] = owners[-1]
                elif fields:
                    owned_nodes[node] = "Word 复杂域"
            for child in node:
                visit(child, owners)
            if instruction_boundary:
                orphan = orphan_instructions.get(paragraph_of(node))
                if orphan is not None:
                    orphan.separate(node)

        visit(story, inherited_owners)
        for instruction in (*fields, *orphan_instructions.values()):
            check_instruction(instruction)

    visit_story(root)
    boundaries_by_paragraph: dict[Any, list[tuple[int, str]]] = {}
    for node, owner in boundary_nodes:
        paragraph = paragraph_of(node)
        if paragraph in paragraph_text:
            offsets = paragraph_text[paragraph].boundaries.get(node)
            if offsets is not None:
                boundaries_by_paragraph.setdefault(paragraph, []).extend((offset, owner) for offset in offsets)
    for paragraph, text in paragraph_text.items():
        for span in text.spans():
            owner = next((owned_nodes[node] for node, _ in span.locations if node in owned_nodes), None)
            if owner is None:
                owner = next(
                    (
                        name
                        for offset, name in boundaries_by_paragraph.get(paragraph, ())
                        if span.start < offset < span.end
                    ),
                    None,
                )
            if owner is not None:
                reject(span.key, span.start, span.end, paragraph, owner)

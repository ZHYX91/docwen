"""Footnotes and endnotes extraction, ID mapping, inline references, ZIP fallback."""

from __future__ import annotations

import re
from collections.abc import Callable
from typing import Any
from zipfile import ZipFile

from docx.opc.constants import RELATIONSHIP_TYPE as RT
from lxml import etree

from docwen_core.docx_bookmarks import build_docx_bookmark_inventory, prove_bookmark_name
from docwen_core.docx_parsing.format_features import DocxMarkdownSyntaxConfig, StyleDetectorConfig
from docwen_core.docx_parsing.xml_ns import NS_W
from docwen_plugin_document.shared.markdown_runs import (
    _run_is_hidden,
    render_paragraph_runs,
)
from docwen_plugin_document.shared.note_target_structure import NOTE_WRAPPERS, is_pure_note_target


def _extract_notes_with_status(
    doc,
    docx_path: str | None,
    part_name: str,
    note_tag: str,
    *,
    preserve_formatting: bool = True,
    syntax_config: DocxMarkdownSyntaxConfig | None = None,
    style_detector_config: StyleDetectorConfig | None = None,
    math_renderer: Callable[[Any, str], str] | None = None,
) -> tuple[dict[int, str], bool]:
    """Extract notes from python-docx part or ZIP fallback.

    *ref_tag* is derived by replacing trailing ``"s"`` (e.g. ``"footnotes"``
    → ``"footnote"`` → ``"footnoteRef"``, or manually for ``"endnoteRef"``).
    """
    # Derive ref tag: "footnotes" → "footnoteRef", "endnotes" → "endnoteRef"
    ref_tag = note_tag.rstrip("s") + "Ref"

    elem = None
    source_part = f"word/{part_name}.xml"
    package_part_failed = False
    try:
        part = getattr(doc.part, f"{part_name}_part", None)
        if part is None:
            relation_type = RT.FOOTNOTES if part_name == "footnotes" else RT.ENDNOTES
            part = next(
                (
                    relation.target_part
                    for relation in getattr(doc.part, "rels", {}).values()
                    if relation.reltype == relation_type and not relation.is_external
                ),
                None,
            )
        if part is not None:
            source_part = str(part.partname).lstrip("/")
            elem = getattr(part, "element", None)
            if elem is None:
                elem = etree.fromstring(part.blob)
    except Exception:
        package_part_failed = True

    if elem is None and docx_path:
        elem, zip_part_failed = _load_notes_elem_from_zip(docx_path, part_name)
        if elem is not None:
            source_part = f"word/{part_name}.xml"
            package_part_failed = False
        elif zip_part_failed:
            package_part_failed = True

    if elem is None:
        return {}, package_part_failed

    notes: dict[int, str] = {}
    w_ns = NS_W
    for note_elem in elem.findall(f"{{{w_ns}}}{note_tag}"):
        w_id_raw = note_elem.get(f"{{{w_ns}}}id")
        if w_id_raw is None:
            continue
        if _is_system_note(note_elem, w_ns):
            continue
        content = _extract_note_content(
            note_elem,
            w_ns,
            ref_tag,
            preserve_formatting=preserve_formatting,
            syntax_config=syntax_config,
            style_detector_config=style_detector_config,
            style_parent=doc,
            math_renderer=math_renderer,
            source_part=source_part,
        )
        # A present empty definition is not unreadable and has lost no text.
        # Do not conflate it with an unsupported nontext-only note payload.
        nontext_payload = any(
            etree.QName(node).localname in {"drawing", "pict", "object", "tbl", "sym", "oMath", "oMathPara"}
            for node in note_elem.iter()
            if isinstance(node.tag, str)
            and not any(
                ancestor.tag in {f"{{{w_ns}}}del", f"{{{w_ns}}}moveFrom"}
                or (ancestor.tag == f"{{{w_ns}}}r" and _run_is_hidden(ancestor, doc))
                for ancestor in node.iterancestors()
            )
        )
        if content.strip() or not nontext_payload:
            notes[int(w_id_raw)] = content
    return notes, False


def _load_notes_elem_from_zip(docx_path: str, part_name: str) -> tuple[Any | None, bool]:
    """Load one notes element and report whether a present part was unreadable."""
    target = f"word/{part_name}.xml"
    try:
        with ZipFile(docx_path, "r") as zf:
            if target not in zf.namelist():
                return None, False
            data = zf.read(target)
        return etree.fromstring(data), False
    except Exception:
        return None, True


def _is_system_note(elem, w_ns: str) -> bool:
    ntype = elem.get(f"{{{w_ns}}}type")
    return ntype in ("separator", "continuationSeparator")


def _extract_note_content(
    elem,
    w_ns: str,
    ref_tag: str,
    *,
    preserve_formatting: bool = True,
    syntax_config: DocxMarkdownSyntaxConfig | None = None,
    style_detector_config: StyleDetectorConfig | None = None,
    style_parent: Any = None,
    math_renderer: Callable[[Any, str], str] | None = None,
    source_part: str | None = None,
) -> str:
    """Read notes in XML order through the body text and formula renderer."""

    def render_math(element: Any) -> str:
        if math_renderer is not None:
            part = source_part or f"word/{ref_tag.removesuffix('Ref')}s.xml"
            return math_renderer(element, part)
        from docwen_plugin_document.to_markdown.formula_extractor import extract_formula_from_element

        return extract_formula_from_element(element)

    para_texts: list[str] = []
    for child in elem:
        tag = child.tag.split("}")[-1] if "}" in (child.tag or "") else (child.tag or "")
        if tag != "p":
            continue
        # Keep the original XML attached to its part, so formula diagnostics
        # retain their exact occurrence path. Suppress only the native marker
        # and its adjacent writer-owned separator through the shared walker.
        marker_spans: dict[Any, tuple[Any, str]] = {}
        for run in child.iter(f"{{{w_ns}}}r"):
            if any(ancestor.tag in {f"{{{w_ns}}}del", f"{{{w_ns}}}moveFrom"} for ancestor in run.iterancestors()):
                continue
            if _run_is_hidden(run, style_parent):
                continue
            if run.find(f"{{{w_ns}}}{ref_tag}") is None:
                continue
            marker_spans[run] = (run, "")
            separator = run.getnext()
            if (
                separator is not None
                and separator.tag == f"{{{w_ns}}}r"
                and _is_reference_separator_run(separator, w_ns)
            ):
                marker_spans[separator] = (separator, "")
        text = render_paragraph_runs(
            child,
            preserve_formatting=preserve_formatting,
            syntax_config=syntax_config or DocxMarkdownSyntaxConfig(),
            style_detector_config=style_detector_config,
            style_parent=style_parent,
            math_renderer=render_math,
            reference_spans=marker_spans,
        )
        if text:
            para_texts.append(text)
    # Paragraph boundaries must not become soft line breaks on the next
    # Markdown import. Breaks within a paragraph remain single newlines.
    return "\n\n".join(para_texts)


def _is_reference_separator_run(run: Any, w_ns: str) -> bool:
    """Recognize the one writer-owned space after a note reference mark.

    The separator is structurally distinct from authored text: it is one
    adjacent run containing only one preserve-space ``w:t``. Authored leading
    whitespace in the following content run and later paragraphs remains
    untouched.
    """

    children = list(run)
    if len(children) != 1 or children[0].tag != f"{{{w_ns}}}t":
        return False
    text = children[0]
    return text.text == " " and text.get("{http://www.w3.org/XML/1998/namespace}space") == "preserve"


def build_note_definitions(
    notes: dict[int, str],
    id_map: dict[int, str],
) -> str:
    """Build Markdown definition block for notes.

    Args:
        notes: {word_id: content}
        id_map: {word_id: display_id} — e.g. 5→"1" or 9→"endnote:1"

    Returns:
        Markdown definition lines with ``[^id]: content`` format.
    """
    lines: list[str] = []

    def display_order(word_id: int) -> tuple[int, int]:
        display_id = id_map.get(word_id, str(word_id))
        suffix = display_id.rsplit(":", 1)[-1]
        return (int(suffix), word_id) if suffix.isdigit() else (2**31 - 1, word_id)

    for word_id in sorted(notes.keys(), key=display_order):
        display_id = id_map.get(word_id, str(word_id))
        content = notes[word_id]
        rendered = _format_multiline_content(content)
        lines.append(f"[^{display_id}]: {rendered}")
    return "\n".join(lines)


def _format_multiline_content(content: str) -> str:
    """Wrap continuation lines with 4-space indentation."""
    parts = content.split("\n")
    if len(parts) <= 1:
        return content
    return parts[0] + "\n" + "\n".join(f"    {p}" for p in parts[1:])


def _native_marker_is_run_payload(reference: Any) -> bool:
    q = f"{{{NS_W}}}"
    run = reference.getparent()
    if run is None or run.tag != q + "r":
        return False
    for ancestor in run.iterancestors():
        if ancestor.tag == q + "p":
            return True
        if ancestor.tag not in NOTE_WRAPPERS:
            return False
    return False


def _note_bookmark_targets(doc) -> dict[str, tuple[str, int]]:
    """Resolve unique balanced bookmarks containing exactly one note marker."""
    inventory = build_docx_bookmark_inventory(doc)
    elements = list(doc.element.iter())
    positions = {id(element): index for index, element in enumerate(elements)}
    targets: dict[str, tuple[str, int]] = {}
    for start in inventory.starts:
        if start.name is None or start.part_name != str(doc.part.partname):
            continue
        proof = prove_bookmark_name(inventory, start.name)
        if not proof.valid or proof.end is None:
            continue
        begin = positions.get(id(start.element))
        end = positions.get(id(proof.end.element))
        if begin is None or end is None:
            continue
        contents = elements[begin + 1 : end]
        references = [
            element
            for element in contents
            if element.tag in {f"{{{NS_W}}}footnoteReference", f"{{{NS_W}}}endnoteReference"}
        ]
        if len(references) != 1 or any(element.tag == f"{{{NS_W}}}t" and element.text for element in contents):
            continue
        reference = references[0]
        q = f"{{{NS_W}}}"
        if not _native_marker_is_run_payload(reference):
            continue
        if any(
            ancestor.tag in {q + "del", q + "moveFrom"} or (ancestor.tag == q + "r" and _run_is_hidden(ancestor, doc))
            for ancestor in reference.iterancestors()
        ):
            continue
        # A target is the marker itself, not an arbitrary visible range that
        # happens to contain one. Run formatting contributes no body payload.
        if not is_pure_note_target(contents, reference):
            continue
        raw_id = reference.get(f"{{{NS_W}}}id", "")
        if not raw_id.isdecimal() or int(raw_id) <= 0:
            continue
        kind = "footnote" if reference.tag == f"{{{NS_W}}}footnoteReference" else "endnote"
        targets[start.name.casefold()] = (kind, int(raw_id))
    return targets


class NoteExtractor:
    """Aggregate footnote/endnote extraction, mapping, reference text, and
    Markdown definitions block."""

    def __init__(
        self,
        doc,
        docx_path: str | None = None,
        *,
        typed_endnotes: bool = True,
        preserve_formatting: bool = True,
        syntax_config: DocxMarkdownSyntaxConfig | None = None,
        style_detector_config: StyleDetectorConfig | None = None,
        math_renderer: Callable[[Any, str], str] | None = None,
        missing_placeholder: str = "[Note could not be read; check the source document]",
    ) -> None:
        self._missing_placeholder = missing_placeholder
        self._endnote_prefix = "endnote:" if typed_endnotes else "endnote-"
        self.footnotes, self.footnote_part_failed = _extract_notes_with_status(
            doc,
            docx_path,
            "footnotes",
            "footnote",
            preserve_formatting=preserve_formatting,
            syntax_config=syntax_config,
            style_detector_config=style_detector_config,
            math_renderer=math_renderer,
        )
        self.endnotes, self.endnote_part_failed = _extract_notes_with_status(
            doc,
            docx_path,
            "endnotes",
            "endnote",
            preserve_formatting=preserve_formatting,
            syntax_config=syntax_config,
            style_detector_config=style_detector_config,
            math_renderer=math_renderer,
        )
        # Display IDs are assigned lazily from the first body reference.
        # Footnotes and endnotes own independent per-file domains.
        self.footnote_id_map: dict[int, str] = {}
        self.endnote_id_map: dict[int, str] = {}
        self._referenced_footnote_ids: set[int] = set()
        self._referenced_endnote_ids: set[int] = set()
        self._note_bookmarks = _note_bookmark_targets(doc)

    def get_noteref_text(self, instruction: str) -> str | None:
        """Recover only number-valued fields pointing at a proven native note."""
        match = re.fullmatch(
            r'\s*NOTEREF\s+(?:"([A-Za-z_][A-Za-z0-9_]{0,39})"|([A-Za-z_][A-Za-z0-9_]{0,39}))(?:\s+\\[hf])*\s*',
            instruction,
            re.IGNORECASE,
        )
        if match is None:
            return None
        target = self._note_bookmarks.get((match[1] or match[2]).casefold())
        return self.get_reference_text(*target) if target is not None else None

    def get_reference_text(self, ref_type: str, word_id: int) -> str:
        """Return an inline reference numbered by first use in its note domain."""
        if ref_type == "footnote":
            referenced = getattr(self, "_referenced_footnote_ids", None)
            if referenced is None:
                referenced = set()
                self._referenced_footnote_ids = referenced
            referenced.add(word_id)
            display_id = self.footnote_id_map.get(word_id)
            if display_id is None:
                display_id = str(len(self.footnote_id_map) + 1)
                self.footnote_id_map[word_id] = display_id
        elif ref_type == "endnote":
            referenced = getattr(self, "_referenced_endnote_ids", None)
            if referenced is None:
                referenced = set()
                self._referenced_endnote_ids = referenced
            referenced.add(word_id)
            display_id = self.endnote_id_map.get(word_id)
            if display_id is None:
                display_id = f"{getattr(self, '_endnote_prefix', 'endnote:')}{len(self.endnote_id_map) + 1}"
                self.endnote_id_map[word_id] = display_id
        else:
            display_id = str(word_id)
        return f"[^{display_id}]"

    def missing_definitions(self) -> tuple[tuple[str, int], ...]:
        """Return referenced source notes whose definitions were unreadable."""
        return tuple(
            (kind, note_id)
            for kind, notes, referenced in (
                ("footnote", self.footnotes, getattr(self, "_referenced_footnote_ids", set())),
                ("endnote", self.endnotes, getattr(self, "_referenced_endnote_ids", set())),
            )
            for note_id in sorted(referenced - notes.keys())
        )

    def definition_loss_counts(self) -> dict[str, int]:
        """Return referenced note definitions lost because their part failed to load."""
        footnote_ids = getattr(self, "_referenced_footnote_ids", set())
        endnote_ids = getattr(self, "_referenced_endnote_ids", set())
        return {
            "footnotes": len(footnote_ids - self.footnotes.keys()),
            "endnotes": len(endnote_ids - self.endnotes.keys()),
        }

    def build_definitions_block(self) -> str:
        """Build combined Markdown definitions block."""
        for word_id in sorted(self.footnotes):
            if word_id not in self.footnote_id_map:
                self.footnote_id_map[word_id] = str(len(self.footnote_id_map) + 1)
        for word_id in sorted(self.endnotes):
            if word_id not in self.endnote_id_map:
                self.endnote_id_map[word_id] = (
                    f"{getattr(self, '_endnote_prefix', 'endnote:')}{len(self.endnote_id_map) + 1}"
                )
        parts: list[str] = []
        if self.footnotes:
            parts.append(build_note_definitions(self.footnotes, self.footnote_id_map))
        if self.endnotes:
            parts.append(build_note_definitions(self.endnotes, self.endnote_id_map))
        for notes, mapping, referenced in (
            (self.footnotes, self.footnote_id_map, getattr(self, "_referenced_footnote_ids", set())),
            (self.endnotes, self.endnote_id_map, getattr(self, "_referenced_endnote_ids", set())),
        ):
            for word_id in sorted(referenced - notes.keys()):
                parts.append(f"[^{mapping[word_id]}]: {self._missing_placeholder}")
        return "\n".join(p for p in parts if p)

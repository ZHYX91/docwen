"""Footnotes and endnotes extraction, ID mapping, inline references, ZIP fallback."""

from __future__ import annotations

import re
from typing import Any
from zipfile import ZipFile

from docwen_core.docx_bookmarks import build_docx_bookmark_inventory, prove_bookmark_name
from docwen_core.docx_parsing.format_features import DocxMarkdownSyntaxConfig, StyleDetectorConfig
from docwen_core.docx_parsing.xml_ns import NS_W
from docwen_plugin_document.shared.markdown_runs import (
    _run_is_hidden,
    append_formatted_run_text,
    resolve_run_style_type,
)


def _extract_notes_with_status(
    doc,
    docx_path: str | None,
    part_name: str,
    note_tag: str,
    *,
    preserve_formatting: bool = True,
    syntax_config: DocxMarkdownSyntaxConfig | None = None,
    style_detector_config: StyleDetectorConfig | None = None,
) -> tuple[dict[int, str], bool]:
    """Extract notes from python-docx part or ZIP fallback.

    *ref_tag* is derived by replacing trailing ``"s"`` (e.g. ``"footnotes"``
    → ``"footnote"`` → ``"footnoteRef"``, or manually for ``"endnoteRef"``).
    """
    # Derive ref tag: "footnotes" → "footnoteRef", "endnotes" → "endnoteRef"
    ref_tag = note_tag.rstrip("s") + "Ref"

    elem = None
    package_part_failed = False
    try:
        part = getattr(doc.part, f"{part_name}_part", None)
        if part is not None:
            elem = part.element
    except Exception:
        package_part_failed = True

    if elem is None and docx_path:
        elem, zip_part_failed = _load_notes_elem_from_zip(docx_path, part_name)
        if elem is not None:
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
        )
        if content.strip():
            notes[int(w_id_raw)] = content
    return notes, False


def _load_notes_elem_from_zip(docx_path: str, part_name: str) -> tuple[Any | None, bool]:
    """Load one notes element and report whether a present part was unreadable."""
    import lxml.etree as etree

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
) -> str:
    """Extract note Markdown while preserving supported run formatting and breaks."""

    para_texts: list[str] = []
    for child in elem:
        tag = child.tag.split("}")[-1] if "}" in (child.tag or "") else (child.tag or "")
        if tag != "p":
            continue
        run_texts: list[str] = []
        separator_expected = False
        for run in child.iter(f"{{{w_ns}}}r"):
            if run.find(f"{{{w_ns}}}{ref_tag}") is not None:
                separator_expected = True
                continue
            if separator_expected and _is_reference_separator_run(run, w_ns):
                separator_expected = False
                continue
            separator_expected = False
            _append_note_run(
                run_texts,
                run,
                w_ns,
                preserve_formatting=preserve_formatting,
                syntax_config=syntax_config or DocxMarkdownSyntaxConfig(),
                style_detector_config=style_detector_config,
                style_parent=style_parent,
            )
        if run_texts:
            para_texts.append("".join(run_texts))
    return "\n".join(para_texts)


def _append_note_run(
    rendered: list[str],
    run: Any,
    w_ns: str,
    *,
    preserve_formatting: bool,
    syntax_config: DocxMarkdownSyntaxConfig,
    style_detector_config: StyleDetectorConfig | None,
    style_parent: Any,
) -> None:
    """Use the body renderer's syntax, code-span padding and run coalescing."""

    parts: list[str] = []
    for child in run:
        if child.tag == f"{{{w_ns}}}t":
            parts.append(child.text or "")
        elif child.tag == f"{{{w_ns}}}br":
            parts.append("\n")
        elif child.tag == f"{{{w_ns}}}tab":
            parts.append("\t")
    text = "".join(parts)
    for index, segment in enumerate(text.split("\n")):
        if index:
            rendered.append("\n")
        if not segment:
            continue
        if preserve_formatting:
            append_formatted_run_text(
                rendered,
                segment,
                run,
                syntax_config=syntax_config,
                style_detector_config=style_detector_config,
                run_style_type=resolve_run_style_type(run, style_parent, style_detector_config),
            )
        else:
            rendered.append(segment)


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
        if any(
            ancestor.tag in {q + "del", q + "moveFrom"} or (ancestor.tag == q + "r" and _run_is_hidden(ancestor))
            for ancestor in reference.iterancestors()
        ):
            continue
        # A target is the marker itself, not an arbitrary visible range that
        # happens to contain one. Run formatting contributes no body payload.
        containers = {q + name for name in ("r", "rPr", "ins", "moveTo", "bookmarkStart", "bookmarkEnd")}
        if any(
            element is not reference
            and element.tag not in containers
            and not any(ancestor.tag == q + "rPr" for ancestor in element.iterancestors())
            for element in contents
        ):
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
    ) -> None:
        self._endnote_prefix = "endnote:" if typed_endnotes else "endnote-"
        self.footnotes, self.footnote_part_failed = _extract_notes_with_status(
            doc,
            docx_path,
            "footnotes",
            "footnote",
            preserve_formatting=preserve_formatting,
            syntax_config=syntax_config,
            style_detector_config=style_detector_config,
        )
        self.endnotes, self.endnote_part_failed = _extract_notes_with_status(
            doc,
            docx_path,
            "endnotes",
            "endnote",
            preserve_formatting=preserve_formatting,
            syntax_config=syntax_config,
            style_detector_config=style_detector_config,
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

    def definition_loss_counts(self) -> dict[str, int]:
        """Return referenced note definitions lost because their part failed to load."""
        footnote_ids = getattr(self, "_referenced_footnote_ids", set())
        endnote_ids = getattr(self, "_referenced_endnote_ids", set())
        return {
            "footnotes": (
                len(footnote_ids - self.footnotes.keys()) if getattr(self, "footnote_part_failed", False) else 0
            ),
            "endnotes": (len(endnote_ids - self.endnotes.keys()) if getattr(self, "endnote_part_failed", False) else 0),
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
        return "\n".join(p for p in parts if p)

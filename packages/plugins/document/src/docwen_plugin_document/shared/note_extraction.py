"""Footnotes and endnotes extraction, ID mapping, inline references, ZIP fallback."""

from __future__ import annotations

from typing import Any
from zipfile import ZipFile

from docwen_core.docx_parsing.xml_ns import NS_W


def _extract_notes_with_status(
    doc,
    docx_path: str | None,
    part_name: str,
    note_tag: str,
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
        content = _extract_note_content(note_elem, w_ns, ref_tag)
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


def _extract_note_content(elem, w_ns: str, ref_tag: str) -> str:
    """Extract text content from a footnote/endnote element.

    Processes paragraphs individually, skipping runs that contain the
    reference marker element (e.g. ``footnoteRef`` or ``endnoteRef``)
    to avoid including the auto-numbering character in the content.
    Multi-paragraph notes are joined with newlines.
    """
    para_texts: list[str] = []
    for child in elem:
        tag = child.tag.split("}")[-1] if "}" in (child.tag or "") else (child.tag or "")
        if tag != "p":
            continue
        run_texts: list[str] = []
        separator_expected = False
        for run in child:
            run_tag = run.tag.split("}")[-1] if "}" in (run.tag or "") else (run.tag or "")
            if run_tag != "r":
                separator_expected = False
                continue
            if run.find(f"{{{w_ns}}}{ref_tag}") is not None:
                separator_expected = True
                continue
            if separator_expected and _is_reference_separator_run(run, w_ns):
                separator_expected = False
                continue
            separator_expected = False
            for t in run.findall(f"{{{w_ns}}}t"):
                if t.text:
                    run_texts.append(t.text)
        if run_texts:
            para_texts.append("".join(run_texts))
    return "\n".join(para_texts)


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


class NoteExtractor:
    """Aggregate footnote/endnote extraction, mapping, reference text, and
    Markdown definitions block."""

    def __init__(self, doc, docx_path: str | None = None) -> None:
        self.footnotes, self.footnote_part_failed = _extract_notes_with_status(
            doc,
            docx_path,
            "footnotes",
            "footnote",
        )
        self.endnotes, self.endnote_part_failed = _extract_notes_with_status(
            doc,
            docx_path,
            "endnotes",
            "endnote",
        )
        # Display IDs are assigned lazily from the first body reference.
        # Footnotes and endnotes own independent per-file domains.
        self.footnote_id_map: dict[int, str] = {}
        self.endnote_id_map: dict[int, str] = {}
        self._referenced_footnote_ids: set[int] = set()
        self._referenced_endnote_ids: set[int] = set()

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
                display_id = f"endnote:{len(self.endnote_id_map) + 1}"
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
                self.endnote_id_map[word_id] = f"endnote:{len(self.endnote_id_map) + 1}"
        parts: list[str] = []
        if self.footnotes:
            parts.append(build_note_definitions(self.footnotes, self.footnote_id_map))
        if self.endnotes:
            parts.append(build_note_definitions(self.endnotes, self.endnote_id_map))
        return "\n".join(p for p in parts if p)

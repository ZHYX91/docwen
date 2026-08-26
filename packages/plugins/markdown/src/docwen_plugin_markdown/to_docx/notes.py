"""Footnote/endnote foundation model for MD→DOCX conversion.

Provides extraction, ID mapping, OOXML element creation, and a
note-aware renderer that integrates mistune's footnotes plugin
output with python-docx.

Associated findings: F-F1-022, F-F1-023, F-F1-024, F-F1-027, F-F3-007
"""

from __future__ import annotations

import posixpath
import re
import unicodedata
from io import BytesIO
from pathlib import Path
from typing import Any, NoReturn
from zipfile import BadZipFile, ZipFile

import lxml.etree as etree

# ── OOXML constants ─────────────────────────────────────────────────────
WML_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"

# Mistune uppercases footnote keys; endnote prefix is always uppercase
# in the parsed AST.
ENDNOTE_PREFIX_UC = "ENDNOTE-"
_ENDNOTE_PFX_LEN = len(ENDNOTE_PREFIX_UC)

_NOTE_DEFINITION_RE = re.compile(r"^(?P<lead> {0,3})\[\^(?P<label>[^\]\\\s]+)\]:[ \t]*(?P<body>.*)$")
_NOTE_REFERENCE_RE = re.compile(r"(?<!\\)\[\^(?P<label>[^\]\\\s]+)\]")
_FENCE_OPEN_RE = re.compile(r"^ {0,3}(?P<fence>`{3,}|~{3,})(?P<info>[^\r\n]*)$")


def _note_syntax_invalid(message: str) -> NoReturn:
    raise NoteWritebackError(
        "MD2DOCX-NOTE-SYNTAX-INVALID",
        message,
        error_type="invalid_input",
    )


def _note_identity(label: str) -> tuple[str, str, str]:
    """Return ``(kind, normalized_id, spelling)`` for one authored label."""

    folded = label.casefold()
    if folded.startswith("footnote:"):
        kind = "footnote"
        note_id = label[len("footnote:") :]
        spelling = "explicit"
    elif folded.startswith("endnote:"):
        kind = "endnote"
        note_id = label[len("endnote:") :]
        spelling = "canonical"
    elif folded.startswith("endnote-"):
        _note_syntax_invalid(f"Unsupported note label '[^{label}]'; use the current endnote form '[^endnote:id]'.")
    else:
        kind = "footnote"
        note_id = label
        spelling = "default"

    if not note_id or any(char.isspace() or char in "[]\\" for char in note_id):
        _note_syntax_invalid(f"Invalid {kind} identifier in '[^{label}]'.")
    normalized_id = unicodedata.normalize("NFC", note_id).casefold()
    if not normalized_id:
        _note_syntax_invalid(f"Invalid {kind} identifier in '[^{label}]'.")
    return kind, normalized_id, spelling


def _rewrite_reference_segments(
    text: str,
    internal_keys: dict[tuple[str, str], str],
) -> str:
    """Rewrite note references outside inline-code spans."""

    def rewrite_segment(segment: str) -> str:
        def replace(match: re.Match[str]) -> str:
            kind, normalized_id, _spelling = _note_identity(match.group("label"))
            return f"[^{internal_keys[(kind, normalized_id)]}]"

        return _NOTE_REFERENCE_RE.sub(replace, segment)

    output: list[str] = []
    cursor = 0
    while cursor < len(text):
        tick = text.find("`", cursor)
        if tick < 0:
            output.append(rewrite_segment(text[cursor:]))
            break
        output.append(rewrite_segment(text[cursor:tick]))
        run_end = tick
        while run_end < len(text) and text[run_end] == "`":
            run_end += 1
        delimiter = text[tick:run_end]
        close = text.find(delimiter, run_end)
        if close < 0:
            output.append(text[tick:])
            break
        close_end = close + len(delimiter)
        output.append(text[tick:close_end])
        cursor = close_end
    return "".join(output)


def normalize_note_syntax(md_body: str) -> str:
    """Validate and normalize the frozen Obsidian note syntax for Mistune.

    The authored Markdown is never written back.  This request-local projection
    maps both footnote spellings to one internal key and maps the canonical
    endnote spelling to an endnote-prefixed internal key.  Definitions
    indented by two spaces or one tab are kept inside the note body.
    """

    lines = md_body.splitlines(keepends=True)
    fenced_lines: set[int] = set()
    definition_labels: dict[int, str] = {}
    continuation_leads: dict[int, int] = {}
    definitions: dict[tuple[str, str], tuple[str, str]] = {}
    reference_identities: list[tuple[str, str]] = []

    fence_char: str | None = None
    fence_length = 0
    active_definition_lead: int | None = None

    for index, line in enumerate(lines):
        text = line.rstrip("\r\n")

        if active_definition_lead is not None:
            if not text.strip():
                continuation_leads[index] = active_definition_lead
                continue
            lead = " " * active_definition_lead
            if text.startswith(f"{lead}\t"):
                continuation_leads[index] = active_definition_lead
                continue
            if text.startswith(lead):
                remainder = text[active_definition_lead:]
                spaces = len(remainder) - len(remainder.lstrip(" "))
                if spaces >= 2:
                    continuation_leads[index] = active_definition_lead
                    continue
            active_definition_lead = None

        if fence_char is not None:
            fenced_lines.add(index)
            if re.fullmatch(rf" {{0,3}}{re.escape(fence_char)}{{{fence_length},}}[ \t]*", text):
                fence_char = None
                fence_length = 0
            continue

        fence_match = _FENCE_OPEN_RE.match(text)
        if fence_match is not None:
            fence = fence_match.group("fence")
            fence_char = fence[0]
            fence_length = len(fence)
            fenced_lines.add(index)
            continue

        definition_match = _NOTE_DEFINITION_RE.match(text)
        if definition_match is None:
            continue

        label = definition_match.group("label")
        kind, normalized_id, spelling = _note_identity(label)
        identity = (kind, normalized_id)
        previous = definitions.get(identity)
        if previous is not None:
            previous_label, previous_spelling = previous
            if kind == "footnote" and previous_spelling != spelling:
                _note_syntax_invalid(
                    "Default and explicit footnote definitions normalize to the same identifier: "
                    f"'[^{previous_label}]' and '[^{label}]'."
                )
            _note_syntax_invalid(f"Duplicate note definition for '[^{label}]'.")
        definitions[identity] = (label, spelling)
        definition_labels[index] = label
        active_definition_lead = len(definition_match.group("lead"))

    for index, line in enumerate(lines):
        if index in fenced_lines or index in definition_labels or index in continuation_leads:
            continue
        text = line.rstrip("\r\n")
        cursor = 0
        while cursor < len(text):
            tick = text.find("`", cursor)
            segment_end = len(text) if tick < 0 else tick
            for match in _NOTE_REFERENCE_RE.finditer(text, cursor, segment_end):
                kind, normalized_id, _spelling = _note_identity(match.group("label"))
                reference_identities.append((kind, normalized_id))
            if tick < 0:
                break
            run_end = tick
            while run_end < len(text) and text[run_end] == "`":
                run_end += 1
            delimiter = text[tick:run_end]
            close = text.find(delimiter, run_end)
            if close < 0:
                break
            cursor = close + len(delimiter)

    for identity in reference_identities:
        if identity not in definitions:
            kind, normalized_id = identity
            _note_syntax_invalid(f"Missing {kind} definition for normalized identifier '{normalized_id}'.")

    ordered_identities: list[tuple[str, str]] = []
    for identity in [*reference_identities, *definitions.keys()]:
        if identity not in ordered_identities:
            ordered_identities.append(identity)
    internal_keys: dict[tuple[str, str], str] = {}
    for identity in ordered_identities:
        label, spelling = definitions[identity]
        if identity[0] == "footnote":
            note_id = label[len("footnote:") :] if spelling == "explicit" else label
            internal_keys[identity] = note_id
        else:
            note_id = label[len("endnote:") :]
            internal_keys[identity] = f"ENDNOTE-{note_id}"

    rewritten: list[str] = []
    for index, line in enumerate(lines):
        newline = line[len(line.rstrip("\r\n")) :]
        text = line.rstrip("\r\n")
        if index in definition_labels:
            match = _NOTE_DEFINITION_RE.match(text)
            assert match is not None
            kind, normalized_id, _spelling = _note_identity(definition_labels[index])
            start, end = match.span("label")
            text = text[:start] + internal_keys[(kind, normalized_id)] + text[end:]
        elif index in continuation_leads and text.strip():
            lead_length = continuation_leads[index]
            lead = text[:lead_length]
            remainder = text[lead_length:]
            if remainder.startswith("\t"):
                text = lead + "    " + remainder[1:]
            else:
                spaces = len(remainder) - len(remainder.lstrip(" "))
                if 2 <= spaces < 4:
                    text = lead + "    " + remainder[spaces:]
        elif index not in fenced_lines:
            text = _rewrite_reference_segments(text, internal_keys)
        rewritten.append(text + newline)
    return "".join(rewritten)


# ── Note extraction ─────────────────────────────────────────────────────


def _extract_plain_text(node: dict[str, Any]) -> str:
    """Extract plain text from an AST subtree (inline nodes)."""
    parts: list[str] = []
    ctype = node.get("type", "")
    if ctype == "text":
        parts.append(node.get("raw", "") or node.get("text", ""))
    elif ctype == "softbreak":
        parts.append(" ")
    elif ctype == "linebreak":
        parts.append("\n")
    children = node.get("children", [])
    for child in children:
        parts.append(_extract_plain_text(child))
    return "".join(parts)


def _extract_inline_children_per_para(
    footnote_item_children: list[dict[str, Any]],
) -> list[list[dict[str, Any]]]:
    """Extract a list of inline-children lists from footnote_item AST children.

    Each item in *footnote_item_children* is typically a paragraph node.
    Returns one entry per paragraph — a list of inline AST dicts
    (``text``, ``strong``, ``emphasis``, ``codespan``, …).
    """
    result: list[list[dict[str, Any]]] = []
    for child in footnote_item_children:
        if child.get("type") == "paragraph":
            result.append(child.get("children", []))
        else:
            # Non-paragraph child (edge case) — flatten to text
            text = _extract_plain_text(child)
            if text:
                result.append([{"type": "text", "raw": text}])
    return result if result else [[{"type": "text", "raw": ""}]]


def extract_notes_from_ast(
    ast_nodes: list[dict[str, Any]],
) -> tuple[list[dict[str, Any]], NoteContext]:
    """Extract footnote/endnote definitions from a mistune AST.

    Scans for the ``footnotes`` wrapper node, removes it, and populates
    a :class:`NoteContext` with per-paragraph inline children for each
    definition key.  Inline formatting (bold, italic, code, …) is
    preserved in the stored AST fragments.

    Returns:
        (cleaned_ast, note_context)
    """
    note_ctx = NoteContext()
    cleaned: list[dict[str, Any]] = []

    for node in ast_nodes:
        if node.get("type") == "footnotes":
            for item in node.get("children", []):
                if item.get("type") != "footnote_item":
                    continue
                attrs = item.get("attrs", {})
                key: str = attrs.get("key", "")
                if not key:
                    continue
                para_children = _extract_inline_children_per_para(item.get("children", []))
                if key.upper().startswith(ENDNOTE_PREFIX_UC):
                    clean_id = key[_ENDNOTE_PFX_LEN:]
                    note_ctx._endnote_children[clean_id] = para_children
                else:
                    note_ctx._footnote_children[key] = para_children
        else:
            cleaned.append(node)

    return cleaned, note_ctx


# ── Style name → NoteContext attribute name mapping ──────────────────────
_NOTE_STYLE_NAMES: dict[str, str] = {
    "footnote text": "footnote_text_style",
    "footnote reference": "footnote_ref_style",
    "endnote text": "endnote_text_style",
    "endnote reference": "endnote_ref_style",
}

_WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


class NoteWritebackError(RuntimeError):
    """Stable fail-closed error for an invalid note graph or atomic write."""

    def __init__(self, diagnostic_code: str, message: str, *, error_type: str) -> None:
        super().__init__(message)
        self.diagnostic_code = diagnostic_code
        self.error_type = error_type


def _lowest_unused_positive(used: set[int]) -> int:
    candidate = 1
    while candidate in used:
        candidate += 1
    return candidate


def _resolve_style_id_by_name(doc, style_name: str) -> str | None:
    """Find the **styleId** whose ``w:name`` equals *style_name* (case-insensitive).

    Iterates ``doc.styles`` (the python-docx public API). Returns ``None`` when
    not found, so the caller keeps the current English default.
    """
    target = style_name.lower()
    for style in doc.styles:
        style_elem = style._element
        name_elem = style_elem.find(f"{{{_WORD_NS}}}name")
        if name_elem is None:
            continue
        val = name_elem.get(f"{{{_WORD_NS}}}val", "")
        if val.lower() == target:
            sid = style_elem.get(f"{{{_WORD_NS}}}styleId")
            if sid:
                return sid
    return None


# ── NoteContext ──────────────────────────────────────────────────────────


class NoteContext:
    """Manages footnote/endnote ID mapping and element creation.

    Maps markdown note identifiers to Word footnote/endnote IDs.
    Stores per-paragraph inline AST children for formatting-preserving
    OOXML body element creation, and created elements for downstream
    consumption (e.g. writing to ``footnotes.xml`` / ``endnotes.xml``).
    """

    def __init__(self) -> None:
        # Markdown key → list of per-paragraph inline AST children
        # (each paragraph is list[dict] of inline nodes like text,
        #  strong, emphasis, codespan, …)
        self._footnote_children: dict[str, list[list[dict[str, Any]]]] = {}
        self._endnote_children: dict[str, list[list[dict[str, Any]]]] = {}

        # Markdown key → Word ID
        self._footnote_id_map: dict[str, int] = {}
        self._endnote_id_map: dict[str, int] = {}

        # Word ID counters (Word uses 1-based IDs)
        self._footnote_counter: int = 1
        self._endnote_counter: int = 1
        self._footnote_used_ids: set[int] = set()
        self._endnote_used_ids: set[int] = set()

        # Created OOXML body elements (for downstream write-back)
        self.footnote_elements: list[etree._Element] = []
        self.endnote_elements: list[etree._Element] = []

        # Style IDs (defaults; can be overridden per document)
        self.footnote_text_style: str = "FootnoteText"
        self.footnote_ref_style: str = "FootnoteReference"
        self.endnote_text_style: str = "EndnoteText"
        self.endnote_ref_style: str = "EndnoteReference"

    def resolve_note_styles(self, doc) -> None:
        """Resolve footnote/endnote style IDs from *doc*'s styles part.

        For a blank ``Document()`` the built-in IDs are the English defaults
        (already set in ``__init__``), so unmatched names keep the correct
        fallback.  For template-based documents this picks up custom styleIds
        (e.g. Chinese Word / WPS using ``'a5'`` instead of ``'FootnoteText'``).
        """
        for style_name, attr_name in _NOTE_STYLE_NAMES.items():
            resolved = _resolve_style_id_by_name(doc, style_name)
            if resolved is not None:
                setattr(self, attr_name, resolved)

    # ── ID mapping ──────────────────────────────────────────────────

    def reserve_existing_ids(
        self,
        *,
        footnote_ids: set[int],
        endnote_ids: set[int],
    ) -> None:
        """Reserve every existing positive note ID before body rendering."""

        if self._footnote_id_map or self._endnote_id_map or self.footnote_elements or self.endnote_elements:
            raise NoteWritebackError(
                "MD2DOCX-NOTE-PREPARE-ORDER",
                "Existing note IDs must be audited before any new note body or reference is created.",
                error_type="conversion_failed",
            )
        self._footnote_used_ids.update(footnote_ids)
        self._endnote_used_ids.update(endnote_ids)
        self._footnote_counter = _lowest_unused_positive(self._footnote_used_ids)
        self._endnote_counter = _lowest_unused_positive(self._endnote_used_ids)

    def get_footnote_word_id(self, md_key: str) -> int | None:
        """Return the Word ID for a footnote markdown key.

        If the key has not been seen before, a new Word ID is allocated
        and a footnote body element is created and stored.
        """
        if md_key not in self._footnote_children:
            return None

        if md_key not in self._footnote_id_map:
            word_id = self._footnote_counter
            self._footnote_id_map[md_key] = word_id
            self._footnote_used_ids.add(word_id)
            self._footnote_counter = _lowest_unused_positive(self._footnote_used_ids)

            children = self._footnote_children[md_key]
            elem = _create_footnote_element(
                word_id,
                children,
                self.footnote_text_style,
                self.footnote_ref_style,
            )
            self.footnote_elements.append(elem)

        return self._footnote_id_map[md_key]

    def get_endnote_word_id(self, md_key: str) -> int | None:
        """Return the Word ID for an endnote markdown key.

        Accepts keys with or without the ``ENDNOTE-`` prefix.
        """
        # Normalise: strip prefix if present
        clean_key = md_key[_ENDNOTE_PFX_LEN:] if md_key.upper().startswith(ENDNOTE_PREFIX_UC) else md_key

        if clean_key not in self._endnote_children:
            return None

        if clean_key not in self._endnote_id_map:
            word_id = self._endnote_counter
            self._endnote_id_map[clean_key] = word_id
            self._endnote_used_ids.add(word_id)
            self._endnote_counter = _lowest_unused_positive(self._endnote_used_ids)

            children = self._endnote_children[clean_key]
            elem = _create_endnote_element(
                word_id,
                children,
                self.endnote_text_style,
                self.endnote_ref_style,
            )
            self.endnote_elements.append(elem)

        return self._endnote_id_map[clean_key]

    # ── Reference run creation ──────────────────────────────────────

    def create_footnote_ref_run(self, md_key: str) -> etree._Element | None:
        """Create an OOXML ``w:r`` run for an inline footnote reference.

        Returns ``None`` if the footnote key is not defined.
        """
        word_id = self.get_footnote_word_id(md_key)
        if word_id is None:
            return None
        return _create_footnote_ref_run(word_id, self.footnote_ref_style)

    def create_endnote_ref_run(self, md_key: str) -> etree._Element | None:
        """Create an OOXML ``w:r`` run for an inline endnote reference.

        Returns ``None`` if the endnote key is not defined.
        """
        word_id = self.get_endnote_word_id(md_key)
        if word_id is None:
            return None
        return _create_endnote_ref_run(word_id, self.endnote_ref_style)

    # ── Convenience queries ─────────────────────────────────────────

    @property
    def has_footnotes(self) -> bool:
        return len(self._footnote_children) > 0

    @property
    def has_endnotes(self) -> bool:
        return len(self._endnote_children) > 0

    @property
    def has_notes(self) -> bool:
        return self.has_footnotes or self.has_endnotes


# ── OOXML element factories ─────────────────────────────────────────────


# ── Inline formatting context keys (used as a bit-set of str flags) ──────
# fmt dict may contain any of these str keys mapped to True.
_DC_BOLD = "b"
_DC_ITALIC = "i"
_DC_STRIKE = "strike"
_DC_SUPERSCRIPT = "superscript"
_DC_SUBSCRIPT = "subscript"
_DC_UNDERLINE = "underline"
_DC_HIGHLIGHT = "highlight"


def _render_note_inlines(
    para_elem: etree._Element,
    children: list[dict[str, Any]],
    fmt: dict[str, bool] | None = None,
) -> None:
    """Render inline AST nodes as OOXML ``w:r`` runs into *para_elem*.

    *fmt* accumulates inherited formatting (bold, italic, …) from
    ancestor wrapper nodes.  Leaf ``text`` nodes produce a single run
    with the combined properties; wrapper nodes recurse with an
    augmented context so that nested formatting is preserved.

    Supports: text, strong, emphasis, strikethrough, codespan,
    softbreak, linebreak, superscript, subscript, underline (insert),
    highlight (mark), and arbitrary nested children.
    """
    _f = fmt or {}
    for child in children:
        ctype = child.get("type", "")
        if ctype == "text":
            _nr_text(para_elem, child, _f)
        elif ctype == "strong":
            _render_note_inlines(para_elem, child.get("children", []), _fmt_set(_f, _DC_BOLD))
        elif ctype == "emphasis":
            _render_note_inlines(para_elem, child.get("children", []), _fmt_set(_f, _DC_ITALIC))
        elif ctype == "strikethrough":
            _render_note_inlines(para_elem, child.get("children", []), _fmt_set(_f, _DC_STRIKE))
        elif ctype == "codespan":
            _nr_codespan(para_elem, child)
        elif ctype == "softbreak":
            _nr_softbreak(para_elem)
        elif ctype == "linebreak":
            _nr_linebreak(para_elem)
        elif ctype == "superscript":
            _render_note_inlines(para_elem, child.get("children", []), _fmt_set(_f, _DC_SUPERSCRIPT))
        elif ctype == "subscript":
            _render_note_inlines(para_elem, child.get("children", []), _fmt_set(_f, _DC_SUBSCRIPT))
        elif ctype in ("underline", "insert"):
            _render_note_inlines(para_elem, child.get("children", []), _fmt_set(_f, _DC_UNDERLINE))
        elif ctype in ("highlight", "mark"):
            _render_note_inlines(para_elem, child.get("children", []), _fmt_set(_f, _DC_HIGHLIGHT))
        else:
            nested = child.get("children", [])
            if nested:
                _render_note_inlines(para_elem, nested, _f)


def _fmt_set(fmt: dict[str, bool], key: str) -> dict[str, bool]:
    """Return a shallow copy of *fmt* with *key* = True."""
    return {**fmt, key: True}


def _nr_text(
    para: etree._Element,
    node: dict[str, Any],
    fmt: dict[str, bool] | None = None,
) -> None:
    """Create a ``w:r`` run with the text from *node* and properties from *fmt*."""
    text = node.get("raw", "") or node.get("text", "")
    if not text:
        return
    r = etree.SubElement(para, f"{{{WML_NS}}}r")
    _apply_fmt(r, fmt)
    t = etree.SubElement(r, f"{{{WML_NS}}}t")
    if text.startswith(" ") or text.endswith(" "):
        t.set(f"{{{XML_NS}}}space", "preserve")
    t.text = text


def _apply_fmt(r: etree._Element, fmt: dict[str, bool] | None) -> None:
    """Create ``w:rPr`` on *r* with the formatting flags in *fmt* (if any)."""
    if not fmt:
        return
    rPr = etree.SubElement(r, f"{{{WML_NS}}}rPr")
    if fmt.get(_DC_BOLD):
        etree.SubElement(rPr, f"{{{WML_NS}}}b")
    if fmt.get(_DC_ITALIC):
        etree.SubElement(rPr, f"{{{WML_NS}}}i")
    if fmt.get(_DC_STRIKE):
        etree.SubElement(rPr, f"{{{WML_NS}}}strike")
    if fmt.get(_DC_SUPERSCRIPT):
        va = etree.SubElement(rPr, f"{{{WML_NS}}}vertAlign")
        va.set(f"{{{WML_NS}}}val", "superscript")
    if fmt.get(_DC_SUBSCRIPT):
        va = etree.SubElement(rPr, f"{{{WML_NS}}}vertAlign")
        va.set(f"{{{WML_NS}}}val", "subscript")
    if fmt.get(_DC_UNDERLINE):
        u = etree.SubElement(rPr, f"{{{WML_NS}}}u")
        u.set(f"{{{WML_NS}}}val", "single")
    if fmt.get(_DC_HIGHLIGHT):
        hl = etree.SubElement(rPr, f"{{{WML_NS}}}highlight")
        hl.set(f"{{{WML_NS}}}val", "yellow")


def _nr_codespan(para: etree._Element, node: dict[str, Any]) -> None:
    text = node.get("raw", "") or node.get("text", "")
    if not text:
        return
    r = etree.SubElement(para, f"{{{WML_NS}}}r")
    rPr = etree.SubElement(r, f"{{{WML_NS}}}rPr")
    rFonts = etree.SubElement(rPr, f"{{{WML_NS}}}rFonts")
    rFonts.set(f"{{{WML_NS}}}ascii", "Consolas")
    rFonts.set(f"{{{WML_NS}}}hAnsi", "Consolas")
    shd = etree.SubElement(rPr, f"{{{WML_NS}}}shd")
    shd.set(f"{{{WML_NS}}}val", "clear")
    shd.set(f"{{{WML_NS}}}fill", "D9D9D9")
    t = etree.SubElement(r, f"{{{WML_NS}}}t")
    if text.startswith(" ") or text.endswith(" "):
        t.set(f"{{{XML_NS}}}space", "preserve")
    t.text = text


def _nr_softbreak(para: etree._Element) -> None:
    r = etree.SubElement(para, f"{{{WML_NS}}}r")
    etree.SubElement(r, f"{{{WML_NS}}}br")


def _nr_linebreak(para: etree._Element) -> None:
    r = etree.SubElement(para, f"{{{WML_NS}}}r")
    etree.SubElement(r, f"{{{WML_NS}}}br")


def _create_footnote_element(
    footnote_id: int,
    para_children: list[list[dict[str, Any]]],
    text_style_id: str,
    ref_style_id: str,
) -> etree._Element:
    """Create a ``w:footnote`` OOXML element with formatted paragraphs.

    *para_children* is a list of inline-child lists, one per paragraph.
    Each inline child is a mistune AST dict (text, strong, emphasis,
    codespan, softbreak, …).
    """
    footnote = etree.Element(f"{{{WML_NS}}}footnote")
    footnote.set(f"{{{WML_NS}}}id", str(footnote_id))

    for idx, children in enumerate(para_children):
        p = etree.SubElement(footnote, f"{{{WML_NS}}}p")

        pPr = etree.SubElement(p, f"{{{WML_NS}}}pPr")
        pStyle = etree.SubElement(pPr, f"{{{WML_NS}}}pStyle")
        pStyle.set(f"{{{WML_NS}}}val", text_style_id)

        if idx == 0:
            # Footnote reference mark run
            r1 = etree.SubElement(p, f"{{{WML_NS}}}r")
            rPr1 = etree.SubElement(r1, f"{{{WML_NS}}}rPr")
            rStyle = etree.SubElement(rPr1, f"{{{WML_NS}}}rStyle")
            rStyle.set(f"{{{WML_NS}}}val", ref_style_id)
            etree.SubElement(r1, f"{{{WML_NS}}}footnoteRef")

            # Space run
            r2 = etree.SubElement(p, f"{{{WML_NS}}}r")
            t2 = etree.SubElement(r2, f"{{{WML_NS}}}t")
            t2.set(f"{{{XML_NS}}}space", "preserve")
            t2.text = " "

        _render_note_inlines(p, children)

    return footnote


def _create_endnote_element(
    endnote_id: int,
    para_children: list[list[dict[str, Any]]],
    text_style_id: str,
    ref_style_id: str,
) -> etree._Element:
    """Create a ``w:endnote`` OOXML element with formatted paragraphs."""
    endnote = etree.Element(f"{{{WML_NS}}}endnote")
    endnote.set(f"{{{WML_NS}}}id", str(endnote_id))

    for idx, children in enumerate(para_children):
        p = etree.SubElement(endnote, f"{{{WML_NS}}}p")

        pPr = etree.SubElement(p, f"{{{WML_NS}}}pPr")
        pStyle = etree.SubElement(pPr, f"{{{WML_NS}}}pStyle")
        pStyle.set(f"{{{WML_NS}}}val", text_style_id)

        if idx == 0:
            # Endnote reference mark
            r1 = etree.SubElement(p, f"{{{WML_NS}}}r")
            rPr1 = etree.SubElement(r1, f"{{{WML_NS}}}rPr")
            rStyle = etree.SubElement(rPr1, f"{{{WML_NS}}}rStyle")
            rStyle.set(f"{{{WML_NS}}}val", ref_style_id)
            etree.SubElement(r1, f"{{{WML_NS}}}endnoteRef")

            r2 = etree.SubElement(p, f"{{{WML_NS}}}r")
            t2 = etree.SubElement(r2, f"{{{WML_NS}}}t")
            t2.set(f"{{{XML_NS}}}space", "preserve")
            t2.text = " "

        _render_note_inlines(p, children)

    return endnote


def _create_footnote_ref_run(footnote_id: int, ref_style_id: str) -> etree._Element:
    """Create an inline ``w:r`` for a footnote reference in body text."""
    r = etree.Element(f"{{{WML_NS}}}r")
    rPr = etree.SubElement(r, f"{{{WML_NS}}}rPr")
    rStyle = etree.SubElement(rPr, f"{{{WML_NS}}}rStyle")
    rStyle.set(f"{{{WML_NS}}}val", ref_style_id)
    footnoteRef = etree.SubElement(r, f"{{{WML_NS}}}footnoteReference")
    footnoteRef.set(f"{{{WML_NS}}}id", str(footnote_id))
    return r


def _create_endnote_ref_run(endnote_id: int, ref_style_id: str) -> etree._Element:
    """Create an inline ``w:r`` for an endnote reference in body text."""
    r = etree.Element(f"{{{WML_NS}}}r")
    rPr = etree.SubElement(r, f"{{{WML_NS}}}rPr")
    rStyle = etree.SubElement(rPr, f"{{{WML_NS}}}rStyle")
    rStyle.set(f"{{{WML_NS}}}val", ref_style_id)
    endnoteRef = etree.SubElement(r, f"{{{WML_NS}}}endnoteReference")
    endnoteRef.set(f"{{{WML_NS}}}id", str(endnote_id))
    return r


# ── Pipeline entry point ────────────────────────────────────────────────


def process_md_body_with_notes(
    md_body: str,
) -> tuple[list[dict[str, Any]], NoteContext]:
    """Parse markdown and extract footnote/endnote definitions.

    Parses the full markdown with mistune (footnotes plugin enabled),
    extracts definitions from the ``footnotes`` AST node, and returns
    the cleaned AST together with a populated :class:`NoteContext`.

    This is the main entry point for wiring notes into the conversion
    pipeline.
    """
    from docwen_plugin_markdown.mistune_extensions import parse_markdown_text

    ast_nodes = parse_markdown_text(normalize_note_syntax(md_body))
    return extract_notes_from_ast(ast_nodes)


# ── DOCX part writeback ──────────────────────────────────────────────────

# OOXML package namespaces and constants
_RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
_OFFICE_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

_FOOTNOTES_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
_ENDNOTES_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"
_FOOTNOTES_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"
_ENDNOTES_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"

# Base XML templates — Word requires separator (id=-1) and
# continuationSeparator (id=0) in every footnotes/endnotes part.
_BASE_FOOTNOTES_XML = (
    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    b'<w:footnotes xmlns:w="http://schemas.openxmlformats.org/'
    b'wordprocessingml/2006/main"'
    b' xmlns:r="http://schemas.openxmlformats.org/officeDocument/'
    b'2006/relationships">'
    b'<w:footnote w:type="separator" w:id="-1">'
    b'<w:p><w:pPr><w:spacing w:after="0" w:line="240" '
    b'w:lineRule="auto"/></w:pPr>'
    b"<w:r><w:separator/></w:r></w:p>"
    b"</w:footnote>"
    b'<w:footnote w:type="continuationSeparator" w:id="0">'
    b'<w:p><w:pPr><w:spacing w:after="0" w:line="240" '
    b'w:lineRule="auto"/></w:pPr>'
    b"<w:r><w:continuationSeparator/></w:r></w:p>"
    b"</w:footnote>"
    b"</w:footnotes>"
)

_BASE_ENDNOTES_XML = (
    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    b'<w:endnotes xmlns:w="http://schemas.openxmlformats.org/'
    b'wordprocessingml/2006/main"'
    b' xmlns:r="http://schemas.openxmlformats.org/officeDocument/'
    b'2006/relationships">'
    b'<w:endnote w:type="separator" w:id="-1">'
    b'<w:p><w:pPr><w:spacing w:after="0" w:line="240" '
    b'w:lineRule="auto"/></w:pPr>'
    b"<w:r><w:separator/></w:r></w:p>"
    b"</w:endnote>"
    b'<w:endnote w:type="continuationSeparator" w:id="0">'
    b'<w:p><w:pPr><w:spacing w:after="0" w:line="240" '
    b'w:lineRule="auto"/></w:pPr>'
    b"<w:r><w:continuationSeparator/></w:r></w:p>"
    b"</w:endnote>"
    b"</w:endnotes>"
)

_NOTE_SPECS = (
    (
        "footnote",
        "word/footnotes.xml",
        "word/_rels/footnotes.xml.rels",
        _FOOTNOTES_REL_TYPE,
        _FOOTNOTES_CT,
    ),
    (
        "endnote",
        "word/endnotes.xml",
        "word/_rels/endnotes.xml.rels",
        _ENDNOTES_REL_TYPE,
        _ENDNOTES_CT,
    ),
)
_NOTE_ID = re.compile(r"[+-]?\d+")
_NOTE_TYPES = frozenset({"normal", "separator", "continuationSeparator", "continuationNotice"})
_RELATIONSHIP_ATTRIBUTE_NAMES = frozenset({"id", "embed", "link"})


def prepare_note_context_for_document(document, note_ctx: NoteContext) -> None:
    """Audit request-template note graphs and seed collision-free ID allocation."""

    if not note_ctx.has_notes:
        return
    buffer = BytesIO()
    try:
        document.save(buffer)
        footnote_ids, endnote_ids = _audit_note_package(buffer.getvalue())
    except NoteWritebackError:
        raise
    except (BadZipFile, etree.XMLSyntaxError, KeyError, OSError, ValueError) as exc:
        raise NoteWritebackError(
            "MD2DOCX-NOTE-PART-INVALID",
            f"Existing DOCX note graph is invalid: {exc}",
            error_type="invalid_input",
        ) from exc
    note_ctx.reserve_existing_ids(footnote_ids=footnote_ids, endnote_ids=endnote_ids)


def _audit_note_package(
    blob: bytes,
    *,
    pending_footnote_ids: set[int] | None = None,
    pending_endnote_ids: set[int] | None = None,
) -> tuple[set[int], set[int]]:
    with ZipFile(BytesIO(blob), "r") as archive:
        names = [item.filename for item in archive.infolist()]
        if len(names) != len(set(names)):
            _note_invalid("DOCX package contains duplicate ZIP members.")
        name_set = set(names)
        rels_root = etree.fromstring(archive.read("word/_rels/document.xml.rels"))
        ct_root = etree.fromstring(archive.read("[Content_Types].xml"))
        result: list[set[int]] = []
        for note_name, part_name, rels_name, rel_type, content_type in _NOTE_SPECS:
            result.append(
                _audit_note_part(
                    archive,
                    name_set,
                    rels_root,
                    ct_root,
                    note_name=note_name,
                    part_name=part_name,
                    rels_name=rels_name,
                    rel_type=rel_type,
                    content_type=content_type,
                )
            )
        referenced_footnotes, referenced_endnotes = _audit_document_note_references(
            archive,
            footnote_ids=result[0] | (pending_footnote_ids or set()),
            endnote_ids=result[1] | (pending_endnote_ids or set()),
        )
        if (pending_footnote_ids or set()) - referenced_footnotes:
            _note_invalid("A pending footnote body has no matching main-document reference.")
        if (pending_endnote_ids or set()) - referenced_endnotes:
            _note_invalid("A pending endnote body has no matching main-document reference.")
        return result[0], result[1]


def _audit_note_part(
    archive: ZipFile,
    names: set[str],
    document_rels: etree._Element,
    content_types: etree._Element,
    *,
    note_name: str,
    part_name: str,
    rels_name: str,
    rel_type: str,
    content_type: str,
) -> set[int]:
    part_present = part_name in names
    document_relationships = [
        rel
        for rel in document_rels.findall(f"{{{_RELS_NS}}}Relationship")
        if rel.get("Type") == rel_type or _resolve_part_target("word/document.xml", rel.get("Target", "")) == part_name
    ]
    overrides = [
        item for item in content_types.findall(f"{{{_CT_NS}}}Override") if item.get("PartName") == f"/{part_name}"
    ]
    if not part_present:
        if document_relationships or overrides or rels_name in names:
            _note_invalid(f"{part_name} has dangling relationship, content-type, or relationship-part records.")
        return set()
    if len(document_relationships) != 1:
        _note_invalid(f"{part_name} must have exactly one owning document relationship.")
    relationship = document_relationships[0]
    if (
        relationship.get("Type") != rel_type
        or relationship.get("TargetMode") not in {None, "Internal"}
        or _resolve_part_target("word/document.xml", relationship.get("Target", "")) != part_name
    ):
        _note_invalid(f"{part_name} has an invalid owning document relationship.")
    if len(overrides) != 1 or overrides[0].get("ContentType") != content_type:
        _note_invalid(f"{part_name} has an invalid content-type override.")

    root = etree.fromstring(archive.read(part_name))
    expected_root = f"{{{WML_NS}}}{note_name}s"
    expected_child = f"{{{WML_NS}}}{note_name}"
    if root.tag != expected_root:
        _note_invalid(f"{part_name} has the wrong root element.")
    all_ids: set[int] = set()
    positive_ids: set[int] = set()
    by_id: dict[int, str] = {}
    for element in root:
        if element.tag != expected_child:
            _note_invalid(f"{part_name} contains an unexpected direct child element.")
        raw_id = element.get(f"{{{WML_NS}}}id")
        if not isinstance(raw_id, str) or _NOTE_ID.fullmatch(raw_id) is None:
            _note_invalid(f"{part_name} contains a missing or malformed note ID.")
        note_id = int(raw_id)
        if not -(2**31) <= note_id < 2**31 or note_id in all_ids:
            _note_invalid(f"{part_name} contains an out-of-range or duplicate note ID.")
        note_type = element.get(f"{{{WML_NS}}}type", "normal")
        if note_type not in _NOTE_TYPES:
            _note_invalid(f"{part_name} contains an unknown note type.")
        if (note_id > 0) != (note_type == "normal"):
            _note_invalid(f"{part_name} note ID/type ownership is invalid.")
        all_ids.add(note_id)
        by_id[note_id] = note_type
        if note_id > 0:
            positive_ids.add(note_id)
    if by_id.get(-1) != "separator" or by_id.get(0) != "continuationSeparator":
        _note_invalid(f"{part_name} does not contain the required reserved separator IDs.")
    _audit_note_relationships(archive, names, root, part_name=part_name, rels_name=rels_name)
    return positive_ids


def _audit_note_relationships(
    archive: ZipFile,
    names: set[str],
    note_root: etree._Element,
    *,
    part_name: str,
    rels_name: str,
) -> None:
    referenced_ids: set[str] = set()
    for element in note_root.iter():
        for raw_name, value in element.attrib.items():
            qname = etree.QName(raw_name)
            if qname.namespace == _OFFICE_REL_NS and qname.localname in _RELATIONSHIP_ATTRIBUTE_NAMES:
                if not value:
                    _note_invalid(f"{part_name} contains a blank relationship reference.")
                referenced_ids.add(str(value))
    if rels_name not in names:
        if referenced_ids:
            _note_invalid(f"{part_name} references relationships but has no relationship part.")
        return
    root = etree.fromstring(archive.read(rels_name))
    if root.tag != f"{{{_RELS_NS}}}Relationships":
        _note_invalid(f"{rels_name} has the wrong root element.")
    relationships: dict[str, etree._Element] = {}
    for relationship in root:
        if relationship.tag != f"{{{_RELS_NS}}}Relationship":
            _note_invalid(f"{rels_name} contains an unexpected direct child element.")
        rel_id = relationship.get("Id", "")
        target = relationship.get("Target", "")
        if not rel_id or rel_id in relationships or not relationship.get("Type") or not target:
            _note_invalid(f"{rels_name} contains a malformed or duplicate relationship.")
        target_mode = relationship.get("TargetMode")
        if target_mode not in {None, "Internal", "External"}:
            _note_invalid(f"{rels_name} contains an invalid TargetMode.")
        if target_mode != "External" and _resolve_part_target(part_name, target) not in names:
            _note_invalid(f"{rels_name} targets a missing package part.")
        relationships[rel_id] = relationship
    missing = referenced_ids - relationships.keys()
    if missing:
        _note_invalid(f"{part_name} references missing relationship IDs: {sorted(missing)!r}.")


def _audit_document_note_references(
    archive: ZipFile,
    *,
    footnote_ids: set[int],
    endnote_ids: set[int],
) -> tuple[set[int], set[int]]:
    document = etree.fromstring(archive.read("word/document.xml"))
    observed: list[set[int]] = []
    for element_name, valid_ids in (("footnoteReference", footnote_ids), ("endnoteReference", endnote_ids)):
        observed_ids: set[int] = set()
        for element in document.iter(f"{{{WML_NS}}}{element_name}"):
            raw_id = element.get(f"{{{WML_NS}}}id")
            if not isinstance(raw_id, str) or _NOTE_ID.fullmatch(raw_id) is None or int(raw_id) not in valid_ids:
                _note_invalid(f"word/document.xml contains a dangling or malformed {element_name}.")
            observed_ids.add(int(raw_id))
        observed.append(observed_ids)
    return observed[0], observed[1]


def _resolve_part_target(owner_part: str, target: str) -> str:
    clean = target.replace("\\", "/")
    if not clean or ":" in clean.split("/", 1)[0]:
        return ""
    if clean.startswith("/"):
        resolved = posixpath.normpath(clean.lstrip("/"))
    else:
        resolved = posixpath.normpath(posixpath.join(posixpath.dirname(owner_part), clean))
    if resolved in {"", ".", ".."} or resolved.startswith("../"):
        return ""
    return resolved


def _note_invalid(message: str) -> NoReturn:
    raise NoteWritebackError(
        "MD2DOCX-NOTE-PART-INVALID",
        message,
        error_type="invalid_input",
    )


def write_notes_to_docx(docx_path: str, note_ctx: NoteContext) -> None:
    """Write footnote/endnote body elements into a saved DOCX file.

    Opens the DOCX as a ZIP, reads or creates ``word/footnotes.xml``
    and ``word/endnotes.xml``, appends the elements stored in
    *note_ctx*, updates document relationships and content types,
    then writes back to *docx_path*.

    This must be called **after** ``Document.save()`` so the ZIP
    contains all body content.
    """
    if not note_ctx.footnote_elements and not note_ctx.endnote_elements:
        return

    original = Path(docx_path)
    try:
        import tempfile
        import zipfile

        new_footnote_ids = _new_note_ids(
            note_ctx.footnote_elements,
            note_name="footnote",
        )
        new_endnote_ids = _new_note_ids(
            note_ctx.endnote_elements,
            note_name="endnote",
        )
        existing_footnotes, existing_endnotes = _audit_note_package(
            original.read_bytes(),
            pending_footnote_ids=new_footnote_ids,
            pending_endnote_ids=new_endnote_ids,
        )
        if new_footnote_ids & existing_footnotes or new_endnote_ids & existing_endnotes:
            _note_invalid("A new note ID collides with the existing note domain.")

        with tempfile.TemporaryDirectory(prefix=".dw-notes-", dir=original.parent) as tmpdir:
            tmp_path = Path(tmpdir) / "notes_writeback.docx"

            with (
                zipfile.ZipFile(str(original), "r") as zf_in,
                zipfile.ZipFile(str(tmp_path), "w", zipfile.ZIP_DEFLATED) as zf_out,
            ):
                names = set(zf_in.namelist())
                rels_root = etree.fromstring(zf_in.read("word/_rels/document.xml.rels"))
                ct_root = etree.fromstring(zf_in.read("[Content_Types].xml"))

                fn_bytes: bytes | None = None
                if note_ctx.footnote_elements:
                    if "word/footnotes.xml" in names:
                        fn_bytes = _append_elements_to_xml(
                            zf_in.read("word/footnotes.xml"),
                            note_ctx.footnote_elements,
                        )
                    else:
                        fn_bytes = _build_notes_xml(_BASE_FOOTNOTES_XML, note_ctx.footnote_elements)
                        _ensure_relationship(rels_root, "footnotes.xml", _FOOTNOTES_REL_TYPE)
                        _ensure_content_type(ct_root, "/word/footnotes.xml", _FOOTNOTES_CT)

                en_bytes: bytes | None = None
                if note_ctx.endnote_elements:
                    if "word/endnotes.xml" in names:
                        en_bytes = _append_elements_to_xml(
                            zf_in.read("word/endnotes.xml"),
                            note_ctx.endnote_elements,
                        )
                    else:
                        en_bytes = _build_notes_xml(_BASE_ENDNOTES_XML, note_ctx.endnote_elements)
                        _ensure_relationship(rels_root, "endnotes.xml", _ENDNOTES_REL_TYPE)
                        _ensure_content_type(ct_root, "/word/endnotes.xml", _ENDNOTES_CT)

                modified_rels = etree.tostring(
                    rels_root,
                    xml_declaration=True,
                    encoding="UTF-8",
                    standalone=True,
                )
                modified_ct = etree.tostring(
                    ct_root,
                    xml_declaration=True,
                    encoding="UTF-8",
                    standalone=True,
                )

                for item in zf_in.infolist():
                    if item.filename == "word/footnotes.xml" and fn_bytes is not None:
                        zf_out.writestr(item, fn_bytes)
                    elif item.filename == "word/endnotes.xml" and en_bytes is not None:
                        zf_out.writestr(item, en_bytes)
                    elif item.filename == "word/_rels/document.xml.rels":
                        zf_out.writestr(item, modified_rels)
                    elif item.filename == "[Content_Types].xml":
                        zf_out.writestr(item, modified_ct)
                    else:
                        zf_out.writestr(item, zf_in.read(item.filename))

                if fn_bytes is not None and "word/footnotes.xml" not in names:
                    zf_out.writestr("word/footnotes.xml", fn_bytes)
                if en_bytes is not None and "word/endnotes.xml" not in names:
                    zf_out.writestr("word/endnotes.xml", en_bytes)

            _audit_note_package(tmp_path.read_bytes())
            tmp_path.replace(original)
    except NoteWritebackError:
        raise
    except (BadZipFile, etree.XMLSyntaxError, KeyError, OSError, ValueError) as exc:
        raise NoteWritebackError(
            "MD2DOCX-NOTE-WRITEBACK-ERROR",
            f"Atomic DOCX note writeback failed: {exc}",
            error_type="conversion_failed",
        ) from exc


def _new_note_ids(
    elements: list[etree._Element],
    *,
    note_name: str,
) -> set[int]:
    expected_tag = f"{{{WML_NS}}}{note_name}"
    new_ids: set[int] = set()
    for element in elements:
        raw_id = element.get(f"{{{WML_NS}}}id")
        if element.tag != expected_tag or not isinstance(raw_id, str) or _NOTE_ID.fullmatch(raw_id) is None:
            _note_invalid(f"New {note_name} element has an invalid shape or ID.")
        note_id = int(raw_id)
        if not 0 < note_id < 2**31 or note_id in new_ids:
            _note_invalid(f"New {note_name} ID collides with the existing note domain.")
        new_ids.add(note_id)
    return new_ids


def _append_elements_to_xml(
    xml_bytes: bytes,
    elements: list[etree._Element],
) -> bytes:
    """Parse *xml_bytes*, append each element in *elements* to the root,
    and return the serialised result."""
    root = etree.fromstring(xml_bytes)
    for elem in elements:
        root.append(elem)
    return etree.tostring(
        root,
        xml_declaration=True,
        encoding="UTF-8",
        standalone=True,
    )


def _build_notes_xml(
    template_bytes: bytes,
    elements: list[etree._Element],
) -> bytes:
    """Parse a *template_bytes* base XML, append *elements*, and serialise."""
    root = etree.fromstring(template_bytes)
    for elem in elements:
        root.append(elem)
    return etree.tostring(
        root,
        xml_declaration=True,
        encoding="UTF-8",
        standalone=True,
    )


def _ensure_relationship(
    rels_root: etree._Element,
    target: str,
    rel_type: str,
) -> None:
    """Add a ``<Relationship>`` to *rels_root* if one targeting *target*
    does not already exist."""
    used_ids: set[str] = set()
    for rel in rels_root.findall(f"{{{_RELS_NS}}}Relationship"):
        rel_id = rel.get("Id", "")
        if not rel_id or rel_id in used_ids:
            _note_invalid("word/_rels/document.xml.rels contains a malformed or duplicate relationship ID.")
        used_ids.add(rel_id)
        same_target = _resolve_part_target("word/document.xml", rel.get("Target", "")) == f"word/{target}"
        same_type = rel.get("Type") == rel_type
        if same_target or same_type:
            if same_target and same_type and rel.get("TargetMode") not in {"External"}:
                return
            _note_invalid("A conflicting document relationship prevents note-part ownership.")

    candidate = 1
    while f"rId{candidate}" in used_ids:
        candidate += 1
    new_id = f"rId{candidate}"
    elem = etree.SubElement(rels_root, f"{{{_RELS_NS}}}Relationship")
    elem.set("Id", new_id)
    elem.set("Type", rel_type)
    elem.set("Target", target)


def _ensure_content_type(
    ct_root: etree._Element,
    part_name: str,
    content_type: str,
) -> None:
    """Add a ``<Override>`` to *ct_root* for *part_name* if one does not
    already exist."""
    for override in ct_root.findall(f"{{{_CT_NS}}}Override"):
        if override.get("PartName") == part_name:
            if override.get("ContentType") == content_type:
                return
            _note_invalid(f"Content-type override for {part_name} is occupied by another type.")

    override_elem = etree.SubElement(ct_root, f"{{{_CT_NS}}}Override")
    override_elem.set("PartName", part_name)
    override_elem.set("ContentType", content_type)

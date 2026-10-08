"""Represent repeated Markdown notes as Word cross-references."""

from collections import defaultdict
from copy import deepcopy
from dataclasses import dataclass
from io import BytesIO

from docx import Document
from lxml import etree

from docwen_core.docx_bookmarks import build_docx_bookmark_inventory
from docwen_core.text.numbering import (
    number_to_chinese,
    number_to_circled,
    number_to_roman_lower,
    number_to_roman_upper,
)

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


@dataclass(frozen=True)
class NoteReferenceProjection:
    document_xml: bytes
    deferred_count: int


def _word_letter_mark(number: int, *, uppercase: bool) -> str:
    """OOXML letter numbering repeats one letter: Z, AA, BB, ..., ZZ, AAA."""
    if number <= 0:
        return ""
    repeats, letter = divmod(number - 1, 26)
    return chr((ord("A") if uppercase else ord("a")) + letter) * (repeats + 1)


def _cached_note_marks(document, kind: str) -> dict[int, str | None]:
    """Compute first-reference marks from document and section note settings.

    Page-dependent restarts still require the editor's layout/field update;
    never substitute a repeated occurrence's ordinal for its original mark.
    """
    formats = {
        "decimal": str,
        "lowerRoman": number_to_roman_lower,
        "upperRoman": number_to_roman_upper,
        "lowerLetter": lambda number: _word_letter_mark(number, uppercase=False),
        "upperLetter": lambda number: _word_letter_mark(number, uppercase=True),
        "chineseCounting": number_to_chinese,
        "chineseCountingThousand": number_to_chinese,
        "decimalEnclosedCircle": number_to_circled,
    }
    defaults = {
        "numFmt": "decimal" if kind == "footnote" else "lowerRoman",
        "numStart": "1",
        "numRestart": "continuous",
    }

    def properties(parent, inherited):
        result = dict(inherited)
        node = parent.find(_W + kind + "Pr")
        if node is not None:
            for key in defaults:
                value = node.find(_W + key)
                if value is not None:
                    result[key] = value.get(_W + "val", result[key])
        return result

    defaults = properties(document.settings.element, defaults)
    marks: dict[int, str | None] = {}
    pending = []
    counter = None
    for block in document.element.body:
        pending.extend(block.iter(_W + kind + "Reference"))
        section = block if block.tag == _W + "sectPr" else block.find(_W + "pPr/" + _W + "sectPr")
        if section is None:
            continue
        settings = properties(section, defaults)
        if counter is None or settings["numRestart"] == "eachSect":
            counter = int(settings["numStart"]) - 1
        for reference in pending:
            # A template's custom symbol replaces the automatic mark and does
            # not advance its numbering sequence (OOXML customMarkFollows).
            if reference.get(_W + "customMarkFollows") in {"true", "1", "on"}:
                continue
            note_id = int(reference.get(_W + "id"))
            if note_id in marks:
                continue
            counter += 1
            formatter = formats.get(settings["numFmt"])
            maximum = {
                "lowerRoman": 3999,
                "upperRoman": 3999,
                "chineseCounting": 99,
                "chineseCountingThousand": 99,
                "decimalEnclosedCircle": 50,
            }.get(settings["numFmt"])
            exact_range = counter > 0 and (maximum is None or counter <= maximum)
            marks[note_id] = (
                formatter(counter)
                if settings["numRestart"] != "eachPage" and formatter is not None and exact_range
                else None
            )
        pending.clear()
    return marks


def project_repeated_note_references(
    package: bytes, footnotes: set[int], endnotes: set[int]
) -> NoteReferenceProjection:
    """Only replace repeats of notes owned by this conversion.

    One native reference owns the note. Subsequent references are ordinary
    NOTEREF fields, with a bookmark around the first reference. Inventory the
    complete package after rendering so template and caption bookmarks count.
    """
    document = Document(BytesIO(package))
    root = document.element
    inventory = build_docx_bookmark_inventory(document)
    used_ids = set(inventory.used_id_keys)
    used_names = set(inventory.used_name_keys)
    next_id = 0
    deferred_count = 0
    for kind, owned_ids in (("footnote", footnotes), ("endnote", endnotes)):
        cached_marks = _cached_note_marks(document, kind)
        grouped: dict[int, list] = defaultdict(list)
        for reference in root.iter(_W + kind + "Reference"):
            value = reference.get(_W + "id")
            if value is None:
                raise ValueError("Note reference has no ID")
            grouped[int(value)].append(reference)
        for note_id, references in grouped.items():
            if note_id not in owned_ids or len(references) < 2:
                continue
            while ("numeric", next_id) in used_ids:
                next_id += 1
            if next_id > 2_147_483_647:
                raise ValueError("No portable bookmark ID remains for repeated note")
            bookmark_id = str(next_id)
            used_ids.add(("numeric", next_id))
            stem = f"_Note_{kind}_{note_id}"
            name = stem
            suffix = 0
            while name.casefold() in used_names:
                suffix += 1
                name = f"{stem}_{suffix}"
            if len(name) > 40:
                raise ValueError("Repeated note bookmark name exceeds Word limit")
            used_names.add(name.casefold())
            first_run = references[0].getparent()
            start = etree.Element(_W + "bookmarkStart", {_W + "id": bookmark_id, _W + "name": name})
            end = etree.Element(_W + "bookmarkEnd", {_W + "id": bookmark_id})
            first_run.addprevious(start)
            first_run.addnext(end)
            for reference in references[1:]:
                mark = cached_marks[note_id]
                if mark is None:
                    deferred_count += 1
                run = reference.getparent()
                if run.tag != _W + "r" or any(child.tag not in {_W + "rPr", reference.tag} for child in run):
                    raise ValueError("Repeated note reference is not an isolated generated run")
                properties = run.find(_W + "rPr")
                cached_properties = deepcopy(properties) if properties is not None else etree.Element(_W + "rPr")
                # Native note references supply their own raised mark; plain
                # cached text does not, even when it uses the same character
                # style. Keep the cache consistent before any field update.
                vertical = cached_properties.find(_W + "vertAlign")
                if vertical is None:
                    vertical = etree.SubElement(cached_properties, _W + "vertAlign")
                vertical.set(_W + "val", "superscript")
                # Word drops the cached run formatting while converting a
                # simple field on save. Emit a formatted complex field so its
                # reference style survives edit/save/reopen.
                for tag, value in (
                    ("fldChar", "begin"),
                    ("instrText", f" NOTEREF {name} \\h \\f "),
                    ("fldChar", "separate"),
                    ("t", mark if mark is not None else "?"),
                    ("fldChar", "end"),
                ):
                    field_run = etree.Element(_W + "r")
                    field_run.append(deepcopy(cached_properties))
                    payload = etree.SubElement(field_run, _W + tag)
                    if tag == "fldChar":
                        payload.set(_W + "fldCharType", value)
                        if value == "begin" and mark is None:
                            payload.set(_W + "dirty", "true")
                    else:
                        payload.text = value
                        if tag == "instrText":
                            payload.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                    run.addprevious(field_run)
                run.getparent().remove(run)
    return NoteReferenceProjection(
        etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True), deferred_count
    )

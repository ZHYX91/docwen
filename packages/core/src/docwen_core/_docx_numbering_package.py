"""Low-level package and numbering-part primitives for resolved numbering."""

from __future__ import annotations

import contextlib
import os
import tempfile
from collections.abc import Mapping
from pathlib import Path
from typing import Any
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

import lxml.etree as etree

from docwen_core.docx_numbering_parts import NumberingPartError, open_numbering_part
from docwen_core.docx_numbering_structure import append_numbering_definitions
from docwen_core.file_io import replace_authorized_file
from docwen_core.models.resolved_numbering import (
    HeadingCounterSegment,
    HeadingLiteralSegment,
)

WML_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
NUMBERING_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
NUMBERING_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

_NUMBER_FORMATS = {
    "chinese_lower": "chineseCounting",
    "chinese_upper": "chineseCountingThousand",
    "arabic_half": "decimal",
    "arabic_full": "decimalFullWidth",
    "arabic_circled": "decimalEnclosedCircleChinese",
    "letter_upper": "upperLetter",
    "letter_lower": "lowerLetter",
    "roman_upper": "upperRoman",
    "roman_lower": "lowerRoman",
}


class ResolvedNumberingOoxmlError(ValueError):
    """The validated snapshot cannot be represented by the closed OOXML form."""


def write_heading_numbering_projection(path: str | Path, projection: Any) -> None:
    """Atomically merge the request definitions into a saved DOCX package."""

    if not projection.abstract_nums and not projection.nums:
        return
    package_path = Path(path)
    with ZipFile(package_path) as package:
        infos = package.infolist()
        if len({item.filename for item in infos}) != len(infos):
            raise ResolvedNumberingOoxmlError("DOCX contains duplicate ZIP members")
        data = {item.filename: package.read(item.filename) for item in infos}
    try:
        part = open_numbering_part(data, create=True)
    except NumberingPartError as exc:
        raise ResolvedNumberingOoxmlError(str(exc)) from exc
    assert part is not None
    root = part.element
    existing_abstract, existing_num = _numbering_ids(root)
    projected_abstract, projected_num = _numbering_ids_from_elements(projection)
    if existing_abstract & projected_abstract or existing_num & projected_num:
        raise ResolvedNumberingOoxmlError("resolved Heading numbering IDs collide after save")
    append_numbering_definitions(root, projection.abstract_nums, projection.nums)
    data[part.name] = etree.tostring(
        root,
        encoding="UTF-8",
        xml_declaration=True,
        standalone=True,
    )

    existing_names = {item.filename for item in infos}
    _rewrite_package(package_path, infos, data, new_names=tuple(data.keys() - existing_names))


def _heading_abstract_num(
    abstract_id: int,
    definition: Any,
    *,
    heading_style_ids: Mapping[int, str],
) -> Any:
    element = etree.Element(f"{{{WML_NS}}}abstractNum")
    element.set(f"{{{WML_NS}}}abstractNumId", str(abstract_id))
    multi = etree.SubElement(element, f"{{{WML_NS}}}multiLevelType")
    multi.set(f"{{{WML_NS}}}val", "multilevel")
    for level in definition.levels:
        style_id = heading_style_ids.get(level.level)
        if not style_id:
            raise ResolvedNumberingOoxmlError(f"Heading {level.level} lacks a resolved style")
        node = etree.SubElement(element, f"{{{WML_NS}}}lvl")
        node.set(f"{{{WML_NS}}}ilvl", str(level.level - 1))
        for name, value in (
            ("start", str(level.start)),
            ("numFmt", _NUMBER_FORMATS.get(level.number_format)),
            ("lvlRestart", "0" if level.restart_after_level is None else str(level.restart_after_level)),
            ("suff", level.suffix),
            ("lvlText", _level_text(level)),
        ):
            if value is None:
                raise ResolvedNumberingOoxmlError("Heading number format is not portable")
            child = etree.SubElement(node, f"{{{WML_NS}}}{name}")
            child.set(f"{{{WML_NS}}}val", value)
    return element


def _heading_num(num_id: int, abstract_id: int, starts: Any) -> Any:
    element = etree.Element(f"{{{WML_NS}}}num")
    element.set(f"{{{WML_NS}}}numId", str(num_id))
    reference = etree.SubElement(element, f"{{{WML_NS}}}abstractNumId")
    reference.set(f"{{{WML_NS}}}val", str(abstract_id))
    for start in starts:
        override = etree.SubElement(element, f"{{{WML_NS}}}lvlOverride")
        override.set(f"{{{WML_NS}}}ilvl", str(start.level - 1))
        value = etree.SubElement(override, f"{{{WML_NS}}}startOverride")
        value.set(f"{{{WML_NS}}}val", str(start.value))
    return element


def _level_text(level: Any) -> str:
    parts: list[str] = []
    for segment in level.display:
        if isinstance(segment, HeadingCounterSegment):
            parts.append(f"%{segment.level}")
        elif isinstance(segment, HeadingLiteralSegment):
            parts.append(segment.literal)
        else:
            raise ResolvedNumberingOoxmlError("Heading display segment is not closed")
    return "".join(parts)


def _append_complex_field(container: Any, *, instruction: str, cached_result: str) -> None:
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    begin_run = OxmlElement("w:r")
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    begin.set(qn("w:dirty"), "true")
    begin_run.append(begin)
    instruction_run = OxmlElement("w:r")
    instruction_text = OxmlElement("w:instrText")
    instruction_text.set(XML_SPACE, "preserve")
    instruction_text.text = instruction
    instruction_run.append(instruction_text)
    separate_run = OxmlElement("w:r")
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    separate_run.append(separate)
    result_run = OxmlElement("w:r")
    result = OxmlElement("w:t")
    if cached_result[:1].isspace() or cached_result[-1:].isspace():
        result.set(XML_SPACE, "preserve")
    result.text = cached_result
    result_run.append(result)
    end_run = OxmlElement("w:r")
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    end_run.append(end)
    container.extend((begin_run, instruction_run, separate_run, result_run, end_run))


def _append_text_run(container: Any, value: str) -> None:
    from docx.oxml import OxmlElement

    if not value:
        return
    run = OxmlElement("w:r")
    text = OxmlElement("w:t")
    if value[:1].isspace() or value[-1:].isspace():
        text.set(XML_SPACE, "preserve")
    text.text = value
    run.append(text)
    container.append(run)


def _prove_complex_field(
    payload: tuple[Any, ...],
    cursor: int,
    *,
    instruction: str,
    cached_result: str,
) -> int:
    if cursor + 5 > len(payload):
        raise ResolvedNumberingOoxmlError("caption complex field is incomplete")
    _prove_field_marker(payload[cursor], "begin", dirty=True)
    _prove_instruction_run(payload[cursor + 1], instruction)
    _prove_field_marker(payload[cursor + 2], "separate", dirty=False)
    _prove_text_run(payload[cursor + 3], cached_result)
    _prove_field_marker(payload[cursor + 4], "end", dirty=False)
    return cursor + 5


def _prove_field_marker(run: Any, field_type: str, *, dirty: bool) -> None:
    element = _only_run_child(run, f"{{{WML_NS}}}fldChar")
    expected_attributes = (f"{{{WML_NS}}}fldCharType", f"{{{WML_NS}}}dirty") if dirty else (f"{{{WML_NS}}}fldCharType",)
    if (
        tuple(element.attrib) != expected_attributes
        or element.get(f"{{{WML_NS}}}fldCharType") != field_type
        or element.get(f"{{{WML_NS}}}dirty") != ("true" if dirty else None)
        or element.get(f"{{{WML_NS}}}fldLock") is not None
    ):
        raise ResolvedNumberingOoxmlError("caption field marker is not canonical")


def _prove_instruction_run(run: Any, instruction: str) -> None:
    element = _only_run_child(run, f"{{{WML_NS}}}instrText")
    if tuple(element.attrib) != (XML_SPACE,) or element.get(XML_SPACE) != "preserve" or element.text != instruction:
        raise ResolvedNumberingOoxmlError("caption field instruction is not exact")


def _prove_text_run(run: Any, value: str) -> None:
    element = _only_run_child(run, f"{{{WML_NS}}}t")
    preserve = value[:1].isspace() or value[-1:].isspace()
    expected_attributes = (XML_SPACE,) if preserve else ()
    if (
        tuple(element.attrib) != expected_attributes
        or element.get(XML_SPACE) != ("preserve" if preserve else None)
        or element.text != value
    ):
        raise ResolvedNumberingOoxmlError("caption text run is not exact")


def _only_run_child(run: Any, expected_tag: str) -> Any:
    children = list(run)
    if (
        run.tag != f"{{{WML_NS}}}r"
        or run.attrib
        or len(children) != 1
        or children[0].tag != expected_tag
        or children[0].tail is not None
        or len(children[0]) != 0
    ):
        raise ResolvedNumberingOoxmlError("caption run topology is not canonical")
    return children[0]


def _prove_caption_bookmark_pair(
    payload: tuple[Any, ...],
    start_index: int,
    end_index: int,
    bookmark_name: str,
    bookmark_inventory: Any,
    paragraph_element: Any,
) -> int:
    from docwen_core.docx_bookmarks import prove_bookmark_name

    if end_index >= len(payload):
        raise ResolvedNumberingOoxmlError("caption bookmark range is incomplete")
    start = payload[start_index]
    end = payload[end_index]
    if (
        start.tag != f"{{{WML_NS}}}bookmarkStart"
        or tuple(start.attrib) != (f"{{{WML_NS}}}id", f"{{{WML_NS}}}name")
        or start.get(f"{{{WML_NS}}}name") != bookmark_name
        or end.tag != f"{{{WML_NS}}}bookmarkEnd"
        or tuple(end.attrib) != (f"{{{WML_NS}}}id",)
        or end.get(f"{{{WML_NS}}}id") != start.get(f"{{{WML_NS}}}id")
        or len(start) != 0
        or len(end) != 0
        or start.text is not None
        or end.text is not None
        or start.tail is not None
        or end.tail is not None
    ):
        raise ResolvedNumberingOoxmlError("caption bookmark range is not canonical")
    proof = prove_bookmark_name(bookmark_inventory, bookmark_name, scope_element=paragraph_element)
    if not proof.valid or proof.start is None or proof.end is None:
        raise ResolvedNumberingOoxmlError("caption bookmark is not globally proven")
    if proof.start.element is not start or proof.end.element is not end:
        raise ResolvedNumberingOoxmlError("caption bookmark does not bind its expected number range")
    return end_index + 1


def _elements_equal(left: Any, right: Any) -> bool:
    return (
        left.tag == right.tag
        and tuple(left.attrib.items()) == tuple(right.attrib.items())
        and left.text == right.text
        and left.tail == right.tail
        and len(left) == len(right)
        and all(_elements_equal(left_child, right_child) for left_child, right_child in zip(left, right, strict=True))
    )


def _clear_paragraph_payload(paragraph_element: Any) -> None:
    for child in list(paragraph_element):
        if child.tag != f"{{{WML_NS}}}pPr":
            paragraph_element.remove(child)


def _append_zero_width_bookmark(container: Any, name: str, bookmark_id: str) -> None:
    container.extend((_bookmark("start", name, bookmark_id), _bookmark("end", name, bookmark_id)))


def _bookmark(kind: str, name: str, bookmark_id: str) -> Any:
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    element = OxmlElement("w:bookmarkStart" if kind == "start" else "w:bookmarkEnd")
    element.set(qn("w:id"), bookmark_id)
    if kind == "start":
        element.set(qn("w:name"), name)
    return element


def _require_bookmark_id(bookmark_id: str | None) -> str:
    if bookmark_id is None:
        raise ResolvedNumberingOoxmlError("bookmark name requires an allocated bookmark ID")
    return bookmark_id


def _insert_num_pr(p_pr: Any, num_pr: Any) -> None:
    before = {"pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl"}
    for index, child in enumerate(p_pr):
        if etree.QName(child).localname not in before:
            p_pr.insert(index, num_pr)
            return
    p_pr.append(num_pr)


def _numbering_ids(root: Any) -> tuple[set[int], set[int]]:
    def collect(name: str, attribute: str) -> set[int]:
        output: set[int] = set()
        for element in root.findall(f"{{{WML_NS}}}{name}"):
            raw = element.get(f"{{{WML_NS}}}{attribute}")
            with contextlib.suppress(TypeError, ValueError):
                output.add(int(raw))
        return output

    return collect("abstractNum", "abstractNumId"), collect("num", "numId")


def _existing_numbering_ids(document: Any) -> tuple[set[int], set[int]]:
    """Inventory every live Word numbering definition and reference.

    Templates may carry dangling ``numId`` references in a story or style.  A
    newly allocated instance with that ID would silently activate those
    paragraphs, so references reserve IDs even when their definition is absent.
    """

    abstract_ids: set[int] = set()
    num_ids: set[int] = set()
    try:
        parts = tuple(document.part.package.parts)
    except (AttributeError, TypeError) as exc:
        raise ResolvedNumberingOoxmlError("template package parts cannot be inventoried") from exc
    parser = etree.XMLParser(resolve_entities=False, no_network=True, remove_blank_text=False)
    for part in parts:
        partname = str(getattr(part, "partname", ""))
        is_word_part = str(getattr(part, "content_type", "")).startswith(
            "application/vnd.openxmlformats-officedocument.wordprocessingml."
        )
        if not is_word_part and not (partname.startswith("/word/") and partname.endswith(".xml")):
            continue
        root = getattr(part, "element", None)
        if root is None:
            root = getattr(part, "_element", None)
        if root is None:
            try:
                root = etree.fromstring(part.blob, parser)
            except (AttributeError, etree.XMLSyntaxError, TypeError, ValueError) as exc:
                raise ResolvedNumberingOoxmlError(f"Word XML part cannot be inventoried: {partname}") from exc
        for element in root.iter():
            if element.tag == f"{{{WML_NS}}}abstractNum":
                abstract_ids.add(_word_id(element, "abstractNumId", partname))
            elif element.tag == f"{{{WML_NS}}}abstractNumId":
                abstract_ids.add(_word_id(element, "val", partname))
            elif element.tag == f"{{{WML_NS}}}num":
                num_ids.add(_word_id(element, "numId", partname, minimum=1))
            elif element.tag == f"{{{WML_NS}}}numId":
                num_ids.add(_word_id(element, "val", partname))
    return abstract_ids, num_ids


def _word_id(element: Any, name: str, partname: str, *, minimum: int = 0) -> int:
    raw = element.get(f"{{{WML_NS}}}{name}")
    try:
        value = int(raw)
    except (TypeError, ValueError) as exc:
        raise ResolvedNumberingOoxmlError(f"invalid numbering ID in {partname}") from exc
    if not minimum <= value <= 2_147_483_647:
        raise ResolvedNumberingOoxmlError(f"numbering ID is outside the portable range in {partname}")
    return value


def _numbering_ids_from_elements(projection: Any) -> tuple[set[int], set[int]]:
    abstract = {int(item.get(f"{{{WML_NS}}}abstractNumId")) for item in projection.abstract_nums}
    nums = {int(item.get(f"{{{WML_NS}}}numId")) for item in projection.nums}
    return abstract, nums


def _lowest_unused(values: set[int], *, start: int) -> int:
    candidate = start
    while candidate in values:
        candidate += 1
    return candidate


def _rewrite_package(
    path: Path,
    infos: list[ZipInfo],
    data: dict[str, bytes],
    *,
    new_names: tuple[str, ...],
) -> None:
    descriptor, temporary_name = tempfile.mkstemp(prefix=f".{path.name}.", suffix=".tmp", dir=path.parent)
    os.close(descriptor)
    temporary = Path(temporary_name)
    try:
        with ZipFile(temporary, "w") as output:
            existing = {item.filename for item in infos}
            for item in infos:
                output.writestr(item, data[item.filename])
            for name in new_names:
                if name in existing:
                    continue
                info = ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
                info.compress_type = ZIP_DEFLATED
                info.external_attr = 0o600 << 16
                output.writestr(info, data[name])
        replace_authorized_file(temporary, path)
    finally:
        temporary.unlink(missing_ok=True)

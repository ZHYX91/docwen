"""Proof-only DOCX Heading numbering extraction for the v4 neutral boundary.

Visible paragraph text is always returned unchanged.  A separate numbering
fact is produced only from a direct paragraph ``numPr`` backed by a valid OPC
numbering part, instance, abstract definition, and level.  This module never
parses, strips, or rewrites a number-looking text prefix.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any
from zipfile import ZipFile

import lxml.etree as etree

from docwen_core.docx_numbering_ooxml import (
    CONTENT_TYPES_NS,
    NUMBERING_CONTENT_TYPE,
    NUMBERING_REL_TYPE,
    RELS_NS,
    WML_NS,
)

AMBIGUOUS_VISIBLE_PREFIX_DIAGNOSTIC = "docwen.docx.numbering.ambiguous_number_prefix"
_MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"

_PORTABLE_HEADING_FORMATS = frozenset(
    {
        "chineseCounting",
        "chineseCountingThousand",
        "decimal",
        "decimalFullWidth",
        "decimalEnclosedCircleChinese",
        "upperLetter",
        "lowerLetter",
        "upperRoman",
        "lowerRoman",
    }
)


class DocxNumberingImportError(ValueError):
    """DOCX numbering evidence is malformed or cannot prove one fact."""


@dataclass(frozen=True, slots=True)
class ProvenHeadingNumbering:
    """One list fact separated from, but never inserted into, Heading text."""

    num_id: int
    abstract_num_id: int
    level: int
    number_format: str
    level_text: str
    start: int
    suffix: str
    style_id: str | None


@dataclass(frozen=True, slots=True)
class NumberingImportDiagnostic:
    code: str
    message: str


@dataclass(frozen=True, slots=True)
class ProofOnlyHeadingImport:
    authored_text: str
    numbering: ProvenHeadingNumbering | None
    diagnostics: tuple[NumberingImportDiagnostic, ...]


@dataclass(frozen=True, slots=True)
class _Level:
    number_format: str
    level_text: str
    start: int
    suffix: str
    style_id: str | None


@dataclass(frozen=True, slots=True)
class _Instance:
    abstract_num_id: int
    start_overrides: tuple[tuple[int, int], ...]


class HeadingNumberingProofIndex:
    """Strict index over one authenticated package numbering part."""

    def __init__(
        self,
        *,
        levels: dict[tuple[int, int], _Level] | None = None,
        instances: dict[int, _Instance] | None = None,
    ) -> None:
        self._levels = levels or {}
        self._instances = instances or {}

    @classmethod
    def load(cls, path: str | Path) -> HeadingNumberingProofIndex:
        package_path = Path(path)
        parser = etree.XMLParser(resolve_entities=False, no_network=True, remove_blank_text=False)
        with ZipFile(package_path) as package:
            infos = package.infolist()
            if len({item.filename for item in infos}) != len(infos):
                raise DocxNumberingImportError("DOCX contains duplicate ZIP members")
            names = {item.filename for item in infos}
            required = {"word/_rels/document.xml.rels", "[Content_Types].xml"}
            if not required <= names:
                raise DocxNumberingImportError("DOCX lacks required OPC support parts")
            rels = etree.fromstring(package.read("word/_rels/document.xml.rels"), parser)
            content_types = etree.fromstring(package.read("[Content_Types].xml"), parser)
            if "word/numbering.xml" not in names:
                _prove_numbering_opc_support(rels, content_types, present=False)
                return cls()
            _prove_numbering_opc_support(rels, content_types, present=True)
            numbering = etree.fromstring(package.read("word/numbering.xml"), parser)
        levels, instances = _parse_numbering_part(numbering)
        return cls(levels=levels, instances=instances)

    def prove_paragraph(self, paragraph: Any) -> ProvenHeadingNumbering | None:
        """Return one proven direct binding, or ``None`` for no numbering."""

        p_pr = paragraph._p.pPr
        if p_pr is None:
            return None
        num_pr_nodes = p_pr.findall(f"{{{WML_NS}}}numPr")
        if not num_pr_nodes:
            return None
        if len(num_pr_nodes) != 1:
            raise DocxNumberingImportError("Heading paragraph has duplicate numPr bindings")
        num_pr = num_pr_nodes[0]
        num_ids = num_pr.findall(f"{{{WML_NS}}}numId")
        levels = num_pr.findall(f"{{{WML_NS}}}ilvl")
        if len(num_ids) != 1 or len(levels) != 1:
            raise DocxNumberingImportError("Heading numPr lacks one exact numId and ilvl")
        num_id = _integer_attribute(num_ids[0], "val", minimum=0)
        ilvl = _integer_attribute(levels[0], "val", minimum=0, maximum=8)
        if num_id == 0:
            return None
        instance = self._instances.get(num_id)
        if instance is None:
            raise DocxNumberingImportError("Heading numPr references a missing numbering instance")
        definition = self._levels.get((instance.abstract_num_id, ilvl))
        if definition is None:
            raise DocxNumberingImportError("Heading numPr references a missing numbering level")
        if definition.number_format not in _PORTABLE_HEADING_FORMATS:
            raise DocxNumberingImportError("Heading numbering format is not a portable structured number")
        start = dict(instance.start_overrides).get(ilvl, definition.start)
        return ProvenHeadingNumbering(
            num_id=num_id,
            abstract_num_id=instance.abstract_num_id,
            level=ilvl + 1,
            number_format=definition.number_format,
            level_text=definition.level_text,
            start=start,
            suffix=definition.suffix,
            style_id=definition.style_id,
        )


def import_heading_without_source_mutation(
    paragraph: Any,
    proof_index: HeadingNumberingProofIndex,
    *,
    suspected_visible_prefix: bool = False,
) -> ProofOnlyHeadingImport:
    """Separate proven list semantics while preserving all visible text.

    ``suspected_visible_prefix`` is an observation supplied by the caller; it
    is never parsed into a number here and never authorizes text cleanup.
    """

    authored = paragraph.text
    numbering = proof_index.prove_paragraph(paragraph)
    diagnostics: tuple[NumberingImportDiagnostic, ...] = ()
    if numbering is None and suspected_visible_prefix:
        diagnostics = (
            NumberingImportDiagnostic(
                AMBIGUOUS_VISIBLE_PREFIX_DIAGNOSTIC,
                "Visible Heading prefix has no authenticated Word list semantics and remains authored text.",
            ),
        )
    return ProofOnlyHeadingImport(authored, numbering, diagnostics)


def _prove_numbering_opc_support(rels: Any, content_types: Any, *, present: bool) -> None:
    if rels.tag != f"{{{RELS_NS}}}Relationships" or content_types.tag != f"{{{CONTENT_TYPES_NS}}}Types":
        raise DocxNumberingImportError("DOCX OPC support roots are invalid")
    relationships = [
        item for item in rels if item.get("Type") == NUMBERING_REL_TYPE or item.get("Target") == "numbering.xml"
    ]
    overrides = [item for item in content_types if item.get("PartName") == "/word/numbering.xml"]
    if not present:
        if relationships or overrides:
            raise DocxNumberingImportError("DOCX advertises an absent numbering part")
        return
    if (
        len(relationships) != 1
        or relationships[0].get("Type") != NUMBERING_REL_TYPE
        or relationships[0].get("Target") != "numbering.xml"
        or len(overrides) != 1
        or overrides[0].get("ContentType") != NUMBERING_CONTENT_TYPE
    ):
        raise DocxNumberingImportError("numbering part lacks exact OPC support")


def _parse_numbering_part(root: Any) -> tuple[dict[tuple[int, int], _Level], dict[int, _Instance]]:
    if root.tag != f"{{{WML_NS}}}numbering":
        raise DocxNumberingImportError("numbering.xml root is invalid")
    _reject_unmodeled_numbering_topology(root)
    levels: dict[tuple[int, int], _Level] = {}
    abstract_ids: set[int] = set()
    for abstract in root.findall(f"{{{WML_NS}}}abstractNum"):
        abstract_id = _integer_attribute(abstract, "abstractNumId", minimum=0)
        if abstract_id in abstract_ids:
            raise DocxNumberingImportError("numbering.xml repeats an abstractNumId")
        abstract_ids.add(abstract_id)
        for node in abstract.findall(f"{{{WML_NS}}}lvl"):
            ilvl = _integer_attribute(node, "ilvl", minimum=0, maximum=8)
            key = (abstract_id, ilvl)
            if key in levels:
                raise DocxNumberingImportError("numbering.xml repeats an abstract level")
            levels[key] = _parse_level(node)

    instances: dict[int, _Instance] = {}
    for node in root.findall(f"{{{WML_NS}}}num"):
        num_id = _integer_attribute(node, "numId", minimum=1)
        if num_id in instances:
            raise DocxNumberingImportError("numbering.xml repeats a numId")
        references = node.findall(f"{{{WML_NS}}}abstractNumId")
        if len(references) != 1:
            raise DocxNumberingImportError("numbering instance lacks one abstractNumId")
        abstract_id = _integer_attribute(references[0], "val", minimum=0)
        if abstract_id not in abstract_ids:
            raise DocxNumberingImportError("numbering instance references a missing abstractNum")
        overrides: list[tuple[int, int]] = []
        for override in node.findall(f"{{{WML_NS}}}lvlOverride"):
            ilvl = _integer_attribute(override, "ilvl", minimum=0, maximum=8)
            if [child.tag for child in override] != [f"{{{WML_NS}}}startOverride"]:
                raise DocxNumberingImportError("numbering level override must contain only one startOverride")
            values = override.findall(f"{{{WML_NS}}}startOverride")
            if len(values) != 1:
                raise DocxNumberingImportError("numbering level override lacks one startOverride")
            overrides.append((ilvl, _integer_attribute(values[0], "val", minimum=1)))
        if [item[0] for item in overrides] != sorted({item[0] for item in overrides}):
            raise DocxNumberingImportError("numbering level overrides are duplicate or unordered")
        instances[num_id] = _Instance(abstract_id, tuple(overrides))
    return levels, instances


def _reject_unmodeled_numbering_topology(root: Any) -> None:
    expected_parents = {
        f"{{{WML_NS}}}abstractNum": f"{{{WML_NS}}}numbering",
        f"{{{WML_NS}}}num": f"{{{WML_NS}}}numbering",
        f"{{{WML_NS}}}lvl": f"{{{WML_NS}}}abstractNum",
        f"{{{WML_NS}}}abstractNumId": f"{{{WML_NS}}}num",
        f"{{{WML_NS}}}lvlOverride": f"{{{WML_NS}}}num",
        f"{{{WML_NS}}}startOverride": f"{{{WML_NS}}}lvlOverride",
        f"{{{WML_NS}}}start": f"{{{WML_NS}}}lvl",
        f"{{{WML_NS}}}numFmt": f"{{{WML_NS}}}lvl",
        f"{{{WML_NS}}}lvlText": f"{{{WML_NS}}}lvl",
        f"{{{WML_NS}}}suff": f"{{{WML_NS}}}lvl",
        f"{{{WML_NS}}}pStyle": f"{{{WML_NS}}}lvl",
    }
    for element in root.iter():
        if isinstance(element.tag, str) and element.tag.startswith(f"{{{_MC_NS}}}"):
            raise DocxNumberingImportError("numbering.xml contains unmodeled markup-compatibility topology")
        expected_parent = expected_parents.get(element.tag)
        if expected_parent is None:
            continue
        parent = element.getparent()
        if parent is None or parent.tag != expected_parent:
            raise DocxNumberingImportError("numbering.xml contains a modeled node outside its exact ancestry")


def _parse_level(node: Any) -> _Level:
    starts = node.findall(f"{{{WML_NS}}}start")
    formats = node.findall(f"{{{WML_NS}}}numFmt")
    texts = node.findall(f"{{{WML_NS}}}lvlText")
    suffixes = node.findall(f"{{{WML_NS}}}suff")
    styles = node.findall(f"{{{WML_NS}}}pStyle")
    if len(starts) > 1 or len(formats) != 1 or len(texts) != 1 or len(suffixes) > 1 or len(styles) > 1:
        raise DocxNumberingImportError("numbering level has duplicate or missing core properties")
    number_format = formats[0].get(f"{{{WML_NS}}}val")
    level_text = texts[0].get(f"{{{WML_NS}}}val")
    if not number_format or level_text is None:
        raise DocxNumberingImportError("numbering level has an empty format or level text")
    start = _integer_attribute(starts[0], "val", minimum=1) if starts else 1
    suffix = suffixes[0].get(f"{{{WML_NS}}}val") if suffixes else "tab"
    if suffix not in {"nothing", "space", "tab"}:
        raise DocxNumberingImportError("numbering level suffix is invalid")
    style_id = styles[0].get(f"{{{WML_NS}}}val") if styles else None
    if style_id == "":
        raise DocxNumberingImportError("numbering level has an empty pStyle")
    return _Level(number_format, level_text, start, suffix, style_id)


def _integer_attribute(
    element: Any,
    name: str,
    *,
    minimum: int,
    maximum: int = 2_147_483_647,
) -> int:
    raw = element.get(f"{{{WML_NS}}}{name}")
    try:
        value = int(raw)
    except (TypeError, ValueError) as exc:
        raise DocxNumberingImportError(f"numbering attribute {name} is not an integer") from exc
    if not minimum <= value <= maximum or str(value) != raw:
        raise DocxNumberingImportError(f"numbering attribute {name} is outside the portable range")
    return value


__all__ = [
    "AMBIGUOUS_VISIBLE_PREFIX_DIAGNOSTIC",
    "DocxNumberingImportError",
    "HeadingNumberingProofIndex",
    "NumberingImportDiagnostic",
    "ProofOnlyHeadingImport",
    "ProvenHeadingNumbering",
    "import_heading_without_source_mutation",
]

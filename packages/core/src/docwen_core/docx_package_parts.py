"""Relationship-owned DOCX parts, independent of their conventional filenames.

The package and loaded-document adapters share roles and URI validation. Neither
adapter creates parts or treats an orphan member as an authenticated document part.
"""

from __future__ import annotations

import posixpath
from collections.abc import Iterator, Mapping
from dataclasses import dataclass
from typing import Any
from urllib.parse import unquote, urlsplit
from zipfile import ZipFile

from lxml import etree

WML_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
_OFFICE_RELS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
_WORD_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml."


class DocxPartError(ValueError):
    """A DOCX package cannot identify an unambiguous internal part."""


@dataclass(frozen=True, slots=True)
class PartRole:
    name: str
    relationship_type: str
    content_type: str
    root_name: str
    multiple: bool = False


DOCUMENT = PartRole("document", _OFFICE_RELS + "officeDocument", _WORD_TYPE + "document.main+xml", "document")
HEADER = PartRole("header", _OFFICE_RELS + "header", _WORD_TYPE + "header+xml", "hdr", True)
FOOTER = PartRole("footer", _OFFICE_RELS + "footer", _WORD_TYPE + "footer+xml", "ftr", True)
STYLES = PartRole("styles", _OFFICE_RELS + "styles", _WORD_TYPE + "styles+xml", "styles")
STYLES_EFFECTS = PartRole(
    "stylesWithEffects",
    "http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects",
    "application/vnd.ms-word.stylesWithEffects+xml",
    "styles",
)
NUMBERING = PartRole("numbering", _OFFICE_RELS + "numbering", _WORD_TYPE + "numbering+xml", "numbering")
FOOTNOTES = PartRole("footnotes", _OFFICE_RELS + "footnotes", _WORD_TYPE + "footnotes+xml", "footnotes")
ENDNOTES = PartRole("endnotes", _OFFICE_RELS + "endnotes", _WORD_TYPE + "endnotes+xml", "endnotes")
COMMENTS = PartRole("comments", _OFFICE_RELS + "comments", _WORD_TYPE + "comments+xml", "comments")


@dataclass(frozen=True, slots=True)
class RelatedPart:
    name: str
    element: etree._Element


class _PackageMembers(Mapping[str, bytes]):
    """Read only selected XML while the caller owns the open ZIP handle."""

    def __init__(self, package: ZipFile, names: tuple[str, ...]) -> None:
        self._package = package
        self._names = dict.fromkeys(names)

    def __getitem__(self, name: str) -> bytes:
        return self._package.read(name)

    def __iter__(self) -> Iterator[str]:
        return iter(self._names)

    def __len__(self) -> int:
        return len(self._names)

    def __contains__(self, name: object) -> bool:
        return name in self._names


def relationship_part_name(owner: str) -> str:
    """Return the relationships member for a package root or an exact part."""
    if not owner:
        return "_rels/.rels"
    directory, basename = posixpath.split(owner)
    return posixpath.join(directory, "_rels", basename + ".rels")


def resolve_part_target(owner: str, target: str) -> str:
    """Resolve an internal target relative to its owner without URI guessing."""
    try:
        parsed = urlsplit(target)
    except ValueError as exc:
        raise DocxPartError("Relationship target is not a valid part URI") from exc
    if (
        not target
        or parsed.scheme
        or parsed.netloc
        or parsed.query
        or parsed.fragment
        or "\\" in target
        or any(ord(character) <= 32 for character in target)
        or target.endswith("/")
    ):
        raise DocxPartError("Relationship target is not an internal part URI")
    for segment in target.split("/"):
        decoded = unquote(segment)
        if "/" in decoded or "\\" in decoded or (decoded in {".", ".."} and decoded != segment):
            raise DocxPartError("Relationship target contains an ambiguous URI segment")
    name = posixpath.normpath(
        target.lstrip("/") if target.startswith("/") else posixpath.join(posixpath.dirname(owner), target)
    )
    if name in {"", ".", ".."} or name.startswith("../"):
        raise DocxPartError("Relationship target escapes the package")
    return name


def _parse(raw: bytes, tag: str, description: str) -> etree._Element:
    try:
        root = etree.fromstring(raw, etree.XMLParser(resolve_entities=False, no_network=True))
    except etree.XMLSyntaxError as exc:
        raise DocxPartError(f"Invalid {description} XML") from exc
    if root.tag != tag:
        raise DocxPartError(f"Invalid {description} root")
    return root


class DocxPackageParts:
    """Read-only role lookup over the exact package members supplied by a caller."""

    def __init__(self, members: Mapping[str, bytes]) -> None:
        self.members = members
        if "[Content_Types].xml" not in members:
            raise DocxPartError("DOCX lacks content types")
        self._types = _parse(members["[Content_Types].xml"], f"{{{CONTENT_TYPES_NS}}}Types", "content types")

    @classmethod
    def from_package(cls, package: ZipFile) -> DocxPackageParts:
        names = package.namelist()
        if len(names) != len(set(names)):
            raise DocxPartError("DOCX contains duplicate ZIP members")
        return cls(_PackageMembers(package, tuple(names)))

    def content_type(self, name: str) -> str | None:
        overrides = [
            item
            for item in self._types.findall(f"{{{CONTENT_TYPES_NS}}}Override")
            if item.get("PartName") == f"/{name}"
        ]
        if len(overrides) > 1:
            raise DocxPartError("Part has duplicate content-type overrides")
        if overrides:
            return overrides[0].get("ContentType")
        extension = posixpath.basename(name).rsplit(".", 1)[-1]
        defaults = [
            item
            for item in self._types.findall(f"{{{CONTENT_TYPES_NS}}}Default")
            if item.get("Extension", "").casefold() == extension.casefold()
        ]
        if len(defaults) > 1:
            raise DocxPartError("Part has ambiguous default content types")
        return defaults[0].get("ContentType") if defaults else None

    def relationships(self, owner: str) -> etree._Element:
        name = relationship_part_name(owner)
        if name not in self.members:
            return etree.Element(f"{{{RELS_NS}}}Relationships", nsmap={"pr": RELS_NS})
        root = _parse(self.members[name], f"{{{RELS_NS}}}Relationships", "relationships")
        elements = [item for item in root if isinstance(item.tag, str)]
        if any(item.tag != f"{{{RELS_NS}}}Relationship" for item in elements):
            raise DocxPartError("Relationships contain an invalid child")
        identities = [item.get("Id") for item in elements]
        if any(not identity for identity in identities) or len(identities) != len(set(identities)):
            raise DocxPartError("Relationships have missing or duplicate IDs")
        return root

    def related(self, owner: str, role: PartRole) -> tuple[RelatedPart, ...]:
        relations = [item for item in self.relationships(owner) if item.get("Type") == role.relationship_type]
        if not role.multiple and len(relations) > 1:
            raise DocxPartError(f"DOCX has multiple {role.name} relationships")
        result: dict[str, RelatedPart] = {}
        for relation in relations:
            if relation.get("TargetMode", "Internal") != "Internal":
                raise DocxPartError(f"{role.name} relationship is external")
            name = resolve_part_target(owner, relation.get("Target", ""))
            if name not in self.members:
                raise DocxPartError(f"{role.name} relationship targets a missing part")
            if self.content_type(name) != role.content_type:
                raise DocxPartError(f"{role.name} relationship has the wrong content type")
            result[name] = RelatedPart(name, _parse(self.members[name], f"{{{WML_NS}}}{role.root_name}", role.name))
        return tuple(result.values())

    def one(self, owner: str, role: PartRole, *, required: bool = False) -> RelatedPart | None:
        matches = self.related(owner, role)
        if len(matches) > 1 or (required and not matches):
            raise DocxPartError(f"DOCX requires exactly one {role.name} part")
        return matches[0] if matches else None

    def document(self) -> RelatedPart:
        result = self.one("", DOCUMENT, required=True)
        assert result is not None
        return result

    def stories(self) -> tuple[RelatedPart, ...]:
        document = self.document()
        return (document, *self.related(document.name, HEADER), *self.related(document.name, FOOTER))


def document_story_parts(document: Any) -> tuple[Any, ...]:
    """Enumerate existing loaded body/header/footer owners without creating parts."""
    owner = document.part
    owner_name = str(owner.partname).lstrip("/")
    result = {owner_name: owner}
    roles = {role.relationship_type: role for role in (HEADER, FOOTER)}
    for relation in owner.rels.values():
        role = roles.get(relation.reltype)
        if role is None:
            continue
        if relation.is_external:
            raise DocxPartError(f"{role.name} relationship is external")
        part = relation.target_part
        name = resolve_part_target(owner_name, relation.target_ref)
        if name != str(part.partname).lstrip("/"):
            raise DocxPartError(f"{role.name} relationship target differs from its loaded part")
        if part.content_type != role.content_type or getattr(part, "element", None) is None:
            raise DocxPartError(f"{role.name} relationship has the wrong content type")
        if part.element.tag != f"{{{WML_NS}}}{role.root_name}":
            raise DocxPartError(f"Invalid {role.name} root")
        result[name] = part
    return tuple(result.values())

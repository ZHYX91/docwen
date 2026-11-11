"""The relationship-owned registries and reference stories of a DOCX package."""

from __future__ import annotations

from dataclasses import dataclass
from zipfile import ZipFile

from docwen_core.docx_package_parts import (
    COMMENTS,
    ENDNOTES,
    FOOTER,
    FOOTNOTES,
    HEADER,
    NUMBERING,
    STYLES,
    STYLES_EFFECTS,
    DocxPackageParts,
    RelatedPart,
)


@dataclass(frozen=True, slots=True)
class StylePackageParts:
    primary: RelatedPart
    effects: RelatedPart | None
    references: tuple[RelatedPart, ...]

    @classmethod
    def read(cls, archive: ZipFile) -> StylePackageParts:
        package = DocxPackageParts.from_package(archive)
        document = package.document()
        primary = package.one(document.name, STYLES, required=True)
        assert primary is not None
        references = [document]
        for role in (HEADER, FOOTER, COMMENTS, FOOTNOTES, ENDNOTES, NUMBERING):
            references.extend(package.related(document.name, role))
        return cls(primary, package.one(document.name, STYLES_EFFECTS), tuple(references))

"""Read-only protection inspection for XLSX-to-ODS delivery choices."""

from __future__ import annotations

import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Literal
from xml.etree import ElementTree

_COMPOUND_FILE_HEADER = bytes.fromhex("D0CF11E0A1B11AE1")
_MAX_PROTECTION_XML_BYTES = 8 * 1024 * 1024


@dataclass(frozen=True, slots=True)
class SpreadsheetProtectionInfo:
    """Conservative protection state for one OOXML workbook."""

    path: str
    status: Literal["none", "protected", "unknown"]
    protected_parts: tuple[str, ...] = ()

    @property
    def is_protected(self) -> bool:
        return self.status == "protected"

    @property
    def is_unknown(self) -> bool:
        """Return whether protection could not be inspected safely."""
        return self.status == "unknown"

    @property
    def requires_password(self) -> bool:
        """Return whether the workbook is an encrypted Office package."""
        return "encrypted-package" in self.protected_parts


def inspect_xlsx_protection(path_value: str) -> SpreadsheetProtectionInfo:
    """Inspect workbook/sheet protection without opening or mutating the file.

    Encrypted OOXML is stored in an OLE compound container and is treated as
    protected. Malformed, missing, or oversized inputs fail closed as unknown.
    """

    path = Path(path_value)
    if not path.is_file():
        return SpreadsheetProtectionInfo(path=str(path), status="unknown")
    try:
        with path.open("rb") as stream:
            if stream.read(len(_COMPOUND_FILE_HEADER)) == _COMPOUND_FILE_HEADER:
                return SpreadsheetProtectionInfo(
                    path=str(path),
                    status="protected",
                    protected_parts=("encrypted-package",),
                )
        with zipfile.ZipFile(path) as package:
            member_names = {
                name
                for name in package.namelist()
                if name == "xl/workbook.xml" or (name.startswith("xl/worksheets/") and name.endswith(".xml"))
            }
            protected_parts: list[str] = []
            for name in sorted(member_names):
                info = package.getinfo(name)
                if info.file_size > _MAX_PROTECTION_XML_BYTES:
                    return SpreadsheetProtectionInfo(path=str(path), status="unknown")
                root = ElementTree.fromstring(package.read(name))
                # Protection is a direct workbook/worksheet child. Walking
                # every cell in Python makes large sheets needlessly expensive.
                if any(_protection_enabled(element) for element in root):
                    protected_parts.append(name)
    except (ElementTree.ParseError, OSError, RuntimeError, ValueError, zipfile.BadZipFile):
        return SpreadsheetProtectionInfo(path=str(path), status="unknown")
    return SpreadsheetProtectionInfo(
        path=str(path),
        status="protected" if protected_parts else "none",
        protected_parts=tuple(protected_parts),
    )


def _protection_enabled(element: ElementTree.Element) -> bool:
    tag = element.tag.rsplit("}", 1)[-1]
    flags = (
        ("sheet",)
        if tag == "sheetProtection"
        else ("lockStructure", "lockWindows", "lockRevision")
        if tag == "workbookProtection"
        else ()
    )
    values = [element.get(flag, "false").lower() for flag in flags]
    if any(value not in {"true", "false", "1", "0"} for value in values):
        raise ValueError("Invalid protection flag")
    return any(value in {"true", "1"} for value in values)


__all__ = ["SpreadsheetProtectionInfo", "inspect_xlsx_protection"]

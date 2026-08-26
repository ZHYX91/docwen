"""Bound generated ODS repeated grid spans to the producing XLSX dimensions."""

from __future__ import annotations

import os
import posixpath
import re
import shutil
import tempfile
import xml.sax
from dataclasses import dataclass
from pathlib import Path
from typing import IO, cast
from xml.sax.handler import ContentHandler, feature_namespaces
from xml.sax.saxutils import XMLGenerator
from xml.sax.xmlreader import AttributesNSImpl
from zipfile import ZIP_DEFLATED, BadZipFile, ZipFile, ZipInfo

from lxml import etree

_ODF_TABLE_NS = "urn:oasis:names:tc:opendocument:xmlns:table:1.0"
_OOXML_MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_OOXML_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_TABLE = _ODF_TABLE_NS
_TABLE_NAME = (_TABLE, "name")
_ROW_REPEAT = (_TABLE, "number-rows-repeated")
_COLUMN_REPEAT = (_TABLE, "number-columns-repeated")
_TABLE_ELEMENT = (_TABLE, "table")
_ROW_ELEMENT = (_TABLE, "table-row")
_COLUMN_ELEMENT = (_TABLE, "table-column")
_CELL_ELEMENTS = {
    (_TABLE, "table-cell"),
    (_TABLE, "covered-table-cell"),
}
_CELL_REFERENCE_RE = re.compile(r"\$?([A-Z]{1,3})\$?([1-9][0-9]*)$")


@dataclass(frozen=True, slots=True)
class OdsGridTrim:
    """One ODS table whose repeated row or column grid was bounded."""

    sheet_name: str
    max_row: int
    max_column: int
    removed_repeated_rows: int
    removed_repeated_columns: int


@dataclass(frozen=True, slots=True)
class OdsGridCompaction:
    """Result of a generated-ODS grid-bound pass."""

    output_path: str
    trimmed_sheets: tuple[OdsGridTrim, ...] = ()

    @property
    def changed(self) -> bool:
        return bool(self.trimmed_sheets)

    @property
    def removed_repeated_rows(self) -> int:
        return sum(item.removed_repeated_rows for item in self.trimmed_sheets)

    @property
    def removed_repeated_columns(self) -> int:
        return sum(item.removed_repeated_columns for item in self.trimmed_sheets)


class OdsGridCompactionError(ValueError):
    """The ODS or XLSX boundary package could not be safely interpreted."""


def _positive_int(value: str | None) -> int:
    try:
        return max(1, int(value or "1"))
    except ValueError:
        return 1


def _column_number(letters: str) -> int:
    result = 0
    for character in letters:
        result = result * 26 + ord(character) - ord("A") + 1
    return result


def _dimension_bound(reference: str) -> tuple[int, int]:
    endpoint = reference.split(":", 1)[-1].strip()
    match = _CELL_REFERENCE_RE.fullmatch(endpoint)
    if match is None:
        raise OdsGridCompactionError(f"Unsupported XLSX worksheet dimension: {reference!r}.")
    return int(match.group(2)), _column_number(match.group(1))


def _xlsx_sheet_bounds(path: str | Path) -> dict[str, tuple[int, int]]:
    try:
        with ZipFile(path) as package:
            workbook = etree.fromstring(
                package.read("xl/workbook.xml"),
                parser=etree.XMLParser(resolve_entities=False, no_network=True),
            )
            relationships = etree.fromstring(
                package.read("xl/_rels/workbook.xml.rels"),
                parser=etree.XMLParser(resolve_entities=False, no_network=True),
            )
            targets = {
                relationship.get("Id", ""): relationship.get("Target", "")
                for relationship in relationships.findall(f"./{{{_PACKAGE_REL_NS}}}Relationship")
                if str(relationship.get("Type", "")).endswith("/worksheet")
            }
            bounds: dict[str, tuple[int, int]] = {}
            for sheet in workbook.findall(f"./{{{_OOXML_MAIN_NS}}}sheets/{{{_OOXML_MAIN_NS}}}sheet"):
                relationship_id = sheet.get(f"{{{_OOXML_REL_NS}}}id", "")
                target = targets.get(relationship_id, "")
                if not target:
                    continue
                normalized_target = target.replace("\\", "/").lstrip("/")
                part = (
                    posixpath.normpath(normalized_target)
                    if normalized_target.startswith("xl/")
                    else posixpath.normpath(posixpath.join("xl", normalized_target))
                )
                worksheet = etree.fromstring(
                    package.read(part),
                    parser=etree.XMLParser(resolve_entities=False, no_network=True),
                )
                dimension = worksheet.find(f"./{{{_OOXML_MAIN_NS}}}dimension")
                reference = "A1" if dimension is None else str(dimension.get("ref", "A1"))
                bounds[str(sheet.get("name", ""))] = _dimension_bound(reference)
            return bounds
    except (BadZipFile, KeyError, etree.XMLSyntaxError, OSError) as exc:
        raise OdsGridCompactionError("The XLSX boundary package is invalid or incomplete.") from exc


def _is_semantic_cell(element: etree._Element) -> bool:
    ignored_attributes = {
        f"{{{_TABLE}}}style-name",
        f"{{{_TABLE}}}number-columns-repeated",
    }
    if any(name not in ignored_attributes for name in element.attrib):
        return True
    if (element.text or "").strip():
        return True
    return any(_has_string_tag(child) for child in element)


def _has_string_tag(element: etree._Element) -> bool:
    """Return whether an lxml node represents a named XML element."""
    tag = cast(object, element.tag)
    return isinstance(tag, str)


@dataclass(slots=True)
class _SemanticTableState:
    name: str
    logical_row: int = 0
    row_start: int = 0
    row_repeat: int = 1
    current_row_column: int = 0
    max_row: int = 1
    max_column: int = 1


def _ods_semantic_bounds(source: IO[bytes]) -> dict[str, tuple[int, int]]:
    states: list[_SemanticTableState] = []
    completed: dict[str, tuple[int, int]] = {}
    try:
        iterator = etree.iterparse(
            source,
            events=("start", "end"),
            resolve_entities=False,
            no_network=True,
            huge_tree=True,
        )
        for event, element in iterator:
            if not isinstance(element.tag, str):
                continue
            expanded_name = (etree.QName(element).namespace or "", etree.QName(element).localname)
            if event == "start" and expanded_name == _TABLE_ELEMENT:
                states.append(_SemanticTableState(str(element.get(f"{{{_TABLE}}}name", ""))))
            elif states and event == "start" and expanded_name == _ROW_ELEMENT:
                state = states[-1]
                state.row_start = state.logical_row + 1
                state.row_repeat = _positive_int(element.get(f"{{{_TABLE}}}number-rows-repeated"))
                state.logical_row += state.row_repeat
                state.current_row_column = 0
            elif states and event == "end" and expanded_name in _CELL_ELEMENTS:
                state = states[-1]
                repeat = _positive_int(element.get(f"{{{_TABLE}}}number-columns-repeated"))
                state.current_row_column += repeat
                if _is_semantic_cell(element):
                    state.max_row = max(
                        state.max_row,
                        state.row_start + state.row_repeat - 1,
                    )
                    state.max_column = max(
                        state.max_column,
                        state.current_row_column,
                    )
            elif states and event == "end" and expanded_name == _TABLE_ELEMENT:
                state = states.pop()
                if state.name:
                    completed[state.name] = (state.max_row, state.max_column)

            if event == "end":
                element.clear()
                parent = element.getparent()
                if parent is not None:
                    while element.getprevious() is not None:
                        del parent[0]
    except etree.XMLSyntaxError as exc:
        raise OdsGridCompactionError("The generated ODS content.xml is invalid.") from exc
    return completed


def _merged_bounds(
    xlsx_bounds: dict[str, tuple[int, int]],
    semantic_bounds: dict[str, tuple[int, int]],
) -> dict[str, tuple[int, int]]:
    return {
        name: (
            max(bound[0], semantic_bounds.get(name, (1, 1))[0]),
            max(bound[1], semantic_bounds.get(name, (1, 1))[1]),
        )
        for name, bound in xlsx_bounds.items()
    }


def _attributes_with_repeat(
    attributes: AttributesNSImpl,
    name: tuple[str, str],
    repeat: int,
) -> AttributesNSImpl:
    values = dict(attributes.items())
    qnames = {attribute_name: attributes.getQNameByName(attribute_name) for attribute_name in values}
    if repeat <= 1:
        values.pop(name, None)
        qnames.pop(name, None)
    else:
        values[name] = str(repeat)
        qnames.setdefault(name, f"table:{name[1]}")
    return AttributesNSImpl(values, qnames)


@dataclass(slots=True)
class _TableState:
    name: str
    bound: tuple[int, int] | None
    logical_row: int = 0
    logical_column_definition: int = 0
    current_row_column: int = 0
    removed_rows: int = 0
    removed_columns: int = 0


class _GridBoundHandler(ContentHandler):
    def __init__(
        self,
        output: IO[bytes],
        bounds: dict[str, tuple[int, int]],
    ) -> None:
        super().__init__()
        self._generator = XMLGenerator(output, encoding="UTF-8", short_empty_elements=False)
        self._bounds = bounds
        self._tables: list[_TableState] = []
        self._suppression_depth = 0
        self.trims: dict[str, _TableState] = {}

    def startDocument(self) -> None:
        self._generator.startDocument()

    def endDocument(self) -> None:
        self._generator.endDocument()

    def startPrefixMapping(self, prefix: str | None, uri: str) -> None:
        if self._suppression_depth == 0:
            self._generator.startPrefixMapping(prefix, uri)

    def endPrefixMapping(self, prefix: str | None) -> None:
        if self._suppression_depth == 0:
            self._generator.endPrefixMapping(prefix)

    def startElementNS(
        self,
        name: tuple[str | None, str],
        qname: str | None,
        attrs: AttributesNSImpl,
    ) -> None:
        if self._suppression_depth:
            self._suppression_depth += 1
            return

        if name == _TABLE_ELEMENT:
            table_name = str(attrs.get(_TABLE_NAME, ""))
            self._tables.append(_TableState(table_name, self._bounds.get(table_name)))
            self._generator.startElementNS(name, qname, attrs)
            return

        table = self._tables[-1] if self._tables else None
        if table is None or table.bound is None:
            self._generator.startElementNS(name, qname, attrs)
            return

        max_row, max_column = table.bound
        if name == _COLUMN_ELEMENT:
            repeat = _positive_int(attrs.get(_COLUMN_REPEAT))
            allowed = max(0, min(repeat, max_column - table.logical_column_definition))
            table.logical_column_definition += repeat
            table.removed_columns += repeat - allowed
            if allowed == 0:
                self._suppression_depth = 1
                return
            attrs = _attributes_with_repeat(attrs, _COLUMN_REPEAT, allowed)
        elif name == _ROW_ELEMENT:
            repeat = _positive_int(attrs.get(_ROW_REPEAT))
            allowed = max(0, min(repeat, max_row - table.logical_row))
            table.logical_row += repeat
            table.current_row_column = 0
            table.removed_rows += repeat - allowed
            if allowed == 0:
                self._suppression_depth = 1
                return
            attrs = _attributes_with_repeat(attrs, _ROW_REPEAT, allowed)
        elif name in _CELL_ELEMENTS:
            repeat = _positive_int(attrs.get(_COLUMN_REPEAT))
            allowed = max(0, min(repeat, max_column - table.current_row_column))
            table.current_row_column += repeat
            table.removed_columns += repeat - allowed
            if allowed == 0:
                self._suppression_depth = 1
                return
            attrs = _attributes_with_repeat(attrs, _COLUMN_REPEAT, allowed)

        self._generator.startElementNS(name, qname, attrs)

    def endElementNS(
        self,
        name: tuple[str | None, str],
        qname: str | None,
    ) -> None:
        if self._suppression_depth:
            self._suppression_depth -= 1
            return
        self._generator.endElementNS(name, qname)
        if name == _TABLE_ELEMENT and self._tables:
            table = self._tables.pop()
            if table.removed_rows or table.removed_columns:
                self.trims[table.name] = table

    def characters(self, content: str) -> None:
        if self._suppression_depth == 0:
            self._generator.characters(content)

    def ignorableWhitespace(self, whitespace: str) -> None:
        if self._suppression_depth == 0:
            self._generator.ignorableWhitespace(whitespace)

    def processingInstruction(self, target: str, data: str) -> None:
        if self._suppression_depth == 0:
            self._generator.processingInstruction(target, data)


def _rewrite_content(
    source: IO[bytes],
    output: IO[bytes],
    bounds: dict[str, tuple[int, int]],
) -> dict[str, _TableState]:
    handler = _GridBoundHandler(output, bounds)
    parser = xml.sax.make_parser()
    parser.setFeature(feature_namespaces, True)
    parser.setContentHandler(handler)
    try:
        parser.parse(source)
    except xml.sax.SAXException as exc:
        raise OdsGridCompactionError("The generated ODS content.xml is invalid.") from exc
    return handler.trims


def _clone_zip_info(info: ZipInfo) -> ZipInfo:
    cloned = ZipInfo(info.filename, date_time=info.date_time)
    cloned.compress_type = info.compress_type
    cloned.comment = info.comment
    cloned.extra = info.extra
    cloned.internal_attr = info.internal_attr
    cloned.external_attr = info.external_attr
    cloned.create_system = info.create_system
    cloned.flag_bits = info.flag_bits
    return cloned


def compact_generated_ods_grid(
    ods_path: str | Path,
    boundary_xlsx_path: str | Path,
) -> OdsGridCompaction:
    """Trim only generated repeated grid spans beyond the producing XLSX bounds."""

    ods = Path(ods_path)
    xlsx_bounds = _xlsx_sheet_bounds(boundary_xlsx_path)
    content_temp: Path | None = None
    package_temp: Path | None = None
    try:
        with ZipFile(ods) as source_package:
            if source_package.testzip() is not None:
                raise OdsGridCompactionError("The generated ODS package failed ZIP CRC.")
            try:
                content_info = source_package.getinfo("content.xml")
            except KeyError as exc:
                raise OdsGridCompactionError("The generated ODS package has no content.xml.") from exc
            with source_package.open(content_info) as semantic_source:
                semantic_bounds = _ods_semantic_bounds(semantic_source)
            bounds = _merged_bounds(xlsx_bounds, semantic_bounds)
            with tempfile.NamedTemporaryFile(
                prefix=f".{ods.name}.content-",
                suffix=".xml",
                dir=ods.parent,
                delete=False,
            ) as content_output:
                content_temp = Path(content_output.name)
                with source_package.open(content_info) as content_source:
                    trims = _rewrite_content(content_source, content_output.file, bounds)
            if not trims:
                return OdsGridCompaction(str(ods))

            with tempfile.NamedTemporaryFile(
                prefix=f".{ods.name}.package-",
                suffix=".ods",
                dir=ods.parent,
                delete=False,
            ) as package_output:
                package_temp = Path(package_output.name)
            with ZipFile(package_temp, "w", compression=ZIP_DEFLATED) as destination:
                for info in source_package.infolist():
                    cloned = _clone_zip_info(info)
                    if info.filename == "content.xml":
                        with content_temp.open("rb") as payload, destination.open(cloned, "w") as target:
                            shutil.copyfileobj(payload, target)
                    else:
                        with source_package.open(info) as payload, destination.open(cloned, "w") as target:
                            shutil.copyfileobj(payload, target)
        os.replace(package_temp, ods)
        package_temp = None
    except (BadZipFile, OSError) as exc:
        raise OdsGridCompactionError("The generated ODS package could not be compacted safely.") from exc
    finally:
        if content_temp is not None:
            content_temp.unlink(missing_ok=True)
        if package_temp is not None:
            package_temp.unlink(missing_ok=True)

    result = tuple(
        OdsGridTrim(
            sheet_name=name,
            max_row=state.bound[0] if state.bound is not None else 0,
            max_column=state.bound[1] if state.bound is not None else 0,
            removed_repeated_rows=state.removed_rows,
            removed_repeated_columns=state.removed_columns,
        )
        for name, state in sorted(trims.items())
    )
    return OdsGridCompaction(str(ods), result)

"""Delivery-first XLSX-to-ODS external-link and protection policy."""

from __future__ import annotations

import base64
import binascii
import hashlib
import hmac
import posixpath
import re
import struct
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import cast
from zipfile import ZIP_DEFLATED, BadZipFile, ZipFile, ZipInfo

from lxml import etree

_MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
_NS = {"m": _MAIN_NS, "r": _REL_NS, "pr": _PACKAGE_REL_NS, "ct": _CONTENT_TYPES_NS}
_EXTERNAL_BOOK_PATTERN = r"(?:\d+|[^\]]+\.(?:xlsx|xlsm|xlsb|xls|xltx|xltm|ods|csv))"
_EXTERNAL_FORMULA_RE = re.compile(
    rf"(?:'[^']*\[(?P<quoted_book>{_EXTERNAL_BOOK_PATTERN})\][^']*'!"
    rf"|\[(?P<plain_book>{_EXTERNAL_BOOK_PATTERN})\][^][!+\-*/^&=<>(),;:{{}}\r\n]+!)",
    re.IGNORECASE,
)
_TRUE_VALUES = frozenset({"1", "true", "on"})
_MAX_SPIN_COUNT = 10_000_000
_ODF_OFFICE_NS = "urn:oasis:names:tc:opendocument:xmlns:office:1.0"
_ODF_TABLE_NS = "urn:oasis:names:tc:opendocument:xmlns:table:1.0"


@dataclass(frozen=True, slots=True)
class XlsxOdsPolicyInspection:
    """Presence-only facts needed before a direct XLSX-to-ODS conversion."""

    external_formula_cells: tuple[str, ...]
    password_protected_elements: tuple[str, ...]
    unpassworded_protected_elements: tuple[str, ...]
    external_link_parts_present: bool = False
    external_defined_names: tuple[str, ...] = ()
    unsupported_external_references: tuple[str, ...] = ()
    fidelity_risk_counts: tuple[tuple[str, int], ...] = ()

    @property
    def needs_private_copy(self) -> bool:
        return bool(
            self.external_formula_cells
            or self.external_link_parts_present
            or self.external_defined_names
            or self.unsupported_external_references
            or self.password_protected_elements
        )


@dataclass(frozen=True, slots=True)
class XlsxOdsPreparation:
    """Losses deliberately applied to the backend-owned private XLSX copy."""

    output_path: str
    external_links_flattened: bool
    protection_removed: bool
    flattened_cached_values: tuple[str, ...] = ()
    removed_protection_elements: tuple[str, ...] = ()
    removed_external_defined_names: tuple[str, ...] = ()
    fidelity_risk_counts: tuple[tuple[str, int], ...] = ()


class XlsxOdsPolicyError(ValueError):
    """Typed admission failure safe to show without exposing a credential."""

    def __init__(self, diagnostic_code: str, message: str) -> None:
        super().__init__(message)
        self.diagnostic_code = diagnostic_code


def _parse_xml(payload: bytes) -> etree._Element:
    return etree.fromstring(payload, parser=etree.XMLParser(resolve_entities=False, no_network=True))


def _serialize_xml(root: etree._Element) -> bytes:
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8")


def _element_local_name(element: etree._Element | None) -> str | None:
    """Return a named XML element's local name, excluding comment-like nodes."""
    if element is None:
        return None
    tag = cast(object, element.tag)
    if not isinstance(tag, str):
        return None
    return etree.QName(tag).localname


def _has_external_reference(value: str, source_filename: str) -> bool:
    for match in _EXTERNAL_FORMULA_RE.finditer(value):
        workbook = match.group("quoted_book") or match.group("plain_book") or ""
        if workbook.isdigit() or workbook.casefold() != source_filename.casefold():
            return True
    return False


def _is_enabled(value: str | None) -> bool:
    return str(value or "").strip().lower() in _TRUE_VALUES


def _worksheet_parts(parts: dict[str, bytes]) -> list[tuple[str, str]]:
    workbook_payload = parts.get("xl/workbook.xml")
    rels_payload = parts.get("xl/_rels/workbook.xml.rels")
    if workbook_payload is None or rels_payload is None:
        raise XlsxOdsPolicyError("INVALID_XLSX_PACKAGE", "The XLSX package has no workbook relationship graph.")

    workbook = _parse_xml(workbook_payload)
    rels = _parse_xml(rels_payload)
    targets = {
        rel.get("Id", ""): rel.get("Target", "")
        for rel in rels.findall("./pr:Relationship", namespaces=_NS)
        if str(rel.get("Type", "")).endswith("/worksheet")
    }
    worksheets: list[tuple[str, str]] = []
    for sheet in workbook.findall("./m:sheets/m:sheet", namespaces=_NS):
        rel_id = sheet.get(f"{{{_REL_NS}}}id", "")
        target = targets.get(rel_id, "")
        if not target:
            continue
        normalized_target = target.replace("\\", "/").lstrip("/")
        normalized = (
            posixpath.normpath(normalized_target)
            if normalized_target.startswith("xl/")
            else posixpath.normpath(posixpath.join("xl", normalized_target))
        )
        worksheets.append((sheet.get("name", "worksheet"), normalized))
    return worksheets


def _protection_kind(element: etree._Element, *, workbook: bool) -> str:
    if workbook:
        active = any(_is_enabled(element.get(name)) for name in ("lockStructure", "lockWindows", "lockRevision"))
        hash_value = element.get("workbookHashValue")
        legacy_value = element.get("workbookPassword")
    else:
        active = _is_enabled(element.get("sheet"))
        hash_value = element.get("hashValue")
        legacy_value = element.get("password")
    if not active:
        return "inactive"
    if hash_value or legacy_value:
        return "password"
    return "unpassworded"


def _read_package(path: str | Path) -> tuple[list[ZipInfo], dict[str, bytes]]:
    try:
        with ZipFile(path) as package:
            infos = package.infolist()
            return infos, {info.filename: package.read(info.filename) for info in infos}
    except (BadZipFile, OSError, KeyError, etree.XMLSyntaxError) as exc:
        raise XlsxOdsPolicyError("INVALID_XLSX_PACKAGE", "The XLSX package cannot be inspected safely.") from exc


def _inspect_parts(parts: dict[str, bytes], *, source_filename: str) -> XlsxOdsPolicyInspection:
    workbook_payload = parts.get("xl/workbook.xml")
    if workbook_payload is None:
        raise XlsxOdsPolicyError("INVALID_XLSX_PACKAGE", "The XLSX package has no workbook part.")
    workbook = _parse_xml(workbook_payload)

    passworded: list[str] = []
    unpassworded: list[str] = []
    external_defined_names: list[str] = []
    retained_defined_name_count = 0
    for workbook_protection in workbook.findall(f"{{{_MAIN_NS}}}workbookProtection"):
        kind = _protection_kind(workbook_protection, workbook=True)
        if kind == "password":
            if "workbook" not in passworded:
                passworded.append("workbook")
        elif kind == "unpassworded" and "workbook" not in unpassworded:
            unpassworded.append("workbook")
    for defined_name in workbook.findall("./m:definedNames/m:definedName", namespaces=_NS):
        if _has_external_reference(defined_name.text or "", source_filename):
            external_defined_names.append(str(defined_name.get("name", "?")))
        else:
            retained_defined_name_count += 1

    external_cells: list[str] = []
    unsupported_external_references: list[str] = []
    data_validation_count = 0
    conditional_formatting_count = 0
    for sheet_name, part_name in _worksheet_parts(parts):
        payload = parts.get(part_name)
        if payload is None:
            raise XlsxOdsPolicyError("INVALID_XLSX_PACKAGE", f"The XLSX package is missing worksheet {sheet_name}.")
        worksheet = _parse_xml(payload)
        data_validation_count += sum(_element_local_name(element) == "dataValidation" for element in worksheet.iter())
        conditional_formatting_count += sum(
            _element_local_name(element) == "conditionalFormatting" for element in worksheet.iter()
        )
        protection = worksheet.find(f"{{{_MAIN_NS}}}sheetProtection")
        if protection is not None:
            kind = _protection_kind(protection, workbook=False)
            if kind == "password":
                passworded.append(sheet_name)
            elif kind == "unpassworded":
                unpassworded.append(sheet_name)
        for cell in worksheet.findall(".//m:c", namespaces=_NS):
            formula = cell.find(f"{{{_MAIN_NS}}}f")
            if formula is None:
                continue
            formula_text = formula.text if formula is not None else ""
            if _has_external_reference(formula_text or "", source_filename):
                external_cells.append(f"{sheet_name}!{cell.get('r', '?')}")
        for element in worksheet.iter():
            local_name = _element_local_name(element)
            if local_name is None:
                continue
            parent = element.getparent()
            is_cell_formula = local_name == "f" and _element_local_name(parent) == "c"
            if (
                local_name in {"f", "formula", "formula1", "formula2"}
                and not is_cell_formula
                and _has_external_reference(element.text or "", source_filename)
            ):
                unsupported_external_references.append(f"{sheet_name}:{local_name}")
            for key, value in element.attrib.items():
                if (
                    etree.QName(key).localname.lower() == "formula"
                    and isinstance(value, str)
                    and _has_external_reference(value, source_filename)
                ):
                    unsupported_external_references.append(f"{sheet_name}:@formula")

    worksheet_parts = {part_name for _sheet_name, part_name in _worksheet_parts(parts)}
    for part_name, payload in parts.items():
        if (
            not part_name.endswith(".xml")
            or part_name == "xl/workbook.xml"
            or part_name in worksheet_parts
            or part_name.startswith("xl/externalLinks/")
        ):
            continue
        try:
            root = _parse_xml(payload)
        except etree.XMLSyntaxError:
            continue
        for element in root.iter():
            local_name = _element_local_name(element)
            if local_name is None:
                continue
            if local_name in {"f", "formula", "formula1", "formula2"} and _has_external_reference(
                element.text or "", source_filename
            ):
                unsupported_external_references.append(f"{part_name}:{local_name}")
            for key, value in element.attrib.items():
                if (
                    etree.QName(key).localname.lower() == "formula"
                    and isinstance(value, str)
                    and _has_external_reference(value, source_filename)
                ):
                    unsupported_external_references.append(f"{part_name}:@formula")

    package_feature_counts = (
        (
            "charts",
            sum(re.fullmatch(r"xl/charts/chart\d+\.xml", name) is not None for name in parts),
        ),
        (
            "drawings",
            sum(re.fullmatch(r"xl/drawings/drawing\d+\.xml", name) is not None for name in parts),
        ),
        (
            "tables",
            sum(re.fullmatch(r"xl/tables/table\d+\.xml", name) is not None for name in parts),
        ),
        (
            "pivot_or_slicer_parts",
            sum(
                name.endswith(".xml")
                and "/_rels/" not in name
                and name.startswith(
                    (
                        "xl/pivotTables/",
                        "xl/pivotCache/",
                        "xl/slicers/",
                        "xl/slicerCaches/",
                    )
                )
                for name in parts
            ),
        ),
    )
    fidelity_risk_counts = tuple(
        (name, count)
        for name, count in (
            ("data_validations", data_validation_count),
            ("conditional_formatting_ranges", conditional_formatting_count),
            *package_feature_counts,
            ("defined_names", retained_defined_name_count),
        )
        if count
    )

    return XlsxOdsPolicyInspection(
        external_formula_cells=tuple(external_cells),
        password_protected_elements=tuple(passworded),
        unpassworded_protected_elements=tuple(unpassworded),
        external_link_parts_present=any(name.startswith("xl/externalLinks/") for name in parts),
        external_defined_names=tuple(external_defined_names),
        unsupported_external_references=tuple(unsupported_external_references),
        fidelity_risk_counts=fidelity_risk_counts,
    )


def inspect_xlsx_ods_policy(path: str | Path) -> XlsxOdsPolicyInspection:
    """Inspect direct-XLSX-to-ODS loss facts without opening any link target."""

    _infos, parts = _read_package(path)
    try:
        return _inspect_parts(parts, source_filename=Path(path).name)
    except (KeyError, etree.XMLSyntaxError, ValueError) as exc:
        if isinstance(exc, XlsxOdsPolicyError):
            raise
        raise XlsxOdsPolicyError("INVALID_XLSX_PACKAGE", "The XLSX package cannot be inspected safely.") from exc


def _modern_password_matches(
    password: str,
    *,
    algorithm_name: str | None,
    salt_value: str | None,
    spin_count: str | None,
    expected_hash: str | None,
) -> bool:
    if not all((algorithm_name, salt_value, spin_count, expected_hash)):
        return False
    normalized_algorithm = str(algorithm_name).replace("-", "").lower()
    try:
        count = int(str(spin_count))
        if count < 0 or count > _MAX_SPIN_COUNT:
            return False
        salt = base64.b64decode(str(salt_value), validate=True)
        expected = base64.b64decode(str(expected_hash), validate=True)
        digest = hashlib.new(normalized_algorithm, salt + password.encode("utf-16le")).digest()
        for index in range(count):
            digest = hashlib.new(normalized_algorithm, digest + struct.pack("<I", index)).digest()
    except (ValueError, TypeError, binascii.Error):
        return False
    return hmac.compare_digest(digest, expected)


def _legacy_password_matches(password: str, expected: str | None) -> bool:
    if not expected:
        return False
    from openpyxl.utils.protection import hash_password

    return hmac.compare_digest(hash_password(password).upper(), str(expected).upper())


def _protection_password_matches(element: etree._Element, password: str, *, workbook: bool) -> bool:
    prefix = "workbook" if workbook else ""
    if workbook:
        expected_hash = element.get("workbookHashValue")
        legacy_hash = element.get("workbookPassword")
        algorithm = element.get("workbookAlgorithmName")
        salt = element.get("workbookSaltValue")
        spins = element.get("workbookSpinCount")
    else:
        expected_hash = element.get("hashValue")
        legacy_hash = element.get("password")
        algorithm = element.get("algorithmName")
        salt = element.get("saltValue")
        spins = element.get("spinCount")
    if expected_hash:
        return _modern_password_matches(
            password,
            algorithm_name=algorithm,
            salt_value=salt,
            spin_count=spins,
            expected_hash=expected_hash,
        )
    if legacy_hash:
        return _legacy_password_matches(password, legacy_hash)
    raise AssertionError(f"{prefix or 'sheet'} protection is not passworded")


def _passworded_protection_nodes(
    parts: dict[str, bytes],
) -> list[tuple[str, etree._Element, etree._Element, bool]]:
    nodes: list[tuple[str, etree._Element, etree._Element, bool]] = []
    workbook = _parse_xml(parts["xl/workbook.xml"])
    for workbook_protection in workbook.findall(f"{{{_MAIN_NS}}}workbookProtection"):
        if _protection_kind(workbook_protection, workbook=True) == "password":
            nodes.append(("xl/workbook.xml", workbook, workbook_protection, True))
    for _sheet_name, part_name in _worksheet_parts(parts):
        worksheet = _parse_xml(parts[part_name])
        protection = worksheet.find(f"{{{_MAIN_NS}}}sheetProtection")
        if protection is not None and _protection_kind(protection, workbook=False) == "password":
            nodes.append((part_name, worksheet, protection, False))
    return nodes


def _flatten_external_links(
    parts: dict[str, bytes],
    *,
    source_filename: str,
) -> tuple[tuple[str, ...], tuple[str, ...]]:
    cached_values: list[str] = []
    for _sheet_name, part_name in _worksheet_parts(parts):
        worksheet = _parse_xml(parts[part_name])
        changed = False
        for cell in worksheet.findall(".//m:c", namespaces=_NS):
            formula = cell.find(f"{{{_MAIN_NS}}}f")
            if formula is None:
                continue
            formula_text = formula.text if formula is not None else ""
            if not _has_external_reference(formula_text or "", source_filename):
                continue
            cached = cell.find(f"{{{_MAIN_NS}}}v")
            if cached is None or cached.text is None:
                raise XlsxOdsPolicyError(
                    "EXTERNAL_LINK_CACHED_VALUE_MISSING",
                    "An external formula has no cached value, so an offline result cannot be delivered safely.",
                )
            assert formula is not None
            cell.remove(formula)
            cached_values.append(cached.text)
            changed = True
        if changed:
            parts[part_name] = _serialize_xml(worksheet)

    workbook = _parse_xml(parts["xl/workbook.xml"])
    removed_defined_names: list[str] = []
    for defined_names in workbook.findall(f"{{{_MAIN_NS}}}definedNames"):
        for defined_name in list(defined_names):
            if not _has_external_reference(defined_name.text or "", source_filename):
                continue
            removed_defined_names.append(str(defined_name.get("name", "?")))
            defined_names.remove(defined_name)
        if len(defined_names) == 0:
            workbook.remove(defined_names)
    external_references = workbook.find(f"{{{_MAIN_NS}}}externalReferences")
    if external_references is not None:
        workbook.remove(external_references)
    parts["xl/workbook.xml"] = _serialize_xml(workbook)

    rels = _parse_xml(parts["xl/_rels/workbook.xml.rels"])
    for relationship in list(rels):
        if str(relationship.get("Type", "")).endswith("/externalLink"):
            rels.remove(relationship)
    parts["xl/_rels/workbook.xml.rels"] = _serialize_xml(rels)

    content_types = _parse_xml(parts["[Content_Types].xml"])
    for override in list(content_types):
        if str(override.get("PartName", "")).startswith("/xl/externalLinks/"):
            content_types.remove(override)
    parts["[Content_Types].xml"] = _serialize_xml(content_types)

    for name in tuple(parts):
        if name.startswith("xl/externalLinks/"):
            del parts[name]
    return tuple(cached_values), tuple(removed_defined_names)


def _write_package(
    output_path: str | Path,
    infos: list[ZipInfo],
    parts: dict[str, bytes],
) -> None:
    info_by_name = {info.filename: info for info in infos}
    with ZipFile(output_path, "w", ZIP_DEFLATED) as package:
        for info in infos:
            payload = parts.get(info.filename)
            if payload is not None:
                package.writestr(info, payload)
        for name, payload in parts.items():
            if name not in info_by_name:
                package.writestr(name, payload)


def prepare_xlsx_for_ods(
    input_path: str | Path,
    output_path: str | Path,
    *,
    password: str | None,
    allow_protection_loss: bool,
) -> XlsxOdsPreparation:
    """Create the only backend-visible XLSX copy for selected POLICY-02 B/B."""

    infos, parts = _read_package(input_path)
    try:
        inspection = _inspect_parts(parts, source_filename=Path(input_path).name)
        password_nodes = _passworded_protection_nodes(parts)
    except (KeyError, etree.XMLSyntaxError, ValueError) as exc:
        if isinstance(exc, XlsxOdsPolicyError):
            raise
        raise XlsxOdsPolicyError("INVALID_XLSX_PACKAGE", "The XLSX package cannot be prepared safely.") from exc

    if password_nodes:
        if not password:
            raise XlsxOdsPolicyError(
                "PROTECTION_PASSWORD_REQUIRED",
                "This workbook has password-protected structure or sheets. Enter the password to continue.",
            )
        if not all(
            _protection_password_matches(node, password, workbook=is_workbook)
            for _, _, node, is_workbook in password_nodes
        ):
            raise XlsxOdsPolicyError(
                "PROTECTION_PASSWORD_INVALID",
                "The spreadsheet protection password is invalid.",
            )
        if not allow_protection_loss:
            raise XlsxOdsPolicyError(
                "PROTECTION_LOSS_CONSENT_REQUIRED",
                "ODS delivery requires removing password protection from the private conversion copy. Confirm this loss to continue.",
            )

    for part_name, root, node, _is_workbook in password_nodes:
        parent = node.getparent()
        if parent is not None:
            parent.remove(node)
            parts[part_name] = _serialize_xml(root)

    flattened_cached_values: tuple[str, ...] = ()
    removed_external_defined_names: tuple[str, ...] = ()
    if inspection.unsupported_external_references:
        raise XlsxOdsPolicyError(
            "EXTERNAL_LINK_FORMULA_UNSUPPORTED",
            "External workbook references exist outside cells or defined names and cannot be flattened safely.",
        )
    if inspection.external_formula_cells or inspection.external_link_parts_present or inspection.external_defined_names:
        flattened_cached_values, removed_external_defined_names = _flatten_external_links(
            parts,
            source_filename=Path(input_path).name,
        )

    _write_package(output_path, infos, parts)
    return XlsxOdsPreparation(
        output_path=str(output_path),
        external_links_flattened=bool(
            inspection.external_formula_cells
            or inspection.external_link_parts_present
            or inspection.external_defined_names
        ),
        protection_removed=bool(password_nodes),
        flattened_cached_values=flattened_cached_values,
        removed_protection_elements=inspection.password_protected_elements,
        removed_external_defined_names=removed_external_defined_names,
        fidelity_risk_counts=inspection.fidelity_risk_counts,
    )


def _numeric_value_matches(actual: str, expected: str) -> bool:
    try:
        return Decimal(actual) == Decimal(expected)
    except InvalidOperation:
        return actual == expected


def validate_prepared_ods(path: str | Path, preparation: XlsxOdsPreparation) -> None:
    """Reject a backend success whose real ODS contradicts selected POLICY-02."""

    expected_values = preparation.flattened_cached_values
    matched_values = [False] * len(expected_values)
    has_table_source = False
    has_external_formula = False
    has_ref_error = False
    retained_protection = False
    removed_elements = set(preparation.removed_protection_elements)
    try:
        with ZipFile(path) as package, package.open("content.xml") as stream:
            iterator = etree.iterparse(
                stream,
                events=("end",),
                resolve_entities=False,
                no_network=True,
                huge_tree=True,
            )
            for _event, element in iterator:
                if preparation.external_links_flattened:
                    if element.tag == f"{{{_ODF_TABLE_NS}}}table-source":
                        has_table_source = True
                    if element.tag == f"{{{_ODF_TABLE_NS}}}table-cell":
                        formula = str(element.get(f"{{{_ODF_TABLE_NS}}}formula", ""))
                        if "file:///" in formula or "#[" in formula:
                            has_external_formula = True
                        actual_value = element.get(f"{{{_ODF_OFFICE_NS}}}value")
                        if actual_value is not None:
                            for index, expected in enumerate(expected_values):
                                if not matched_values[index] and _numeric_value_matches(str(actual_value), expected):
                                    matched_values[index] = True
                    if any(
                        "#REF!" in str(value)
                        for value in (
                            *element.attrib.values(),
                            element.text or "",
                            element.tail or "",
                        )
                    ):
                        has_ref_error = True

                if preparation.protection_removed:
                    if (
                        element.tag == f"{{{_ODF_OFFICE_NS}}}spreadsheet"
                        and "workbook" in removed_elements
                        and _is_enabled(element.get(f"{{{_ODF_TABLE_NS}}}structure-protected"))
                    ):
                        retained_protection = True
                    if element.tag == f"{{{_ODF_TABLE_NS}}}table":
                        table_name = str(element.get(f"{{{_ODF_TABLE_NS}}}name", ""))
                        if table_name in removed_elements and _is_enabled(element.get(f"{{{_ODF_TABLE_NS}}}protected")):
                            retained_protection = True

                element.clear()
                parent = element.getparent()
                if parent is not None:
                    while element.getprevious() is not None:
                        del parent[0]
    except (BadZipFile, KeyError, OSError, etree.XMLSyntaxError) as exc:
        raise XlsxOdsPolicyError(
            "POLICY02_FINAL_ARTIFACT_INVALID",
            "The backend reported success but did not produce an inspectable ODS package.",
        ) from exc

    if preparation.external_links_flattened:
        missing_values = [
            expected for expected, matched in zip(expected_values, matched_values, strict=True) if not matched
        ]
        if has_table_source or has_external_formula or missing_values or (expected_values and has_ref_error):
            raise XlsxOdsPolicyError(
                "EXTERNAL_LINK_FLATTENING_NOT_DELIVERED",
                "The ODS artifact did not preserve the cached external value without a live link.",
            )

    if preparation.protection_removed and retained_protection:
        raise XlsxOdsPolicyError(
            "PROTECTION_REMOVAL_NOT_DELIVERED",
            "The ODS artifact still reports workbook or sheet protection.",
        )

"""Business-neutral DOCX caption, field, bookmark, and table-role helpers."""

from __future__ import annotations

import base64
import re
from collections.abc import Callable
from dataclasses import dataclass
from typing import Any

from docwen_core.docx_bookmarks import (
    BOOKMARK_NAME_MAX_LENGTH,
    DocxBookmarkInventory,
    prove_bookmark_name,
)
from docwen_core.models.semantic_document import PORTABLE_TARGET_ID_MAX_LENGTH

_BOOKMARK_PREFIX = "_DW_"
_SHORTHAND_BOOKMARK_PREFIX = "_DWF_S_"
_OBJECT_BOOKMARK_PREFIX = "_DWO_"
_PAIRING_CAPTION_BOOKMARK_PREFIX = "_DWP_C_"
_PAIRING_OBJECT_BOOKMARK_PREFIX = "_DWP_O_"
_PAIRING_TOKEN_RE = re.compile(r"^[0-9A-Z]+$")
BIBLIOGRAPHY_BOOKMARK_NAME = "_DWB_BIBLIOGRAPHY"
CITATION_BOOKMARK_PREFIX = "_DWC_"
BIBLIOGRAPHY_ENTRY_BOOKMARK_PREFIX = "_DWE_"
_CAPTION_KIND_BY_FIELD = {
    "figure": "figure",
    "table": "table",
    "equation": "equation",
    "listing": "listing",
}


@dataclass(frozen=True, slots=True)
class DocxSemanticCaption:
    """A caption proven by SEQ plus a target or internal object pairing."""

    kind: str
    target_id: str | None
    content: str
    bookmark_name: str
    source_form: str = "imported"
    label: str = ""
    cached_number: str = ""
    target_ids: tuple[str, ...] = ()
    target_marker_present: bool = False
    target_marker_malformed: bool = False
    shorthand_marker_present: bool = False
    shorthand_marker_malformed: bool = False
    pairing_tokens: tuple[str, ...] = ()
    pairing_marker_present: bool = False
    pairing_marker_malformed: bool = False


@dataclass(frozen=True, slots=True)
class DocxPairingMarkers:
    """Physical internal pairing markers found inside one DOCX object."""

    tokens: tuple[str, ...]
    present: bool
    malformed: bool


@dataclass(frozen=True, slots=True)
class DocxTargetMarkers:
    """Decoded target markers plus their package-wide physical proof state."""

    target_ids: tuple[str, ...]
    present: bool
    malformed: bool


@dataclass(frozen=True, slots=True)
class DocxSemanticTableMetadata:
    """Header/repeat metadata encoded with standard OOXML role elements."""

    header_rows: int
    header_columns: int
    repeat_header: str


def encode_target_bookmark(target_id: str) -> str:
    """Encode a portable target ID into a reversible OOXML bookmark name."""

    payload = _encode_portable_target(target_id)
    return f"{_BOOKMARK_PREFIX}{payload}"


def decode_target_bookmark(bookmark_name: str) -> str | None:
    """Decode a DocWen semantic bookmark, returning ``None`` for other names."""

    if not bookmark_name.startswith(_BOOKMARK_PREFIX):
        return None
    payload = bookmark_name[len(_BOOKMARK_PREFIX) :]
    if not payload:
        return None
    padded = payload + "=" * ((8 - len(payload) % 8) % 8)
    try:
        decoded = base64.b32decode(padded, casefold=True).decode("ascii")
    except (UnicodeDecodeError, ValueError):
        return None
    if not _is_portable_target_id(decoded):
        return None
    return decoded


def encode_shorthand_bookmark(target_id: str) -> str:
    """Encode caption source-form provenance in a standard bookmark name."""

    payload = _encode_portable_target(target_id)
    return f"{_SHORTHAND_BOOKMARK_PREFIX}{payload}"


def encode_object_bookmark(target_id: str) -> str:
    """Encode the target bound to a physical DOCX object such as a table."""

    payload = _encode_portable_target(target_id)
    return f"{_OBJECT_BOOKMARK_PREFIX}{payload}"


def decode_object_bookmark(bookmark_name: str) -> str | None:
    """Decode a DocWen object-binding bookmark."""

    if not bookmark_name.startswith(_OBJECT_BOOKMARK_PREFIX):
        return None
    payload = bookmark_name[len(_OBJECT_BOOKMARK_PREFIX) :]
    if not payload:
        return None
    padded = payload + "=" * ((8 - len(payload) % 8) % 8)
    try:
        decoded = base64.b32decode(padded, casefold=True).decode("ascii")
    except (UnicodeDecodeError, ValueError):
        return None
    if not _is_portable_target_id(decoded):
        return None
    return decoded


def encode_caption_pairing_bookmark(token: str) -> str:
    """Encode an internal caption-side object pairing marker."""

    return _encode_pairing_bookmark(_PAIRING_CAPTION_BOOKMARK_PREFIX, token)


def encode_object_pairing_bookmark(token: str) -> str:
    """Encode an internal object-side caption pairing marker."""

    return _encode_pairing_bookmark(_PAIRING_OBJECT_BOOKMARK_PREFIX, token)


def encode_citation_bookmark(cluster_id: str) -> str:
    """Encode a neutral citation cluster ID as a reversible bookmark name."""

    return f"{CITATION_BOOKMARK_PREFIX}{_encode_portable_target(cluster_id)}"


def decode_citation_bookmark(bookmark_name: str) -> str | None:
    """Decode a canonical citation-cluster bookmark name."""

    return _decode_portable_semantic_bookmark(
        bookmark_name,
        prefix=CITATION_BOOKMARK_PREFIX,
    )


def encode_bibliography_entry_bookmark(item_id: str) -> str:
    """Encode a neutral bibliography item ID as a reversible bookmark name."""

    return f"{BIBLIOGRAPHY_ENTRY_BOOKMARK_PREFIX}{_encode_portable_target(item_id)}"


def decode_bibliography_entry_bookmark(bookmark_name: str) -> str | None:
    """Decode a canonical bibliography-entry bookmark name."""

    return _decode_portable_semantic_bookmark(
        bookmark_name,
        prefix=BIBLIOGRAPHY_ENTRY_BOOKMARK_PREFIX,
    )


def inspect_caption_pairing_markers(
    element: Any,
    *,
    bookmark_inventory: DocxBookmarkInventory,
) -> DocxPairingMarkers:
    """Inspect balanced internal caption-side pairing markers."""

    return _inspect_pairing_markers(
        element,
        prefix=_PAIRING_CAPTION_BOOKMARK_PREFIX,
        bookmark_inventory=bookmark_inventory,
    )


def inspect_object_pairing_markers(
    element: Any,
    *,
    bookmark_inventory: DocxBookmarkInventory,
) -> DocxPairingMarkers:
    """Inspect balanced internal object-side pairing markers."""

    return _inspect_pairing_markers(
        element,
        prefix=_PAIRING_OBJECT_BOOKMARK_PREFIX,
        bookmark_inventory=bookmark_inventory,
    )


def inspect_caption_target_markers(
    element: Any,
    *,
    bookmark_inventory: DocxBookmarkInventory,
) -> DocxTargetMarkers:
    """Inspect target bookmarks that prove a caption address."""

    return _inspect_target_markers(
        element,
        prefix=_BOOKMARK_PREFIX,
        decoder=decode_target_bookmark,
        bookmark_inventory=bookmark_inventory,
    )


def inspect_object_target_markers(
    element: Any,
    *,
    bookmark_inventory: DocxBookmarkInventory,
) -> DocxTargetMarkers:
    """Inspect target bookmarks that bind a physical object."""

    return _inspect_target_markers(
        element,
        prefix=_OBJECT_BOOKMARK_PREFIX,
        decoder=decode_object_bookmark,
        bookmark_inventory=bookmark_inventory,
    )


def inspect_shorthand_target_markers(
    element: Any,
    *,
    bookmark_inventory: DocxBookmarkInventory,
) -> DocxTargetMarkers:
    """Inspect optional caption source-form provenance bookmarks."""

    return _inspect_target_markers(
        element,
        prefix=_SHORTHAND_BOOKMARK_PREFIX,
        decoder=_decode_shorthand_bookmark,
        bookmark_inventory=bookmark_inventory,
    )


def extract_object_target(
    element: Any,
    *,
    bookmark_inventory: DocxBookmarkInventory,
) -> str | None:
    """Return the explicit target bound inside a DOCX object element."""

    markers = inspect_object_target_markers(
        element,
        bookmark_inventory=bookmark_inventory,
    )
    if markers.malformed or len(markers.target_ids) != 1:
        return None
    return markers.target_ids[0]


def _decode_shorthand_bookmark(bookmark_name: str) -> str | None:
    if not bookmark_name.startswith(_SHORTHAND_BOOKMARK_PREFIX):
        return None
    payload = bookmark_name[len(_SHORTHAND_BOOKMARK_PREFIX) :]
    padded = payload + "=" * ((8 - len(payload) % 8) % 8)
    try:
        decoded = base64.b32decode(padded, casefold=True).decode("ascii")
    except (UnicodeDecodeError, ValueError):
        return None
    return decoded if _is_portable_target_id(decoded) else None


def extract_semantic_caption(
    paragraph_element: Any,
    *,
    bookmark_inventory: DocxBookmarkInventory,
) -> DocxSemanticCaption | None:
    """Extract the legacy target-bound caption used by Markdown conversion."""

    extracted = _extract_semantic_caption(
        paragraph_element,
        allow_internal_pairing=False,
        bookmark_inventory=bookmark_inventory,
    )
    if (
        extracted is None
        or extracted.target_id is None
        or extracted.target_marker_malformed
        or extracted.shorthand_marker_malformed
    ):
        return None
    return extracted


def extract_neutral_semantic_caption(
    paragraph_element: Any,
    *,
    bookmark_inventory: DocxBookmarkInventory,
    proven_object_kind: str | None = None,
) -> DocxSemanticCaption | None:
    """Extract a neutral caption proven by a target or internal object pairing.

    ``proven_object_kind`` is a caller-supplied fallback for editors that
    localize the identifier of a ``SEQ`` field while preserving DocWen's
    caption and physical-object bookmarks.  Callers must derive the hint from
    an independently proven object binding.  Known identifiers always win;
    the hint applies only when exactly one otherwise-unknown ``SEQ`` remains.
    """

    return _extract_semantic_caption(
        paragraph_element,
        allow_internal_pairing=True,
        bookmark_inventory=bookmark_inventory,
        proven_object_kind=proven_object_kind,
    )


def _extract_semantic_caption(
    paragraph_element: Any,
    *,
    allow_internal_pairing: bool,
    bookmark_inventory: DocxBookmarkInventory,
    proven_object_kind: str | None = None,
) -> DocxSemanticCaption | None:
    """Extract a caption without relying on adjacency to its owning object."""

    from docx.oxml.ns import qn

    target_markers = inspect_caption_target_markers(
        paragraph_element,
        bookmark_inventory=bookmark_inventory,
    )
    shorthand_markers = inspect_shorthand_target_markers(
        paragraph_element,
        bookmark_inventory=bookmark_inventory,
    )
    pairing = inspect_caption_pairing_markers(
        paragraph_element,
        bookmark_inventory=bookmark_inventory,
    )
    target_id = target_markers.target_ids[0] if len(target_markers.target_ids) == 1 else None
    shorthand_target = shorthand_markers.target_ids[0] if len(shorthand_markers.target_ids) == 1 else None
    if (
        not target_markers.present
        and not shorthand_markers.present
        and (not allow_internal_pairing or not pairing.present)
    ):
        return None

    fields = _field_instruction_results(paragraph_element)
    known_candidates: list[tuple[str, str]] = []
    unknown_candidates: list[str] = []
    for instruction, cached_result in fields:
        match = re.match(r"\s*SEQ\s+([^\\\s]+)", instruction, re.IGNORECASE)
        if match is None:
            continue
        known_kind = _CAPTION_KIND_BY_FIELD.get(match.group(1).casefold())
        if known_kind is None:
            unknown_candidates.append(cached_result.strip())
        else:
            known_candidates.append((known_kind, cached_result.strip()))

    if known_candidates:
        kind, cached_number = known_candidates[0]
    elif proven_object_kind is not None and len(unknown_candidates) == 1:
        kind = proven_object_kind
        cached_number = unknown_candidates[0]
    else:
        kind = None
        cached_number = ""
    if kind is None:
        return None

    visible = "".join(element.text or "" for element in paragraph_element.iter(qn("w:t"))).strip()
    prefix, separator, content = visible.partition(":")
    if not separator:
        return None
    label = prefix.strip()
    if cached_number and label.endswith(cached_number):
        label = label[: -len(cached_number)].strip()
    if not label:
        label = kind.title()
    return DocxSemanticCaption(
        kind=kind,
        target_id=target_id,
        content=content.strip(),
        bookmark_name=(encode_target_bookmark(target_id) if target_id is not None else ""),
        source_form=("shorthand" if target_id is not None and shorthand_target == target_id else "imported"),
        label=label,
        cached_number=cached_number,
        target_ids=target_markers.target_ids,
        target_marker_present=target_markers.present,
        target_marker_malformed=target_markers.malformed,
        shorthand_marker_present=shorthand_markers.present,
        shorthand_marker_malformed=(
            shorthand_markers.malformed
            or (shorthand_markers.present and (len(shorthand_markers.target_ids) != 1 or shorthand_target != target_id))
        ),
        pairing_tokens=pairing.tokens,
        pairing_marker_present=pairing.present,
        pairing_marker_malformed=pairing.malformed,
    )


def _encode_portable_target(target_id: str) -> str:
    if not _is_portable_target_id(target_id):
        raise ValueError(
            f"target_id must be an ASCII portable identifier of at most {PORTABLE_TARGET_ID_MAX_LENGTH} characters"
        )
    return base64.b32encode(target_id.encode("ascii")).decode("ascii").rstrip("=")


def _decode_portable_semantic_bookmark(bookmark_name: str, *, prefix: str) -> str | None:
    if not bookmark_name.startswith(prefix):
        return None
    payload = bookmark_name[len(prefix) :]
    if not payload:
        return None
    padded = payload + "=" * ((8 - len(payload) % 8) % 8)
    try:
        decoded = base64.b32decode(padded, casefold=True).decode("ascii")
    except (UnicodeDecodeError, ValueError):
        return None
    if not _is_portable_target_id(decoded):
        return None
    canonical = f"{prefix}{base64.b32encode(decoded.encode('ascii')).decode('ascii').rstrip('=')}"
    return decoded if canonical.casefold() == bookmark_name.casefold() else None


def _is_portable_target_id(target_id: str) -> bool:
    return (
        len(target_id) <= PORTABLE_TARGET_ID_MAX_LENGTH
        and re.fullmatch(r"[A-Za-z][A-Za-z0-9._-]*", target_id) is not None
    )


def _encode_pairing_bookmark(prefix: str, token: str) -> str:
    if _PAIRING_TOKEN_RE.fullmatch(token) is None:
        raise ValueError("pairing token must contain only uppercase ASCII letters and digits")
    bookmark_name = f"{prefix}{token}"
    if len(bookmark_name) > BOOKMARK_NAME_MAX_LENGTH:
        raise ValueError("pairing bookmark name exceeds Word's 40-character limit")
    return bookmark_name


def _inspect_pairing_markers(
    element: Any,
    *,
    prefix: str,
    bookmark_inventory: DocxBookmarkInventory,
) -> DocxPairingMarkers:
    values, present, malformed = _inspect_decoded_bookmarks(
        element,
        prefix=prefix,
        decoder=lambda name: _decode_pairing_bookmark(name, prefix=prefix),
        bookmark_inventory=bookmark_inventory,
    )
    return DocxPairingMarkers(tokens=values, present=present, malformed=malformed)


def _inspect_target_markers(
    element: Any,
    *,
    prefix: str,
    decoder: Callable[[str], str | None],
    bookmark_inventory: DocxBookmarkInventory,
) -> DocxTargetMarkers:
    values, present, malformed = _inspect_decoded_bookmarks(
        element,
        prefix=prefix,
        decoder=decoder,
        bookmark_inventory=bookmark_inventory,
    )
    return DocxTargetMarkers(target_ids=values, present=present, malformed=malformed)


def _inspect_decoded_bookmarks(
    element: Any,
    *,
    prefix: str,
    decoder: Callable[[str], str | None],
    bookmark_inventory: DocxBookmarkInventory,
) -> tuple[tuple[str, ...], bool, bool]:
    from docx.oxml.ns import qn

    marker_names = [
        item.get(qn("w:name")) or ""
        for item in element.iter(qn("w:bookmarkStart"))
        if (item.get(qn("w:name")) or "").startswith(prefix)
    ]
    values: list[str] = []
    malformed = False
    for bookmark_name in marker_names:
        decoded = decoder(bookmark_name)
        if decoded is None:
            malformed = True
            continue
        values.append(decoded)
        if not prove_bookmark_name(
            bookmark_inventory,
            bookmark_name,
            scope_element=element,
        ).valid:
            malformed = True
    return tuple(values), bool(marker_names), malformed


def _decode_pairing_bookmark(bookmark_name: str, *, prefix: str) -> str | None:
    if not bookmark_name.startswith(prefix):
        return None
    token = bookmark_name[len(prefix) :]
    if _PAIRING_TOKEN_RE.fullmatch(token) is None or len(bookmark_name) > BOOKMARK_NAME_MAX_LENGTH:
        return None
    return token


def render_semantic_reference_text(
    paragraph_element: Any,
    *,
    bookmark_inventory: DocxBookmarkInventory,
) -> str | None:
    """Replace supported REF fields with canonical ``@target-id`` tokens."""

    from docx.oxml.ns import qn

    output: list[str] = []
    field_instruction: list[str] = []
    field_result: list[str] = []
    in_field = False
    in_result = False
    found_reference = False

    for element in paragraph_element.iter():
        if element.tag == qn("w:fldSimple"):
            instruction = element.get(qn("w:instr")) or ""
            replacement = _reference_token_from_instruction(
                instruction,
                bookmark_inventory=bookmark_inventory,
            )
            if replacement is not None:
                output.append(replacement)
                found_reference = True
            else:
                output.extend(text.text or "" for text in element.iter(qn("w:t")))
            continue
        if element.tag == qn("w:fldChar"):
            field_type = element.get(qn("w:fldCharType")) or ""
            if field_type == "begin":
                in_field = True
                in_result = False
                field_instruction = []
                field_result = []
            elif in_field and field_type == "separate":
                in_result = True
            elif in_field and field_type == "end":
                instruction = "".join(field_instruction)
                replacement = _reference_token_from_instruction(
                    instruction,
                    bookmark_inventory=bookmark_inventory,
                )
                if replacement is not None:
                    output.append(replacement)
                    found_reference = True
                else:
                    output.append("".join(field_result))
                in_field = False
                in_result = False
            continue
        if element.tag == qn("w:instrText") and in_field:
            field_instruction.append(element.text or "")
            continue
        if element.tag == qn("w:t"):
            if in_field:
                if in_result:
                    field_result.append(element.text or "")
            elif not _has_ancestor_tag(element, qn("w:fldSimple")):
                output.append(element.text or "")

    return "".join(output) if found_reference else None


def extract_semantic_table_metadata(
    tbl_element: Any,
    *,
    default_first_row: bool = True,
) -> DocxSemanticTableMetadata:
    """Read the standard ``cnfStyle``/``tblHeader`` metadata written by DocWen."""

    from docx.oxml.ns import qn

    rows = tbl_element.findall(qn("w:tr"))
    header_rows = 0
    repeat_values: list[bool | None] = []
    for row in rows:
        tr_pr = row.find(qn("w:trPr"))
        cnf = tr_pr.find(qn("w:cnfStyle")) if tr_pr is not None else None
        is_header = cnf is not None and (cnf.get(qn("w:firstRow")) or "") in {"1", "true"}
        if not is_header:
            break
        header_rows += 1
        tbl_header = tr_pr.find(qn("w:tblHeader")) if tr_pr is not None else None
        repeat_values.append(None if tbl_header is None else (tbl_header.get(qn("w:val")) or "1") not in {"0", "false"})

    header_columns = 0
    if rows:
        for cell in rows[0].findall(qn("w:tc")):
            tc_pr = cell.find(qn("w:tcPr"))
            cnf = tc_pr.find(qn("w:cnfStyle")) if tc_pr is not None else None
            is_header = cnf is not None and (cnf.get(qn("w:firstColumn")) or "") in {"1", "true"}
            if not is_header:
                break
            grid_span = tc_pr.find(qn("w:gridSpan")) if tc_pr is not None else None
            header_columns += int(grid_span.get(qn("w:val")) or 1) if grid_span is not None else 1

    if repeat_values and all(value is True for value in repeat_values):
        repeat_header = "always"
    elif repeat_values and all(value is False for value in repeat_values):
        repeat_header = "never"
    else:
        repeat_header = "inherit"
    if header_rows == 0 and default_first_row:
        header_rows = 1 if rows else 0
    return DocxSemanticTableMetadata(
        header_rows=header_rows,
        header_columns=header_columns,
        repeat_header=repeat_header,
    )


def _field_instructions(element: Any) -> list[str]:
    from docx.oxml.ns import qn

    instructions = [item.text or "" for item in element.iter(qn("w:instrText"))]
    instructions.extend(item.get(qn("w:instr")) or "" for item in element.iter(qn("w:fldSimple")))
    return instructions


def _field_instruction_results(element: Any) -> list[tuple[str, str]]:
    from docx.oxml.ns import qn

    fields: list[tuple[str, str]] = []
    instruction_parts: list[str] = []
    result_parts: list[str] = []
    in_field = False
    in_result = False
    for item in element.iter():
        if item.tag == qn("w:fldSimple"):
            instruction = item.get(qn("w:instr")) or ""
            result = "".join(text.text or "" for text in item.iter(qn("w:t")))
            fields.append((instruction, result))
            continue
        if item.tag == qn("w:fldChar"):
            field_type = item.get(qn("w:fldCharType")) or ""
            if field_type == "begin":
                in_field = True
                in_result = False
                instruction_parts = []
                result_parts = []
            elif in_field and field_type == "separate":
                in_result = True
            elif in_field and field_type == "end":
                fields.append(("".join(instruction_parts), "".join(result_parts)))
                in_field = False
                in_result = False
            continue
        if item.tag == qn("w:instrText") and in_field:
            instruction_parts.append(item.text or "")
        elif item.tag == qn("w:t") and in_field and in_result:
            result_parts.append(item.text or "")
    return fields


def _reference_token_from_instruction(
    instruction: str,
    *,
    bookmark_inventory: DocxBookmarkInventory,
) -> str | None:
    match = re.match(r"\s*REF\s+(\S+)", instruction, re.IGNORECASE)
    if match is None:
        return None
    bookmark_name = match.group(1)
    target_id = decode_target_bookmark(bookmark_name)
    if target_id is not None and not prove_bookmark_name(bookmark_inventory, bookmark_name).valid:
        return None
    return f"@{target_id}" if target_id is not None else None


def _has_ancestor_tag(element: Any, tag: str) -> bool:
    parent = element.getparent()
    while parent is not None:
        if parent.tag == tag:
            return True
        parent = parent.getparent()
    return False

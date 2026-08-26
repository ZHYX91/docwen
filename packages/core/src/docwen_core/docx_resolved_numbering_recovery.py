"""Proof-only recovery for DOCX packages carrying resolved-v4 signals.

This reader deliberately does not reconstruct a provider document or claim an
exact Markdown round-trip.  The current physical package carries enough
authority to prove Word numbering, caption, reference, Citation, and disabled
occurrence facts, but it does not carry the complete admitted neutral input.
Consequently, callers may use the recovered authored tokens and may preserve
visible authored text, while a separate stable diagnostic records that exact
source bytes are unavailable.
"""

from __future__ import annotations

import hashlib
import re
from pathlib import Path
from typing import Any, cast
from zipfile import ZipFile

from docwen_core import _docx_semantics_v3_fenced as fenced
from docwen_core._docx_recovery_map import (
    RESOLVED_V4_RECOVERY_MAP_NAMESPACE,
    ResolvedV4RecoveryMap,
    compute_physical_projection,
    read_recovery_map,
)
from docwen_core._docx_semantics_v3_fenced_map import parse_fenced_source_map
from docwen_core._docx_semantics_v3_model import (
    ANCHOR_TOPOLOGY_MAP_NAMESPACE,
    CAPTION_STYLE_BINDING_MAP_NAMESPACE,
    REFERENCE_OCCURRENCE_MAP_NAMESPACE,
    SOFT_REFERENCE_MAP_NAMESPACE,
    TARGET_MAP_NAMESPACE,
    DocxSemanticsV3Error,
    RecoveredCaptionV3,
)
from docwen_core._docx_semantics_v3_ooxml import (
    recover_paragraph_children,
    sdt_tag,
    visible_text,
)
from docwen_core._docx_semantics_v3_package import (
    parse_reference_occurrence_map,
    parse_semantic_map,
    parse_soft_reference_map,
    read_owned_map_parts,
    verify_custom_xml_support,
)
from docwen_core._docx_semantics_v3_recovery import (
    DocxSemanticsV3Recovery,
    _parse_v3_caption,
    _prove_caption_object_kind,
)
from docwen_core._docx_semantics_v3_styles import (
    caption_kind_for_paragraph_style,
    parse_caption_style_binding_map,
    prove_caption_style_registry,
)
from docwen_core._docx_semantics_v3_topology import (
    logical_group_elements,
    parse_anchor_topology_map,
)
from docwen_core.docx_citation_ooxml import (
    CITATION_ITEM_MAP_NAMESPACE,
    CITATION_OCCURRENCE_MAP_NAMESPACE,
    parse_citation_occurrence_map,
    read_proven_resolved_citations,
)
from docwen_core.docx_numbering_import import (
    HeadingNumberingProofIndex,
    NumberingImportDiagnostic,
    ProofOnlyHeadingImport,
    import_heading_without_source_mutation,
)
from docwen_core.docx_numbering_occurrence import (
    NUMBERING_OCCURRENCE_MAP_NAMESPACE,
    NumberingOccurrenceIdentity,
    parse_numbering_occurrence_map,
    prove_numbering_occurrence_sdt,
)

RESOLVED_V4_SOURCE_SNAPSHOT_MISSING_DIAGNOSTIC = "docwen.docx.resolved_v4.source_snapshot_missing"
RESOLVED_V4_SOURCE_SNAPSHOT_MISSING_MESSAGE = (
    "Resolved-v4 Word semantics were proven, but this package has no authenticated complete neutral/source "
    "snapshot; output is generic extraction and is not an exact Markdown round-trip."
)

_CAPTION_COUNTER = {
    "figure": "Figure",
    "table": "Table",
    "equation": "Equation",
    "code_block": "Code",
}
_CAPTION_STYLE_KEY = {
    "figure": "figure_caption",
    "table": "table_caption",
    "equation": "equation_caption",
    "code_block": "code_block_caption",
}
_FIELD_FORMAT = r"(?:ARABIC|ALPHABETIC|alphabetic|ROMAN|roman)"
_SUSPECTED_PREFIX_RE = re.compile(
    r"^(?:[0-9]+(?:\.[0-9]+)+[.)、．]?\s+|第[^\s]{1,16}[章节篇部]\s*|[一二三四五六七八九十百千]+[、.．]\s*)"
)


class ResolvedNumberingV4Recovery(DocxSemanticsV3Recovery):
    """Authenticated adapter over the currently frozen v4 physical carriers."""

    is_resolved_v4 = True
    has_complete_source_snapshot = False
    source_recovery_available = False
    _resolved_v4_inline_tokens: dict[str, str]
    _resolved_v4_citations: tuple[Any, ...]
    _resolved_v4_heading_proof: HeadingNumberingProofIndex
    _resolved_v4_heading_imports: dict[int, ProofOnlyHeadingImport]
    _recovery_map: ResolvedV4RecoveryMap | None = None
    _recovery_item_number: int | None = None

    @classmethod
    def load_if_present(
        cls,
        path: str | Path,
        document: Any,
    ) -> ResolvedNumberingV4Recovery | None:
        """Return a strict v4 reader only for an explicit current-package signal.

        Packages that are physically indistinguishable from historical v3 are
        left to the frozen v3 reader.  A future authenticated recovery map must
        remove that ambiguity before exact round-trip can be claimed.
        """

        package_path = Path(path)
        with ZipFile(package_path) as package:
            owned = read_owned_map_parts(package)
            for namespace, (item_number, _root) in owned.items():
                verify_custom_xml_support(package, item_number, namespace)
            targets = []
            if TARGET_MAP_NAMESPACE in owned:
                _number, root = owned[TARGET_MAP_NAMESPACE]
                targets, _anchors = parse_semantic_map(root)
            caption_styles = ()
            if CAPTION_STYLE_BINDING_MAP_NAMESPACE in owned:
                _number, root = owned[CAPTION_STYLE_BINDING_MAP_NAMESPACE]
                caption_styles = parse_caption_style_binding_map(root)
                prove_caption_style_registry(package, caption_styles)
            recovery_map = read_recovery_map(package)
            if recovery_map is not None:
                item_number, _parsed = recovery_map
                verify_custom_xml_support(package, item_number, RESOLVED_V4_RECOVERY_MAP_NAMESPACE)
        if not _has_explicit_resolved_v4_signal(document, owned, targets, caption_styles):
            return None
        recovery = cls._load_proven(package_path, document, owned)
        if recovery_map is not None:
            recovery._bind_recovery_map(package_path, recovery_map[1], recovery_map[0])
        return recovery

    def _bind_recovery_map(
        self,
        package_path: Path,
        value: ResolvedV4RecoveryMap,
        item_number: int,
    ) -> None:
        """Bind the exact-neutral recovery authority without raw file admission.

        The map is authentic when its content digest and the whole-package
        physical projection both match the reopened package.  The three raw
        pointer files live in the request-owned staging tree; exact-neutral
        recovery consumes them only through the authenticated adapter below,
        and generic proof-only recovery never touches them.
        """

        physical = compute_physical_projection(package_path, exclude_item_numbers={item_number})
        if physical != value.physical_sha256:
            raise DocxSemanticsV3Error("resolved-v4 recovery projection digest differs from the package")
        if not re.fullmatch(r"[0-9a-f]{64}", value.source_sha256) or not re.fullmatch(
            r"[0-9a-f]{64}", value.plan_sha256
        ):
            raise DocxSemanticsV3Error("resolved-v4 recovery map identity digests are invalid")
        self._recovery_map = value
        self._recovery_item_number = item_number
        self.has_complete_source_snapshot = True
        self.source_recovery_available = True

    def prove_exact_recovery_raw(
        self,
        *,
        neutral_raw: bytes,
        plan_raw: bytes,
        authored_source: bytes,
    ) -> None:
        """Prove three caller-supplied raw files against the authenticated map."""

        if self._recovery_map is None:
            raise DocxSemanticsV3Error("resolved-v4 recovery map is not bound")
        expected = {item.role: item for item in self._recovery_map.pointers}
        actual = {
            "neutral_raw": neutral_raw,
            "plan_raw": plan_raw,
            "authored_source": authored_source,
        }
        for role, raw in actual.items():
            pointer = expected.get(role)
            if pointer is None:
                raise DocxSemanticsV3Error(f"resolved-v4 recovery map lacks {role} pointer")
            if len(raw) != pointer.bytes:
                raise DocxSemanticsV3Error(f"resolved-v4 recovery {role} byte count differs from its pointer")
            if hashlib.sha256(raw).hexdigest() != pointer.sha256:
                raise DocxSemanticsV3Error(f"resolved-v4 recovery {role} digest differs from its pointer")

    def recover_exact_neutral(self, *, raw_root: Path) -> bytes:
        """Return the authenticated neutral raw bytes from a request-owned tree.

        ``raw_root`` is the staging directory holding the three pointer files.
        Every pointer is a safe relative path inside that tree; no external or
        absolute path is ever admitted.
        """

        if self._recovery_map is None:
            raise DocxSemanticsV3Error("resolved-v4 recovery map is not bound")
        resolved: dict[str, bytes] = {}
        for pointer in self._recovery_map.pointers:
            relative = Path(pointer.relative_path)
            if relative.is_absolute() or any(part in {".", "..", ""} for part in relative.parts):
                raise DocxSemanticsV3Error("resolved-v4 recovery pointer is not a safe relative path")
            candidate = (raw_root / relative).resolve()
            root_resolved = raw_root.resolve()
            if root_resolved not in candidate.parents and candidate != root_resolved:
                raise DocxSemanticsV3Error("resolved-v4 recovery pointer escapes the raw root")
            if not candidate.is_file() or candidate.is_symlink():
                raise DocxSemanticsV3Error("resolved-v4 recovery raw file is missing or is a link")
            raw = candidate.read_bytes()
            if len(raw) != pointer.bytes:
                raise DocxSemanticsV3Error(f"resolved-v4 recovery {pointer.role} byte count differs from its pointer")
            if hashlib.sha256(raw).hexdigest() != pointer.sha256:
                raise DocxSemanticsV3Error(f"resolved-v4 recovery {pointer.role} digest differs from its pointer")
            resolved[pointer.role] = raw
        neutral = resolved.get("neutral_raw")
        if neutral is None:
            raise DocxSemanticsV3Error("resolved-v4 recovery neutral raw bytes are unavailable")
        return neutral

    @classmethod
    def _load_proven(
        cls,
        path: Path,
        document: Any,
        owned: dict[str, tuple[int, Any]],
    ) -> ResolvedNumberingV4Recovery:
        targets: list[Any] = []
        anchors: list[Any] = []
        soft: list[Any] = []
        references: list[Any] = []
        fenced_sources: list[Any] = []
        topology: list[Any] = []
        caption_styles: tuple[Any, ...] = ()
        occurrences: list[NumberingOccurrenceIdentity] = []

        if TARGET_MAP_NAMESPACE in owned:
            _number, root = owned[TARGET_MAP_NAMESPACE]
            targets, anchors = parse_semantic_map(root)
        if ANCHOR_TOPOLOGY_MAP_NAMESPACE in owned:
            _number, root = owned[ANCHOR_TOPOLOGY_MAP_NAMESPACE]
            topology = parse_anchor_topology_map(root)
        if SOFT_REFERENCE_MAP_NAMESPACE in owned:
            _number, root = owned[SOFT_REFERENCE_MAP_NAMESPACE]
            soft = parse_soft_reference_map(root)
        if REFERENCE_OCCURRENCE_MAP_NAMESPACE in owned:
            _number, root = owned[REFERENCE_OCCURRENCE_MAP_NAMESPACE]
            references = parse_reference_occurrence_map(root)
        if CAPTION_STYLE_BINDING_MAP_NAMESPACE in owned:
            _number, root = owned[CAPTION_STYLE_BINDING_MAP_NAMESPACE]
            caption_styles = parse_caption_style_binding_map(root)
        if fenced.FENCED_SOURCE_MAP_NAMESPACE in owned:
            _number, root = owned[fenced.FENCED_SOURCE_MAP_NAMESPACE]
            fenced_sources = parse_fenced_source_map(root)
        if NUMBERING_OCCURRENCE_MAP_NAMESPACE in owned:
            _number, root = owned[NUMBERING_OCCURRENCE_MAP_NAMESPACE]
            occurrences = parse_numbering_occurrence_map(root)

        citations = read_proven_resolved_citations(path)
        citation_records: tuple[Any, ...] = ()
        citation_tokens: dict[str, str] = {}
        citation_ranges: list[tuple[int, int, str]] = []
        if CITATION_OCCURRENCE_MAP_NAMESPACE in owned:
            _number, root = owned[CITATION_OCCURRENCE_MAP_NAMESPACE]
            citation_records = parse_citation_occurrence_map(root).occurrences
            citation_tokens = {item.tag: item.authored_token for item in citation_records}
            citation_ranges = [(item.source_start, item.source_end, item.tag) for item in citation_records]
        elif CITATION_ITEM_MAP_NAMESPACE in owned:
            raise DocxSemanticsV3Error("resolved-v4 Citation item map has no occurrence map")
        _prove_one_source_identity((*soft, *references, *fenced_sources, *occurrences, *citation_records))
        _prove_merged_inline_order(document, (*soft, *references, *citation_records))

        inline_tokens = {
            **{item.tag: item.authored_token for item in soft},
            **{item.tag: item.authored_token for item in references},
            **citation_tokens,
        }

        def parse_caption(paragraph: Any, kind: str, *, required_bookmark: str | None) -> tuple[str, str]:
            return _parse_resolved_v4_caption(
                paragraph,
                kind,
                required_bookmark=required_bookmark,
                inline_tokens=inline_tokens,
            )

        recovery = cast(
            ResolvedNumberingV4Recovery,
            cls._bind_document_evidence(
                document,
                targets,
                anchors,
                soft,
                references,
                caption_styles,
                fenced_sources,
                topology,
                topology_map_present=ANCHOR_TOPOLOGY_MAP_NAMESPACE in owned,
                caption_parser=parse_caption,
            ),
        )
        recovery._resolved_v4_inline_tokens = inline_tokens
        recovery._resolved_v4_citations = citations
        recovery._resolved_v4_heading_proof = HeadingNumberingProofIndex.load(path)
        recovery._resolved_v4_heading_imports = {}
        recovery._bind_disabled_occurrences(
            document,
            occurrences,
            caption_styles,
            soft,
            references,
            citation_ranges,
            parse_caption,
        )
        recovery._prove_target_headings(document, targets)
        recovery.caption_signatures = tuple(
            (item.kind, item.source_id, item.title, item.cached_number) for item in recovery.recovered_captions
        )
        return recovery

    @property
    def source_recovery_diagnostic(self) -> NumberingImportDiagnostic:
        return NumberingImportDiagnostic(
            RESOLVED_V4_SOURCE_SNAPSHOT_MISSING_DIAGNOSTIC,
            RESOLVED_V4_SOURCE_SNAPSHOT_MISSING_MESSAGE,
        )

    def heading_import(self, paragraph: Any) -> ProofOnlyHeadingImport:
        key = id(paragraph._p)
        existing = self._resolved_v4_heading_imports.get(key)
        if existing is not None:
            return existing
        imported = import_heading_without_source_mutation(
            paragraph,
            self._resolved_v4_heading_proof,
            suspected_visible_prefix=_SUSPECTED_PREFIX_RE.match(paragraph.text) is not None,
        )
        self._resolved_v4_heading_imports[key] = imported
        return imported

    def render_paragraph_text(self, paragraph_element: Any) -> str | None:
        text, semantic = recover_paragraph_children(
            list(paragraph_element),
            target_ids_by_bookmark=self._target_ids_by_bookmark,
            soft_tokens_by_tag={**self._soft_tokens_by_tag, **self._resolved_v4_inline_tokens},
            occurrence_tokens_by_tag=self._occurrence_tokens_by_tag,
        )
        return text if semantic else None

    def _prove_target_headings(self, document: Any, targets: list[Any]) -> None:
        from docx.text.paragraph import Paragraph

        headings_by_id = {
            anchor.source_id: element
            for element, anchor in self._block_anchors.items()
            if anchor.owner_kind == "semantic_target" and anchor.block_kind == "heading"
        }
        expected = {item.source_id for item in targets if item.kind == "heading"}
        if set(headings_by_id) != expected:
            raise DocxSemanticsV3Error("resolved-v4 Heading target inventory differs from its physical paragraphs")
        for source_id in sorted(expected):
            self.heading_import(Paragraph(headings_by_id[source_id], document))

    def _bind_disabled_occurrences(
        self,
        document: Any,
        occurrences: list[NumberingOccurrenceIdentity],
        caption_styles: tuple[Any, ...],
        soft_references: list[Any],
        references: list[Any],
        citation_ranges: list[tuple[int, int, str]],
        parse_caption: Any,
    ) -> None:
        from docx.oxml.ns import qn

        body = document.element.body
        physical = [
            item
            for item in body
            if item.tag == qn("w:sdt") and (sdt_tag(item) or "").startswith("docwen-numbering-occurrence-v1:")
        ]
        nested = [
            item
            for item in body.iter(qn("w:sdt"))
            if (sdt_tag(item) or "").startswith("docwen-numbering-occurrence-v1:")
        ]
        if physical != nested or [sdt_tag(item) for item in physical] != [item.tag for item in occurrences]:
            raise DocxSemanticsV3Error("resolved-v4 occurrence physical order/cardinality differs from its map")
        if not occurrences:
            return
        styles = {item.semantic_key: item.resolved_style_id for item in caption_styles}
        recovered = list(self.recovered_captions)
        reference_ranges = [(item.source_start, item.source_end, item.tag) for item in (*soft_references, *references)]
        for wrapper, occurrence in zip(physical, occurrences, strict=True):
            style_id = styles.get(_CAPTION_STYLE_KEY[occurrence.kind])
            if style_id is None:
                raise DocxSemanticsV3Error("resolved-v4 occurrence lacks its caption-style binding")
            allowed = tuple(
                tag
                for start, end, tag in sorted((*reference_ranges, *citation_ranges))
                if occurrence.source_start <= start and end <= occurrence.source_end
            )
            caption, logical_object = prove_numbering_occurrence_sdt(
                wrapper,
                occurrence,
                caption_style_id=style_id,
                allowed_inline_tags=allowed,
            )
            _prove_caption_object_kind(occurrence.kind, logical_group_elements((logical_object,)))
            title, number = parse_caption(caption, occurrence.kind, required_bookmark=None)
            if number:
                raise DocxSemanticsV3Error("disabled resolved-v4 occurrence unexpectedly recovered a number")
            object_elements = logical_group_elements((logical_object,))
            self._block_elements[wrapper] = tuple(wrapper.find(qn("w:sdtContent")))
            recovered.append(
                RecoveredCaptionV3(
                    kind=occurrence.kind,
                    source_id=None,
                    title=title,
                    cached_number="",
                    caption_element=caption,
                    object_elements=object_elements,
                )
            )
        self.recovered_captions = tuple(recovered)


def _prove_one_source_identity(records: tuple[Any, ...]) -> None:
    identities = {item.source_sha256 for item in records}
    if len(identities) > 1:
        raise DocxSemanticsV3Error("resolved-v4 recovery maps do not share one source identity")


def _prove_merged_inline_order(document: Any, records: tuple[Any, ...]) -> None:
    from docx.oxml.ns import qn

    expected = sorted(records, key=lambda item: (item.source_start, item.source_end, item.tag))
    expected_tags = [item.tag for item in expected]
    if len(set(expected_tags)) != len(expected_tags):
        raise DocxSemanticsV3Error("resolved-v4 inline recovery tags are not globally unique")
    expected_set = set(expected_tags)
    physical = [
        tag
        for item in document.element.body.iter(qn("w:sdt"))
        if (tag := sdt_tag(item)) is not None and tag in expected_set
    ]
    if physical != expected_tags:
        raise DocxSemanticsV3Error("resolved-v4 merged inline physical order differs from source authority")


def _has_explicit_resolved_v4_signal(
    document: Any,
    owned: dict[str, tuple[int, Any]],
    targets: list[Any],
    caption_styles: tuple[Any, ...],
) -> bool:
    from docx.oxml.ns import qn

    if {
        NUMBERING_OCCURRENCE_MAP_NAMESPACE,
        CITATION_ITEM_MAP_NAMESPACE,
        CITATION_OCCURRENCE_MAP_NAMESPACE,
    }.intersection(owned):
        return True
    target_by_tag = {item.tag: item for item in targets}
    caption_style_ids = {item.resolved_style_id for item in caption_styles}
    for paragraph in document.element.body.iter(qn("w:p")):
        owner = paragraph.getparent()
        wrapper = None if owner is None else owner.getparent()
        target = target_by_tag.get(sdt_tag(wrapper) or "") if wrapper is not None else None
        if target is not None and target.kind == "heading":
            if paragraph.find(f"{qn('w:pPr')}/{qn('w:numPr')}") is not None:
                return True
            continue
        style = paragraph.find(f"{qn('w:pPr')}/{qn('w:pStyle')}")
        if style is None or style.get(qn("w:val")) not in caption_style_ids:
            continue
        instructions = [item.text or "" for item in paragraph.iter(qn("w:instrText"))]
        if not any("SEQ" in item for item in instructions):
            return True
        kind = target.kind if target is not None else caption_kind_for_paragraph_style(paragraph, caption_styles)
        if kind not in _CAPTION_COUNTER:
            continue
        bookmark = target.bookmark_name if target is not None else None
        try:
            _parse_v3_caption(paragraph, kind, required_bookmark=bookmark)
        except DocxSemanticsV3Error:
            return True
    return False


def _parse_resolved_v4_caption(
    paragraph: Any,
    kind: str,
    *,
    required_bookmark: str | None,
    inline_tokens: dict[str, str],
) -> tuple[str, str]:
    from docx.oxml.ns import qn

    counter = _CAPTION_COUNTER.get(kind)
    if counter is None:
        raise DocxSemanticsV3Error("resolved-v4 caption kind is outside the closed set")
    payload = tuple(item for item in paragraph if item.tag != qn("w:pPr"))
    direct_instructions = [
        item.text or "" for child in payload if child.tag != qn("w:sdt") for item in child.iter(qn("w:instrText"))
    ]
    has_sequence = any(re.fullmatch(rf" SEQ {counter} .+", item) is not None for item in direct_instructions)
    cursor = 0
    bookmark_id: str | None = None
    if not has_sequence:
        if any("SEQ" in item or "STYLEREF" in item for item in direct_instructions):
            raise DocxSemanticsV3Error("resolved-v4 disabled caption contains an unproven numbering field")
        if required_bookmark is not None:
            bookmark_id, cursor = _consume_bookmark_start(payload, cursor, required_bookmark)
            cursor = _consume_bookmark_end(payload, cursor, bookmark_id)
        elif any(item.tag in {qn("w:bookmarkStart"), qn("w:bookmarkEnd")} for item in payload):
            raise DocxSemanticsV3Error("resolved-v4 ID-less caption contains a bookmark")
        return _render_authored_payload(payload[cursor:], inline_tokens), ""

    label, cursor = _consume_text_run(payload, cursor)
    if not label or len(label) > 72:
        raise DocxSemanticsV3Error("resolved-v4 caption localized label is empty or oversized")
    if required_bookmark is not None:
        bookmark_id, cursor = _consume_bookmark_start(payload, cursor, required_bookmark)

    cached_parts: list[str] = []
    if cursor < len(payload) and _field_instruction(payload, cursor).startswith(" STYLEREF "):
        instruction, cached, cursor = _consume_complex_field(payload, cursor)
        if re.fullmatch(r' STYLEREF "[^"\r\n]{1,255}" \\n ', instruction) is None:
            raise DocxSemanticsV3Error("resolved-v4 caption STYLEREF instruction is not portable")
        separator, cursor = _consume_text_run(payload, cursor)
        if not 1 <= len(separator) <= 8:
            raise DocxSemanticsV3Error("resolved-v4 chapter separator is not portable")
        cached_parts.extend((cached, separator))

    instruction, cached, cursor = _consume_complex_field(payload, cursor)
    if not _valid_sequence_instruction(instruction, counter):
        raise DocxSemanticsV3Error("resolved-v4 caption SEQ instruction is not portable")
    cached_parts.append(cached)
    if required_bookmark is not None:
        assert bookmark_id is not None
        cursor = _consume_bookmark_end(payload, cursor, bookmark_id)
    elif any(item.tag in {qn("w:bookmarkStart"), qn("w:bookmarkEnd")} for item in payload):
        raise DocxSemanticsV3Error("resolved-v4 ID-less caption contains a bookmark")

    title = ""
    if cursor < len(payload):
        suffix, cursor = _consume_text_run(payload, cursor)
        if not suffix.startswith(" "):
            raise DocxSemanticsV3Error("resolved-v4 caption authored-content separator is not canonical")
        title = suffix[1:] + _render_authored_payload(payload[cursor:], inline_tokens)
        if not title:
            raise DocxSemanticsV3Error("resolved-v4 caption has an empty authored payload after its separator")
    return title, "".join(cached_parts)


def _consume_complex_field(payload: tuple[Any, ...], cursor: int) -> tuple[str, str, int]:
    from docx.oxml.ns import qn

    if cursor + 5 > len(payload):
        raise DocxSemanticsV3Error("resolved-v4 caption complex field is incomplete")
    _prove_field_marker(payload[cursor], "begin", dirty=True)
    instruction = _one_run_text(payload[cursor + 1], qn("w:instrText"), preserve=True)
    _prove_field_marker(payload[cursor + 2], "separate", dirty=False)
    cached = _one_run_text(payload[cursor + 3], qn("w:t"), preserve=None)
    _prove_field_marker(payload[cursor + 4], "end", dirty=False)
    if not cached:
        raise DocxSemanticsV3Error("resolved-v4 caption cached field result is empty")
    return instruction, cached, cursor + 5


def _field_instruction(payload: tuple[Any, ...], cursor: int) -> str:
    from docx.oxml.ns import qn

    if cursor + 1 >= len(payload):
        return ""
    try:
        return _one_run_text(payload[cursor + 1], qn("w:instrText"), preserve=True)
    except DocxSemanticsV3Error:
        return ""


def _prove_field_marker(run: Any, field_type: str, *, dirty: bool) -> None:
    from docx.oxml.ns import qn

    children = list(run)
    expected = (qn("w:fldCharType"), qn("w:dirty")) if dirty else (qn("w:fldCharType"),)
    if (
        run.tag != qn("w:r")
        or run.attrib
        or len(children) != 1
        or children[0].tag != qn("w:fldChar")
        or tuple(children[0].attrib) != expected
        or children[0].get(qn("w:fldCharType")) != field_type
        or children[0].get(qn("w:dirty")) != ("true" if dirty else None)
        or children[0].text is not None
        or children[0].tail is not None
        or len(children[0]) != 0
    ):
        raise DocxSemanticsV3Error("resolved-v4 caption field marker is not canonical")


def _one_run_text(run: Any, tag: str, *, preserve: bool | None) -> str:
    from docx.oxml.ns import qn

    xml_space = "{http://www.w3.org/XML/1998/namespace}space"
    children = list(run)
    if run.tag != qn("w:r") or run.attrib or len(children) != 1 or children[0].tag != tag or len(children[0]) != 0:
        raise DocxSemanticsV3Error("resolved-v4 caption run topology is not canonical")
    element = children[0]
    value = element.text or ""
    expected_preserve = value[:1].isspace() or value[-1:].isspace()
    if preserve is not None:
        expected_preserve = preserve
    expected_attributes = (xml_space,) if expected_preserve else ()
    if (
        tuple(element.attrib) != expected_attributes
        or element.get(xml_space) != ("preserve" if expected_preserve else None)
        or element.tail is not None
    ):
        raise DocxSemanticsV3Error("resolved-v4 caption text run is not canonical")
    return value


def _consume_text_run(payload: tuple[Any, ...], cursor: int) -> tuple[str, int]:
    from docx.oxml.ns import qn

    if cursor >= len(payload):
        raise DocxSemanticsV3Error("resolved-v4 caption text run is missing")
    return _one_run_text(payload[cursor], qn("w:t"), preserve=None), cursor + 1


def _consume_bookmark_start(payload: tuple[Any, ...], cursor: int, name: str) -> tuple[str, int]:
    from docx.oxml.ns import qn

    if cursor >= len(payload):
        raise DocxSemanticsV3Error("resolved-v4 caption bookmark start is missing")
    element = payload[cursor]
    bookmark_id = element.get(qn("w:id"))
    if (
        element.tag != qn("w:bookmarkStart")
        or tuple(element.attrib) != (qn("w:id"), qn("w:name"))
        or element.get(qn("w:name")) != name
        or bookmark_id is None
        or re.fullmatch(r"(?:0|[1-9][0-9]{0,9})", bookmark_id) is None
        or element.text is not None
        or element.tail is not None
        or len(element) != 0
    ):
        raise DocxSemanticsV3Error("resolved-v4 caption bookmark start is not canonical")
    return bookmark_id, cursor + 1


def _consume_bookmark_end(payload: tuple[Any, ...], cursor: int, bookmark_id: str) -> int:
    from docx.oxml.ns import qn

    if cursor >= len(payload):
        raise DocxSemanticsV3Error("resolved-v4 caption bookmark end is missing")
    element = payload[cursor]
    if (
        element.tag != qn("w:bookmarkEnd")
        or tuple(element.attrib) != (qn("w:id"),)
        or element.get(qn("w:id")) != bookmark_id
        or element.text is not None
        or element.tail is not None
        or len(element) != 0
    ):
        raise DocxSemanticsV3Error("resolved-v4 caption bookmark end is not canonical")
    return cursor + 1


def _render_authored_payload(payload: tuple[Any, ...], inline_tokens: dict[str, str]) -> str:
    from docx.oxml.ns import qn

    output: list[str] = []
    for child in payload:
        if child.tag == qn("w:sdt"):
            tag = sdt_tag(child)
            token = inline_tokens.get(tag or "")
            if token is None:
                raise DocxSemanticsV3Error("resolved-v4 caption contains an unowned inline SDT")
            output.append(token)
            continue
        if (
            list(child.iter(qn("w:instrText")))
            or list(child.iter(qn("w:fldChar")))
            or list(child.iter(qn("w:bookmarkStart")))
            or list(child.iter(qn("w:bookmarkEnd")))
        ):
            raise DocxSemanticsV3Error("resolved-v4 authored caption payload contains an unowned field or bookmark")
        output.append(visible_text(child))
    return "".join(output)


def _valid_sequence_instruction(instruction: str, counter: str) -> bool:
    prefix = re.escape(f" SEQ {counter} ")
    format_switch = re.escape("\\* ")
    patterns = (
        rf"{prefix}{format_switch}{_FIELD_FORMAT} ",
        rf"{prefix}{re.escape('\\r ')}[1-9][0-9]* {format_switch}{_FIELD_FORMAT} ",
        rf"{prefix}{re.escape('\\s ')}[1-6] {format_switch}{_FIELD_FORMAT} ",
    )
    return any(re.fullmatch(pattern, instruction) is not None for pattern in patterns)


def _recovery_pointer(value: ResolvedV4RecoveryMap, role: str) -> Any:
    for item in value.pointers:
        if item.role == role:
            return item
    raise DocxSemanticsV3Error(f"resolved-v4 recovery map lacks {role} pointer")


__all__ = [
    "RESOLVED_V4_SOURCE_SNAPSHOT_MISSING_DIAGNOSTIC",
    "RESOLVED_V4_SOURCE_SNAPSHOT_MISSING_MESSAGE",
    "ResolvedNumberingV4Recovery",
]

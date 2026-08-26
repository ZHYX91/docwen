"""Complete the request-owned DOCX style registry before rendering.

The public identity contract lives in :mod:`docwen_core.docx_styles`.  This
module owns only OOXML mechanics: recognize legacy localized styles, preflight
all conflicts, resolve host and DocWen IDs without losing user formatting,
update the main-document style domain, and keep Word's optional
``stylesWithEffects`` copy in sync.
"""

from __future__ import annotations

import unicodedata
from dataclasses import dataclass
from io import BytesIO
from typing import TYPE_CHECKING, Final, Never
from zipfile import ZIP_DEFLATED, ZIP_STORED, BadZipFile, ZipFile, ZipInfo

import lxml.etree as etree
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.styles.style import BaseStyle

from docwen_core.docx_styles import (
    MANAGED_DOCUMENT_STYLES,
    DocumentStyleCatalog,
    DocumentStyleDefinition,
    DocumentStyleFormat,
)

if TYPE_CHECKING:
    from docx.document import Document as DocumentObject

_PRIMARY_STYLES_PART: Final = "word/styles.xml"
_EFFECTS_STYLES_PART: Final = "word/stylesWithEffects.xml"
_REFERENCE_PART_NAMES: Final = frozenset(
    {
        "word/document.xml",
        "word/comments.xml",
        "word/footnotes.xml",
        "word/endnotes.xml",
        "word/numbering.xml",
    }
)
_REFERENCE_PREFIXES: Final = ("word/header", "word/footer")
_REFERENCE_TAGS: Final = (
    "pStyle",
    "rStyle",
    "tblStyle",
    "basedOn",
    "next",
    "link",
    "numStyleLink",
    "styleLink",
)
_REFERENCE_QNAMES: Final = tuple(qn(f"w:{tag}") for tag in _REFERENCE_TAGS)
_BASE_STYLE_DEFINITIONS: Final = (
    DocumentStyleDefinition("normal", "Normal", "paragraph", "", canonical_name="Normal"),
    DocumentStyleDefinition(
        "default_paragraph_font",
        "DefaultParagraphFont",
        "character",
        "",
        canonical_name="Default Paragraph Font",
    ),
    DocumentStyleDefinition("table_normal", "TableNormal", "table", "", canonical_name="Normal Table"),
    DocumentStyleDefinition("title", "Title", "paragraph", "Normal", canonical_name="Title"),
)
_ALL_STYLE_DEFINITIONS: Final = (*_BASE_STYLE_DEFINITIONS, *MANAGED_DOCUMENT_STYLES)
_CUSTOM_KEYS: Final = frozenset(
    definition.semantic_key for definition in MANAGED_DOCUMENT_STYLES if not definition.is_builtin
)
_LEGACY_STYLE_IDS: Final = {"code_block_caption": ("DocWenListingCaption",)}
_LEGACY_CODE_BLOCK_CAPTION_NAMES: Final = (
    "Listing Caption",
    "Listing-Beschriftung",
    "Leyenda de listado",
    "Légende de code",
    "コードリストキャプション",
    "코드 목록 캡션",
    "Legenda de listagem",
    "Подпись к листингу",
    "Chú thích mã",
    "代码清单题注",
    "程式碼清單題注",
)
_QUOTE_BORDER_COLORS: Final = ("CCCCCC", "C1C1C1", "B6B6B6", "ABABAB", "A0A0A0", "959595", "8A8A8A", "7F7F7F", "747474")


class ManagedStyleCompletionError(RuntimeError):
    """A stable, pre-render style completion failure."""

    def __init__(self, diagnostic_code: str, message: str, *, error_type: str) -> None:
        super().__init__(message)
        self.diagnostic_code = diagnostic_code
        self.error_type = error_type


@dataclass(frozen=True, slots=True)
class ManagedStyleConflict:
    """One preserved template collision and its request-local binding."""

    semantic_key: str
    requested_style_id: str
    resolved_style_id: str
    message: str
    code: str = "MD2DOCX-STYLE-COLLISION-PRESERVED"


@dataclass(frozen=True, slots=True)
class ManagedStyleBindings:
    """Request-owned stable styles, rebound after any save/reopen cycle."""

    styles: tuple[tuple[str, BaseStyle], ...]
    conflicts: tuple[ManagedStyleConflict, ...] = ()

    def get(self, semantic_key: str) -> BaseStyle:
        for key, style in self.styles:
            if key == semantic_key:
                return style
        raise KeyError(semantic_key)

    def style_id(self, semantic_key: str) -> str:
        return self.get(semantic_key).style_id

    @property
    def style_ids(self) -> tuple[tuple[str, str], ...]:
        return tuple((key, style.style_id) for key, style in self.styles)


@dataclass(frozen=True, slots=True)
class _StyleMatch:
    definition: DocumentStyleDefinition
    candidate: etree._Element | None
    output_name: str
    output_style_id: str


@dataclass(frozen=True, slots=True)
class _CompletionPlan:
    matches: tuple[_StyleMatch, ...]
    migrations: tuple[tuple[str, str], ...]
    conflicts: tuple[ManagedStyleConflict, ...] = ()


def complete_managed_styles(
    document: DocumentObject,
    catalog: DocumentStyleCatalog,
    *,
    code_font: str | None = None,
    code_background_color: str | None = None,
) -> tuple[DocumentObject, ManagedStyleBindings]:
    """Return a complete document and its request-local managed bindings.

    All identity and type conflicts are detected before the supplied document
    is mutated.  A save/ZIP/reopen round-trip then updates generic OOXML parts
    (footnotes/endnotes) and Word's parallel ``stylesWithEffects`` registry.
    Callers must discard every proxy obtained from the input document.
    """

    try:
        source = BytesIO()
        document.save(source)
        source_blob = source.getvalue()
        working = Document(BytesIO(source_blob))
        primary_root = working.styles.element
        effects_root = _preflight_serialized_package(source_blob)
        resolved_ids, resolution_conflicts = _resolve_output_style_ids(
            primary_root,
            effects_root,
            catalog,
        )
        plan = _preflight(primary_root, catalog, resolved_ids)
        effects_plan = _preflight(effects_root, catalog, resolved_ids) if effects_root is not None else None
        _apply_primary_plan(
            primary_root,
            plan,
            catalog,
            code_font=code_font,
            code_background_color=code_background_color,
        )

        serialized = BytesIO()
        working.save(serialized)
        completed_blob = _rewrite_serialized_package(
            serialized.getvalue(),
            plan,
            effects_plan,
            catalog,
            code_font=code_font,
            code_background_color=code_background_color,
        )
        completed = Document(BytesIO(completed_blob))
        conflicts = _dedupe_conflicts(
            (*resolution_conflicts, *plan.conflicts, *(effects_plan.conflicts if effects_plan is not None else ()))
        )
        bindings = _bind_and_validate(completed, catalog, plan, conflicts)
        expected_ids = tuple((match.definition.semantic_key, match.output_style_id) for match in plan.matches)
        effects_expected_ids = (
            tuple((match.definition.semantic_key, match.output_style_id) for match in effects_plan.matches)
            if effects_plan is not None
            else None
        )
        _validate_serialized_package(
            completed_blob,
            catalog,
            expected_ids=expected_ids,
            effects_expected_ids=effects_expected_ids,
        )
        return completed, bindings
    except ManagedStyleCompletionError:
        raise
    except (BadZipFile, etree.XMLSyntaxError, OSError, ValueError) as exc:
        raise ManagedStyleCompletionError(
            "MD2DOCX-STYLE-COMPLETION-ERROR",
            f"DOCX managed-style completion failed: {exc}",
            error_type="conversion_failed",
        ) from exc


def _preflight(
    root: etree._Element,
    catalog: DocumentStyleCatalog,
    resolved_ids: tuple[tuple[str, str], ...],
) -> _CompletionPlan:
    resolved_by_key = dict(resolved_ids)
    styles = tuple(root.findall(qn("w:style")))
    by_id: dict[str, etree._Element] = {}
    by_name: dict[str, list[etree._Element]] = {}
    for style in styles:
        style_id = style.get(qn("w:styleId"), "")
        if not style_id or style_id in by_id:
            _conflict("DOCX styles contain a missing or duplicate styleId.")
        by_id[style_id] = style
        name = _style_name(style)
        if name:
            by_name.setdefault(_normalize_name(name), []).append(style)

    claimed: dict[int, str] = {}
    matches: list[_StyleMatch] = []
    migrations: list[tuple[str, str]] = []
    conflicts: list[ManagedStyleConflict] = []
    for definition in _ALL_STYLE_DEFINITIONS:
        output_name = _output_name(definition, catalog)
        output_style_id = resolved_by_key[definition.semantic_key]
        recognition_names = _recognition_names(definition, catalog)
        id_hits = _unique_elements(
            candidate
            for style_id in (output_style_id, definition.style_id, *_LEGACY_STYLE_IDS.get(definition.semantic_key, ()))
            if (candidate := by_id.get(style_id)) is not None
            and _compatible_identity(candidate, definition, recognition_names)
        )
        all_name_hits = _unique_elements(
            candidate for name in recognition_names for candidate in by_name.get(_normalize_name(name), ())
        )
        name_hits = [candidate for candidate in all_name_hits if candidate.get(qn("w:type"), "") == definition.kind]
        stable_hit = by_id.get(definition.style_id)
        if definition.is_builtin:
            if stable_hit is not None and not _compatible_identity(stable_hit, definition, recognition_names):
                _conflict(f"Canonical built-in styleId {definition.style_id!r} is occupied by another identity.")
            if len(all_name_hits) != len(name_hits) or len(name_hits) > 1:
                _conflict(f"Canonical built-in style name for {definition.semantic_key!r} is ambiguous or mistyped.")
            if stable_hit is not None and name_hits and stable_hit is not name_hits[0]:
                _conflict(f"Canonical built-in styleId and name disagree for {definition.semantic_key!r}.")
        candidates = _unique_elements((*id_hits, *name_hits))
        candidate = _select_candidate(candidates, definition, output_style_id)
        if len(candidates) > 1 or len(all_name_hits) != len(name_hits):
            conflicts.append(
                _preserved_collision(
                    definition,
                    output_style_id,
                    "Multiple or wrong-type template styles matched the managed visible identity; "
                    "all non-selected styles were preserved.",
                )
            )
        if candidate is None and definition in _BASE_STYLE_DEFINITIONS:
            _conflict(f"Required base style {definition.style_id!r} is missing.")
        if candidate is not None:
            _require_type(candidate, definition)
            owner = claimed.setdefault(id(candidate), definition.semantic_key)
            if owner != definition.semantic_key:
                _conflict(f"One physical style matches both {owner!r} and {definition.semantic_key!r}.")
            old_id = candidate.get(qn("w:styleId"), "")
            if old_id != output_style_id:
                migrations.append((old_id, output_style_id))
        matches.append(_StyleMatch(definition, candidate, output_name, output_style_id))

    target_ids = [match.output_style_id for match in matches]
    if len(target_ids) != len(set(target_ids)):
        _internal("The managed style registry contains duplicate stable IDs.")
    migration_targets = {target for _source, target in migrations}
    if len(migration_targets) != len(migrations):
        _conflict("Multiple legacy styles would migrate to the same stable styleId.")
    return _CompletionPlan(tuple(matches), tuple(migrations), _dedupe_conflicts(conflicts))


def _resolve_output_style_ids(
    primary_root: etree._Element,
    effects_root: etree._Element | None,
    catalog: DocumentStyleCatalog,
) -> tuple[tuple[tuple[str, str], ...], tuple[ManagedStyleConflict, ...]]:
    """Choose one collision-free managed ID shared by both Word style parts."""

    roots = (primary_root,) if effects_root is None else (primary_root, effects_root)
    used_ids = {
        style_id
        for root in roots
        for style in root.findall(qn("w:style"))
        if (style_id := style.get(qn("w:styleId"), ""))
    }
    resolved: list[tuple[str, str]] = []
    conflicts: list[ManagedStyleConflict] = []
    for definition in _ALL_STYLE_DEFINITIONS:
        if definition in _BASE_STYLE_DEFINITIONS or definition.is_builtin:
            resolved.append(
                (
                    definition.semantic_key,
                    _resolve_builtin_style_id(primary_root, definition, catalog),
                )
            )
            continue
        recognition_names = _recognition_names(definition, catalog)
        recognition_identities = {_normalize_name(name) for name in recognition_names}
        visible_identity_conflict = any(
            _normalize_name(_style_name(candidate)) in recognition_identities
            and candidate.get(qn("w:type"), "") != definition.kind
            for root in roots
            for candidate in root.findall(qn("w:style"))
        )
        stable_is_safe = (
            all(
                (candidate := _style_by_id(root, definition.style_id)) is None
                or _compatible_identity(candidate, definition, recognition_names)
                for root in roots
            )
            and not visible_identity_conflict
        )
        if stable_is_safe:
            output_style_id = definition.style_id
        else:
            output_style_id = _allocate_collision_free_style_id(definition.style_id, used_ids)
            conflicts.append(
                _preserved_collision(
                    definition,
                    output_style_id,
                    f"Template styleId {definition.style_id!r} belongs to another identity or type; "
                    "the template style was preserved and the managed identity received a request-local styleId.",
                )
            )
        used_ids.add(output_style_id)
        resolved.append((definition.semantic_key, output_style_id))
    return tuple(resolved), tuple(conflicts)


def _resolve_builtin_style_id(
    primary_root: etree._Element,
    definition: DocumentStyleDefinition,
    catalog: DocumentStyleCatalog,
) -> str:
    """Reuse one compatible host ID; newly injected built-ins stay canonical."""

    recognition_names = _recognition_names(definition, catalog)
    stable = _style_by_id(primary_root, definition.style_id)
    if stable is not None:
        # Preflight owns the stable-ID conflict diagnostic. Returning the
        # canonical ID here keeps that failure deterministic.
        if _compatible_identity(stable, definition, recognition_names):
            return definition.style_id
        return definition.style_id

    candidates = _unique_elements(
        candidate
        for candidate in primary_root.findall(qn("w:style"))
        if _compatible_identity(candidate, definition, recognition_names)
    )
    if len(candidates) == 1:
        return candidates[0].get(qn("w:styleId"), definition.style_id)
    # Missing and ambiguous cases are distinguished by preflight. Missing
    # managed built-ins are injected with their canonical Word styleId.
    return definition.style_id


def _style_by_id(root: etree._Element, style_id: str) -> etree._Element | None:
    for element in root.findall(qn("w:style")):
        if element.get(qn("w:styleId"), "") == style_id:
            return element
    return None


def _allocate_collision_free_style_id(stable_id: str, used_ids: set[str]) -> str:
    index = 1
    while True:
        candidate = f"{stable_id}DocWen{index}"
        if candidate not in used_ids:
            return candidate
        index += 1


def _compatible_identity(
    candidate: etree._Element,
    definition: DocumentStyleDefinition,
    recognition_names: tuple[str, ...],
) -> bool:
    return candidate.get(qn("w:type"), "") == definition.kind and _normalize_name(_style_name(candidate)) in {
        _normalize_name(name) for name in recognition_names
    }


def _select_candidate(
    candidates: list[etree._Element],
    definition: DocumentStyleDefinition,
    output_style_id: str,
) -> etree._Element | None:
    if not candidates:
        return None
    preferred_ids = (
        output_style_id,
        definition.style_id,
        *_LEGACY_STYLE_IDS.get(definition.semantic_key, ()),
    )
    for style_id in preferred_ids:
        preferred = [item for item in candidates if item.get(qn("w:styleId"), "") == style_id]
        if len(preferred) == 1:
            return preferred[0]
    return candidates[0] if len(candidates) == 1 else None


def _preserved_collision(
    definition: DocumentStyleDefinition,
    resolved_style_id: str,
    message: str,
) -> ManagedStyleConflict:
    return ManagedStyleConflict(
        semantic_key=definition.semantic_key,
        requested_style_id=definition.style_id,
        resolved_style_id=resolved_style_id,
        message=message,
    )


def _dedupe_conflicts(
    conflicts: tuple[ManagedStyleConflict, ...] | list[ManagedStyleConflict],
) -> tuple[ManagedStyleConflict, ...]:
    result: list[ManagedStyleConflict] = []
    seen: set[tuple[str, str, str, str]] = set()
    for conflict in conflicts:
        identity = (
            conflict.semantic_key,
            conflict.requested_style_id,
            conflict.resolved_style_id,
            conflict.code,
        )
        if identity not in seen:
            result.append(conflict)
            seen.add(identity)
    return tuple(result)


def _apply_primary_plan(
    root: etree._Element,
    plan: _CompletionPlan,
    catalog: DocumentStyleCatalog,
    *,
    code_font: str | None,
    code_background_color: str | None,
) -> None:
    migration = dict(plan.migrations)
    resolved_by_stable_id = {match.definition.style_id: match.output_style_id for match in plan.matches}
    for match in plan.matches:
        definition = match.definition
        style = match.candidate
        if style is None:
            style = _new_style(
                definition,
                match.output_name,
                match.output_style_id,
                resolved_by_stable_id.get(definition.based_on, definition.based_on),
                catalog,
                code_font=code_font,
                code_background_color=code_background_color,
            )
            root.append(style)
        else:
            style.set(qn("w:styleId"), match.output_style_id)
            _set_style_name(style, match.output_name)
            _remove_children(style, "aliases")
            if definition.semantic_key in _CUSTOM_KEYS:
                style.set(qn("w:customStyle"), "1")
            else:
                custom_style = qn("w:customStyle")
                if custom_style in style.attrib:
                    del style.attrib[custom_style]
    rewrite_style_refs_in_element(root, migration)


def _rewrite_serialized_package(
    blob: bytes,
    plan: _CompletionPlan,
    effects_plan: _CompletionPlan | None,
    catalog: DocumentStyleCatalog,
    *,
    code_font: str | None,
    code_background_color: str | None,
) -> bytes:
    migration = dict(plan.migrations)
    source = BytesIO(blob)
    target = BytesIO()
    with ZipFile(source, "r") as archive_in, ZipFile(target, "w", allowZip64=True) as archive_out:
        names = archive_in.namelist()
        if _PRIMARY_STYLES_PART not in names:
            _internal("DOCX package has no primary styles.xml part.")
        saw_effects = False
        for item in archive_in.infolist():
            payload = archive_in.read(item.filename)
            if item.filename == _EFFECTS_STYLES_PART:
                saw_effects = True
                if effects_plan is None:
                    _internal("The parallel style part appeared after preflight.")
                payload = _complete_effects_part(
                    payload,
                    effects_plan,
                    catalog,
                    code_font=code_font,
                    code_background_color=code_background_color,
                )
            elif _is_reference_part(item.filename):
                root = etree.fromstring(payload)
                rewrite_style_refs_in_element(root, migration)
                payload = _serialize_xml(root)
            _write_zip_member(archive_out, item, payload)
        if effects_plan is not None and not saw_effects:
            _internal("The parallel style part disappeared after preflight.")
    return target.getvalue()


def _complete_effects_part(
    payload: bytes,
    preflight_plan: _CompletionPlan,
    catalog: DocumentStyleCatalog,
    *,
    code_font: str | None,
    code_background_color: str | None,
) -> bytes:
    effects_root = etree.fromstring(payload)
    resolved_ids = tuple((match.definition.semantic_key, match.output_style_id) for match in preflight_plan.matches)
    plan = _preflight(effects_root, catalog, resolved_ids)
    if _plan_signature(plan) != _plan_signature(preflight_plan):
        _internal("The parallel style registry changed after preflight.")
    _apply_primary_plan(
        effects_root,
        plan,
        catalog,
        code_font=code_font,
        code_background_color=code_background_color,
    )
    return _serialize_xml(effects_root)


def _preflight_serialized_package(
    blob: bytes,
) -> etree._Element | None:
    with ZipFile(BytesIO(blob), "r") as archive:
        names = archive.namelist()
        if len(names) != len(set(names)):
            _internal("DOCX package contains duplicate ZIP members.")
        if _PRIMARY_STYLES_PART not in names:
            _internal("DOCX package has no primary styles.xml part.")
        for name in names:
            if _is_reference_part(name):
                etree.fromstring(archive.read(name))
        if _EFFECTS_STYLES_PART not in names:
            return None
        return etree.fromstring(archive.read(_EFFECTS_STYLES_PART))


def rewrite_style_refs_in_element(root: etree._Element, old_to_new: dict[str, str]) -> int:
    """Rewrite the frozen set of WordprocessingML style-reference values."""

    changed = 0
    if not old_to_new:
        return changed
    for tag in _REFERENCE_QNAMES:
        for element in root.iter(tag):
            current = element.get(qn("w:val"))
            replacement = old_to_new.get(current or "")
            if replacement is not None:
                element.set(qn("w:val"), replacement)
                changed += 1
    return changed


def _plan_signature(
    plan: _CompletionPlan,
) -> tuple[tuple[tuple[str, str | None, str, str], ...], tuple[tuple[str, str], ...]]:
    return (
        tuple(
            (
                match.definition.semantic_key,
                match.candidate.get(qn("w:styleId")) if match.candidate is not None else None,
                match.output_name,
                match.output_style_id,
            )
            for match in plan.matches
        ),
        plan.migrations,
    )


def _new_style(
    definition: DocumentStyleDefinition,
    output_name: str,
    output_style_id: str,
    based_on_style_id: str,
    catalog: DocumentStyleCatalog,
    *,
    code_font: str | None,
    code_background_color: str | None,
) -> etree._Element:
    style = OxmlElement("w:style")
    style.set(qn("w:type"), definition.kind)
    style.set(qn("w:styleId"), output_style_id)
    if definition.semantic_key in _CUSTOM_KEYS:
        style.set(qn("w:customStyle"), "1")
    name = OxmlElement("w:name")
    name.set(qn("w:val"), output_name)
    style.append(name)
    if based_on_style_id:
        based_on = OxmlElement("w:basedOn")
        based_on.set(qn("w:val"), based_on_style_id)
        style.append(based_on)
    _apply_new_style_defaults(
        style,
        definition,
        catalog,
        code_font=code_font,
        code_background_color=code_background_color,
    )
    return style


def _apply_new_style_defaults(
    style: etree._Element,
    definition: DocumentStyleDefinition,
    catalog: DocumentStyleCatalog,
    *,
    code_font: str | None,
    code_background_color: str | None,
) -> None:
    style_format = catalog.format_for(definition.semantic_key)
    if style_format is not None:
        _apply_locale_format(style, style_format, heading_level=_heading_level(definition.semantic_key))
        return

    key = definition.semantic_key
    if key in {"footnote_text", "endnote_text"}:
        _append_flags(style, "semiHidden" if key == "footnote_text" else None, "unhideWhenUsed")
        _append_ui_priority(style, 99)
        p_pr = _child(style, "pPr")
        _set_empty_value(p_pr, "snapToGrid", "0")
        _set_attributes(p_pr, "ind", firstLine="0")
        r_pr = _child(style, "rPr")
        _set_attributes(r_pr, "sz", val="18")
        _set_attributes(r_pr, "szCs", val="18")
    elif key in {"footnote_reference", "endnote_reference"}:
        if key == "footnote_reference":
            _append_flags(style, "semiHidden", "unhideWhenUsed")
        else:
            _append_flags(style, "semiHidden", "unhideWhenUsed")
        _append_ui_priority(style, 99)
        _set_attributes(_child(style, "rPr"), "vertAlign", val="superscript")
    elif key in {"caption", "bibliography", "hyperlink"} or key.endswith("_caption"):
        return
    elif key == "code_block":
        _append_common_custom(style, priority=29)
        p_pr = _child(style, "pPr")
        _set_attributes(
            p_pr,
            "shd",
            val="clear",
            color="auto",
            fill=(code_background_color or "F5F5F5").upper(),
        )
        _set_attributes(p_pr, "spacing", before="120", after="120", line="240", lineRule="auto")
        _set_attributes(p_pr, "ind", firstLine="0")
        _monospace(_child(style, "rPr"), with_size=True, font_name=code_font)
    elif key == "inline_code":
        _append_common_custom(style, priority=29)
        r_pr = _child(style, "rPr")
        _monospace(r_pr, with_size=False, font_name=code_font)
        _set_attributes(
            r_pr,
            "shd",
            val="clear",
            color="auto",
            fill=(code_background_color or "F0F0F0").upper(),
        )
    elif key == "formula_block":
        _append_common_custom(style, priority=29)
        p_pr = _child(style, "pPr")
        _set_attributes(p_pr, "spacing", before="120", after="120")
        _set_attributes(p_pr, "ind", firstLine="0")
        _set_attributes(p_pr, "jc", val="center")
    elif key == "inline_formula":
        _append_common_custom(style, priority=29)
    elif key == "list_block":
        _append_common_custom(style, priority=34)
        p_pr = _child(style, "pPr")
        _set_attributes(p_pr, "ind", left="720", firstLine="0")
        p_pr.append(OxmlElement("w:contextualSpacing"))
    elif key.startswith("horizontal_rule_"):
        _append_common_custom(style, priority=99)
        p_pr = _child(style, "pPr")
        border = _child(_child(p_pr, "pBdr"), "bottom")
        _set_qn_attributes(border, val="single", color="auto", sz=str(4 * int(key[-1])), space="1")
        _set_attributes(p_pr, "spacing", before="120", after="120", line="240", lineRule="auto")
        _set_attributes(p_pr, "ind", firstLine="0")
    elif key in {"table_content", "table_header"}:
        _append_common_custom(style, priority=39)
        p_pr = _child(style, "pPr")
        _set_attributes(p_pr, "spacing", before="0", after="0", line="240", lineRule="auto")
        _set_attributes(p_pr, "ind", firstLine="0")
        _set_attributes(p_pr, "jc", val="center")
        r_pr = _child(style, "rPr")
        _set_attributes(r_pr, "sz", val="21")
        _set_attributes(r_pr, "szCs", val="21")
        if key == "table_header":
            r_pr.append(OxmlElement("w:b"))
            r_pr.append(OxmlElement("w:bCs"))
    elif key in {"three_line_table", "table_grid"}:
        _append_common_custom(style, priority=59)
        tbl_pr = _child(style, "tblPr")
        borders = _child(tbl_pr, "tblBorders")
        edges = (
            ("top", "bottom")
            if key == "three_line_table"
            else (
                "top",
                "left",
                "bottom",
                "right",
                "insideH",
                "insideV",
            )
        )
        for edge in edges:
            _set_qn_attributes(
                _child(borders, edge),
                val="single",
                color="auto",
                sz="12" if key == "three_line_table" else "4",
                space="0",
            )
        if key == "three_line_table":
            first_row = _child(style, "tblStylePr")
            first_row.set(qn("w:type"), "firstRow")
            bottom = _child(_child(_child(first_row, "tcPr"), "tcBorders"), "bottom")
            _set_qn_attributes(bottom, val="single", color="auto", sz="4", space="0")
    elif key == "image_paragraph":
        _append_common_custom(style, priority=39)
        p_pr = _child(style, "pPr")
        _set_attributes(p_pr, "spacing", before="120", after="120", line="240", lineRule="auto")
        _set_attributes(p_pr, "ind", firstLine="0")
        _set_attributes(p_pr, "jc", val="center")
    elif key.startswith("quote_"):
        _append_common_custom(style, priority=29)
        level = int(key.rsplit("_", 1)[1])
        p_pr = _child(style, "pPr")
        left = _child(_child(p_pr, "pBdr"), "left")
        _set_qn_attributes(left, val="single", color=_QUOTE_BORDER_COLORS[level - 1], sz="24", space="12")
        _set_attributes(p_pr, "shd", val="clear", color="auto", fill="F5F5F5")
        _set_attributes(p_pr, "spacing", before="120", after="120")
        _set_attributes(p_pr, "ind", left=str(480 + (level - 1) * 240), right="480", firstLine="0")
        r_pr = _child(style, "rPr")
        _set_attributes(r_pr, "color", val="666666")
        _set_attributes(r_pr, "sz", val="21")
        _set_attributes(r_pr, "szCs", val="21")


def _apply_locale_format(
    style: etree._Element,
    value: DocumentStyleFormat,
    *,
    heading_level: int | None,
) -> None:
    _append_common_custom(style, priority=9 if heading_level is not None else 1)
    p_pr = _child(style, "pPr")
    if value.spacing_before_twip or value.spacing_after_twip:
        _set_attributes(
            p_pr,
            "spacing",
            before=str(value.spacing_before_twip),
            after=str(value.spacing_after_twip),
        )
    indent: dict[str, str] = {}
    if value.first_line_indent_chars:
        indent["firstLineChars"] = str(value.first_line_indent_chars)
        indent["firstLine"] = str(round(value.font_size_pt * 2 * value.first_line_indent_chars / 100 * 10))
    elif value.first_line_indent_cm:
        indent["firstLine"] = str(round(value.first_line_indent_cm / 2.54 * 1440))
    else:
        indent["firstLine"] = "0"
    _set_attributes(p_pr, "ind", **indent)
    _set_attributes(p_pr, "jc", val=value.justification)
    if heading_level is not None:
        _set_attributes(p_pr, "outlineLvl", val=str(heading_level - 1))
    r_pr = _child(style, "rPr")
    _set_attributes(
        r_pr,
        "rFonts",
        ascii=value.ascii_font,
        hAnsi=value.ascii_font,
        eastAsia=value.east_asia_font,
    )
    if value.bold:
        r_pr.append(OxmlElement("w:b"))
        r_pr.append(OxmlElement("w:bCs"))
    half_points = str(round(value.font_size_pt * 2))
    _set_attributes(r_pr, "sz", val=half_points)
    _set_attributes(r_pr, "szCs", val=half_points)


def _append_common_custom(style: etree._Element, *, priority: int) -> None:
    style.append(OxmlElement("w:qFormat"))
    _append_ui_priority(style, priority)


def _append_ui_priority(style: etree._Element, value: int) -> None:
    _set_attributes(style, "uiPriority", val=str(value))


def _append_flags(style: etree._Element, *names: str | None) -> None:
    for name in names:
        if name:
            style.append(OxmlElement(f"w:{name}"))


def _monospace(r_pr: etree._Element, *, with_size: bool, font_name: str | None = None) -> None:
    if font_name:
        _set_attributes(r_pr, "rFonts", ascii=font_name, hAnsi=font_name, eastAsia=font_name, cs=font_name)
    else:
        _set_attributes(r_pr, "rFonts", ascii="Consolas", hAnsi="Consolas", eastAsia="等线", cs="Courier New")
    if with_size:
        _set_attributes(r_pr, "sz", val="20")
        _set_attributes(r_pr, "szCs", val="20")


def _child(parent: etree._Element, name: str) -> etree._Element:
    child = OxmlElement(f"w:{name}")
    parent.append(child)
    return child


def _set_attributes(parent: etree._Element, child_name: str, **values: str) -> etree._Element:
    child = _child(parent, child_name)
    _set_qn_attributes(child, **values)
    return child


def _set_empty_value(parent: etree._Element, child_name: str, value: str) -> etree._Element:
    return _set_attributes(parent, child_name, val=value)


def _set_qn_attributes(element: etree._Element, **values: str) -> None:
    for name, value in values.items():
        element.set(qn(f"w:{name}"), value)


def _heading_level(semantic_key: str) -> int | None:
    if not semantic_key.startswith("heading_"):
        return None
    return int(semantic_key.rsplit("_", 1)[1])


def _bind_and_validate(
    document: DocumentObject,
    catalog: DocumentStyleCatalog,
    plan: _CompletionPlan,
    conflicts: tuple[ManagedStyleConflict, ...],
) -> ManagedStyleBindings:
    by_id = {style.style_id: style for style in document.styles}
    bindings: list[tuple[str, BaseStyle]] = []
    managed_matches = {
        match.definition.semantic_key: match for match in plan.matches if match.definition in MANAGED_DOCUMENT_STYLES
    }
    for definition in MANAGED_DOCUMENT_STYLES:
        match = managed_matches[definition.semantic_key]
        style = by_id.get(match.output_style_id)
        if style is None:
            _internal(f"Managed style {match.output_style_id!r} is missing after completion.")
        style_element = style._element
        if style_element is None:
            _internal(f"Managed style {match.output_style_id!r} has no OOXML definition.")
        _require_type(style_element, definition)
        if _style_name(style_element) != _output_name(definition, catalog):
            _internal(f"Managed style {match.output_style_id!r} has the wrong output name.")
        if style_element.find(qn("w:aliases")) is not None:
            _internal(f"Managed style {match.output_style_id!r} still has aliases.")
        bindings.append((definition.semantic_key, style))
    return ManagedStyleBindings(tuple(bindings), conflicts)


def _validate_serialized_package(
    blob: bytes,
    catalog: DocumentStyleCatalog,
    *,
    expected_ids: tuple[tuple[str, str], ...],
    effects_expected_ids: tuple[tuple[str, str], ...] | None,
) -> None:
    expected_by_key = dict(expected_ids)
    with ZipFile(BytesIO(blob), "r") as archive:
        entry_names = [item.filename for item in archive.infolist()]
        if len(entry_names) != len(set(entry_names)):
            _internal("Final DOCX package contains duplicate ZIP member names.")
        primary = etree.fromstring(archive.read(_PRIMARY_STYLES_PART))
        primary_by_id = _style_elements_by_id(primary)
        style_types = {style_id: element.get(qn("w:type"), "") for style_id, element in primary_by_id.items()}
        for definition in MANAGED_DOCUMENT_STYLES:
            expected_id = expected_by_key[definition.semantic_key]
            element = primary_by_id.get(expected_id)
            if element is None or style_types.get(expected_id) != definition.kind:
                _internal(f"Managed style {expected_id!r} failed package validation.")
            if _style_name(element) != _output_name(definition, catalog):
                _internal(f"Managed style {expected_id!r} has the wrong final output name.")
            if element.find(qn("w:aliases")) is not None:
                _internal(f"Managed style {expected_id!r} regained aliases after rendering.")
        for item in archive.infolist():
            if not (_is_reference_part(item.filename) or item.filename == _PRIMARY_STYLES_PART):
                continue
            root = etree.fromstring(archive.read(item.filename))
            _validate_reference_root(root, style_types, item.filename)
        if _EFFECTS_STYLES_PART in archive.namelist():
            if effects_expected_ids is None:
                _internal("Parallel style registry appeared without a resolved identity map.")
            effects_expected_by_key = dict(effects_expected_ids)
            effects = etree.fromstring(archive.read(_EFFECTS_STYLES_PART))
            effects_by_id = _style_elements_by_id(effects)
            effects_types = {style_id: element.get(qn("w:type"), "") for style_id, element in effects_by_id.items()}
            for definition in MANAGED_DOCUMENT_STYLES:
                expected_id = effects_expected_by_key[definition.semantic_key]
                element = effects_by_id.get(expected_id)
                if element is None or element.get(qn("w:type")) != definition.kind:
                    _internal(f"Parallel style {expected_id!r} is missing or invalid.")
                if _style_name(element) != _output_name(definition, catalog):
                    _internal(f"Parallel style {expected_id!r} has the wrong output name.")
                if element.find(qn("w:aliases")) is not None:
                    _internal(f"Parallel style {expected_id!r} regained aliases after rendering.")
            _validate_reference_root(effects, effects_types, _EFFECTS_STYLES_PART)


def validate_managed_style_package(
    blob: bytes,
    catalog: DocumentStyleCatalog,
    bindings: ManagedStyleBindings | None = None,
) -> None:
    """Validate final DOCX style identities and every supported style reference.

    The renderer and the notes/numbering ZIP writers run after initial style
    completion, so the final serialized package must be checked again before
    an artifact is registered.
    """

    try:
        expected_ids = (
            bindings.style_ids
            if bindings is not None
            else tuple((definition.semantic_key, definition.style_id) for definition in MANAGED_DOCUMENT_STYLES)
        )
        _validate_serialized_package(
            blob,
            catalog,
            expected_ids=expected_ids,
            effects_expected_ids=expected_ids,
        )
    except ManagedStyleCompletionError:
        raise
    except (BadZipFile, etree.XMLSyntaxError, KeyError, OSError, ValueError) as exc:
        raise ManagedStyleCompletionError(
            "MD2DOCX-STYLE-COMPLETION-ERROR",
            f"Final DOCX managed-style validation failed: {exc}",
            error_type="conversion_failed",
        ) from exc


def _validate_reference_root(root: etree._Element, style_types: dict[str, str], part_name: str) -> None:
    expected_types = {qn("w:pStyle"): "paragraph", qn("w:rStyle"): "character", qn("w:tblStyle"): "table"}
    for tag in _REFERENCE_QNAMES:
        for element in root.iter(tag):
            style_id = element.get(qn("w:val"), "")
            if style_id not in style_types:
                _internal(f"{part_name} references missing style {style_id!r}.")
            expected = expected_types.get(tag)
            if expected is not None and style_types[style_id] != expected:
                _internal(f"{part_name} references {style_id!r} with the wrong OOXML type.")


def _recognition_names(
    definition: DocumentStyleDefinition,
    catalog: DocumentStyleCatalog,
) -> tuple[str, ...]:
    if definition.is_builtin:
        assert definition.canonical_name is not None
        return (definition.canonical_name,)
    names = catalog.recognition_names_for(definition.semantic_key)
    if definition.semantic_key == "code_block_caption":
        names += _LEGACY_CODE_BLOCK_CAPTION_NAMES
    return names


def _output_name(definition: DocumentStyleDefinition, catalog: DocumentStyleCatalog) -> str:
    if definition.is_builtin:
        assert definition.canonical_name is not None
        return definition.canonical_name
    return catalog.name_for(definition.semantic_key)


def _style_name(style: etree._Element) -> str:
    name = style.find(qn("w:name"))
    return name.get(qn("w:val"), "") if name is not None else ""


def _set_style_name(style: etree._Element, value: str) -> None:
    name = style.find(qn("w:name"))
    if name is None:
        name = OxmlElement("w:name")
        style.insert(0, name)
    name.set(qn("w:val"), value)


def _remove_children(style: etree._Element, child_name: str) -> None:
    for child in style.findall(qn(f"w:{child_name}")):
        style.remove(child)


def _require_type(element: etree._Element, definition: DocumentStyleDefinition) -> None:
    actual = element.get(qn("w:type"), "")
    if actual != definition.kind:
        _conflict(f"Style {definition.semantic_key!r} must be {definition.kind}, not {actual or 'unspecified'}.")


def _style_elements_by_id(root: etree._Element) -> dict[str, etree._Element]:
    result: dict[str, etree._Element] = {}
    for element in root.findall(qn("w:style")):
        style_id = element.get(qn("w:styleId"), "")
        if not style_id or style_id in result:
            _internal("A style definition part contains missing or duplicate style IDs.")
        result[style_id] = element
    return result


def _normalize_name(value: str) -> str:
    return unicodedata.normalize("NFC", value).casefold()


def _unique_elements(elements) -> list[etree._Element]:
    result: list[etree._Element] = []
    seen: set[int] = set()
    for element in elements:
        if id(element) not in seen:
            result.append(element)
            seen.add(id(element))
    return result


def _is_reference_part(name: str) -> bool:
    return name in _REFERENCE_PART_NAMES or (name.endswith(".xml") and name.startswith(_REFERENCE_PREFIXES))


def _serialize_xml(root: etree._Element) -> bytes:
    return etree.tostring(
        root,
        encoding="UTF-8",
        xml_declaration=True,
        standalone=True,
    )


def _write_zip_member(archive: ZipFile, source: ZipInfo, payload: bytes) -> None:
    target = ZipInfo(source.filename, source.date_time)
    target.comment = source.comment
    target.extra = source.extra
    target.create_system = source.create_system
    target.create_version = source.create_version
    target.extract_version = source.extract_version
    target.flag_bits = source.flag_bits
    target.volume = source.volume
    target.internal_attr = source.internal_attr
    target.external_attr = source.external_attr
    target.compress_type = source.compress_type if source.compress_type in {ZIP_STORED, ZIP_DEFLATED} else ZIP_DEFLATED
    archive.writestr(target, payload)


def _conflict(message: str) -> Never:
    raise ManagedStyleCompletionError(
        "MD2DOCX-STYLE-CONFLICT",
        message,
        error_type="invalid_input",
    )


def _internal(message: str) -> Never:
    raise ManagedStyleCompletionError(
        "MD2DOCX-STYLE-COMPLETION-ERROR",
        message,
        error_type="conversion_failed",
    )


__all__ = [
    "ManagedStyleBindings",
    "ManagedStyleCompletionError",
    "ManagedStyleConflict",
    "complete_managed_styles",
    "rewrite_style_refs_in_element",
    "validate_managed_style_package",
]

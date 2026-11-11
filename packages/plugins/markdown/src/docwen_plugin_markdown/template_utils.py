"""Template loading helpers for MD→DOCX conversion.

The runtime owns template discovery and content validation. This plugin only
loads the resolved path it receives, or creates the built-in fallback when no
template was selected.
"""

# pyright: reportGeneralTypeIssues=false

from __future__ import annotations

import re
from copy import deepcopy
from dataclasses import dataclass
from pathlib import Path
from typing import Any, cast
from zipfile import BadZipFile, ZipFile

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.section import Section
from docx.shared import Emu, Mm, Pt
from docx.text.paragraph import Paragraph
from docx.text.run import Run
from lxml import etree

from docwen_plugin_markdown.template_errors import BodyPlaceholderPlacementError, TemplatePackageError
from docwen_plugin_markdown.to_docx.sections import section_properties_at
from docwen_plugin_markdown.yaml_processor import BODY_PLACEHOLDER_ALIASES

# ── Built-in fallback template ──────────────────────────────────────────────

_BUILTIN_PLACEHOLDER = "{{正文}}"


def body_content_extent(doc: Any, placeholder: Any = None) -> tuple[int, int]:
    """Return usable column width/page height at the template body location."""
    properties = section_properties_at(doc, placeholder)
    section = Section(properties, doc.part) if properties is not None else None

    def length(name: str, default_mm: float) -> int:
        value = getattr(section, name, None)
        return int(Mm(default_mm) if value is None else value)

    width = length("page_width", 210) - length("left_margin", 25.4) - length("right_margin", 25.4) - length("gutter", 0)
    height = length("page_height", 297) - length("top_margin", 25.4) - length("bottom_margin", 25.4)
    columns = properties.find(qn("w:cols")) if properties is not None else None
    if columns is not None:
        widths = [int(item.get(qn("w:w"), "0")) * 635 for item in columns.findall(qn("w:col"))]
        if widths and all(value > 0 for value in widths):
            width = min(width, *widths)
        else:
            count = max(1, int(columns.get(qn("w:num"), "1")))
            spacing = int(columns.get(qn("w:space"), "720")) * 635
            width = (width - (count - 1) * spacing) // count
    return max(1, width), max(1, height)


# Default theme font fallback values: 宋体 (SimSun) for East-Asian,
# Calibri for Latin, 10.5 pt (五号) default size.
_DEFAULT_BODY_FONT: dict[str, Any] = {
    "name": "Calibri",
    "east_asia": "宋体",
    "size": Pt(10.5),
    "bold": False,
    "italic": False,
    "underline": False,
}


@dataclass(frozen=True, slots=True)
class BodyParagraphFormat:
    """Authored paragraph properties, independent of python-docx's API subset.

    The anchor owns visual formatting; the generated paragraph owns its style,
    numbering and document identity. Sections and range anchors stay at their
    original location through the body splice, never in this projection.
    """

    properties: tuple[etree._Element, ...]
    attributes: tuple[tuple[str, str], ...] = ()
    namespaces: tuple[tuple[str | None, str], ...] = ()
    # A style mapping can introduce competing indentation that the anchor did
    # not inherit. Retain its original style context only for those direct axes.
    source_style_id: str | None = None
    source_indent_attributes: tuple[tuple[str, str], ...] = ()

    @property
    def indent_attributes(self) -> tuple[tuple[str, str], ...]:
        indent = next((part for part in self.properties if part.tag == qn("w:ind")), None)
        return _property_attributes(indent) if indent is not None else ()

    def apply_to(self, paragraph: Any) -> None:
        projected_indent = False
        if self.properties or self.attributes:
            p_pr = paragraph._p.get_or_add_pPr()
            _preserve_property_namespaces(p_pr, self.namespaces)
            for name, value in self.attributes:
                if name not in p_pr.attrib:
                    p_pr.set(name, value)
            owned = {part.tag: part for part in p_pr}
            for source in self.properties:
                if not isinstance(getattr(source, "tag", None), str):
                    p_pr.append(deepcopy(source))
                    continue
                existing = owned.get(source.tag)
                if existing is None:
                    _insert_property(p_pr, source, _PARAGRAPH_PROPERTY_ORDER)
                    projected_indent |= source.tag == qn("w:ind")
                elif source.tag == qn("w:rPr"):
                    # Paragraph-mark formatting affects empty paragraphs too.
                    # Merge it without overwriting a generated mark's format.
                    _preserve_property_namespaces(existing, tuple(source.nsmap.items()))
                    for name, value in source.attrib.items():
                        if name not in existing.attrib:
                            existing.set(name, value)
                    owned_mark = {part.tag for part in existing}
                    for part in source:
                        if not isinstance(getattr(part, "tag", None), str):
                            existing.append(deepcopy(part))
                            continue
                        if part.tag not in owned_mark:
                            _insert_property(existing, part, _RUN_PROPERTY_ORDER)
        target_style = paragraph.style
        if (
            projected_indent
            and self.source_style_id is not None
            and target_style is not None
            and target_style.style_id != self.source_style_id
        ):
            indent = paragraph._p.pPr.ind if paragraph._p.pPr is not None else None
            if indent is not None:
                _preserve_mapped_indentation(paragraph, indent, dict(self.source_indent_attributes))


def inherited_indent_attributes(paragraph: Any) -> dict[str, str]:
    """Read indentation only, without flattening any other style properties."""
    return _indent_attributes(paragraph, include_direct=False)


def effective_indent_attributes(paragraph: Any) -> dict[str, str]:
    """Resolve direct formatting, styles and defaults for origin measurement."""
    return _indent_attributes(paragraph, include_direct=True)


def _indent_attributes(paragraph: Any, *, include_direct: bool) -> dict[str, str]:
    inherited: dict[str, str] = {}
    related = [(qn("w:firstLineChars"), qn("w:hangingChars"))]
    if include_direct:
        # Measuring the origin also resolves mutually exclusive physical axes.
        # The projection helper above retains its original attribute payload.
        related.append((qn("w:firstLine"), qn("w:hanging")))

    def inherit(source: Any) -> None:
        # A nearer character declaration selects/cancels the related group;
        # an older opposite-direction character value cannot revive it.
        blocked = {name for group in related if any(key in inherited for key in group) for name in group}
        for name, value in source.attrib.items():
            if name not in blocked:
                inherited.setdefault(name, value)

    direct = paragraph._p.pPr.ind if paragraph._p.pPr is not None else None
    if include_direct and direct is not None:
        inherit(direct)
    style = paragraph.style
    seen: set[Any] = set()
    while style is not None and style._element not in seen:
        seen.add(style._element)
        properties = style._element.find(qn("w:pPr"))
        source = properties.find(qn("w:ind")) if properties is not None else None
        if source is not None:
            inherit(source)
        style = style.base_style
    defaults = paragraph.part.styles.element.find(
        f"{qn('w:docDefaults')}/{qn('w:pPrDefault')}/{qn('w:pPr')}/{qn('w:ind')}"
    )
    if defaults is not None:
        inherit(defaults)
    return inherited


def _preserve_mapped_indentation(paragraph: Any, indent: Any, source: dict[str, str]) -> None:
    """Protect direct anchor axes from newly introduced output-style values.

    Inherited character units outrank later lengths in Word. Character zero
    cancels earlier related character units, but leaves inherited lengths in
    effect (MS-OI29500, 2.1.44). Restore only a changed related source value;
    an unchanged custom body style keeps its original inheritance verbatim.
    """
    target = inherited_indent_attributes(paragraph)
    for length_name, character_name in (
        ("left", "leftChars"),
        ("right", "rightChars"),
        ("start", "startChars"),
        ("end", "endChars"),
    ):
        length, characters = qn(f"w:{length_name}"), qn(f"w:{character_name}")
        if length in indent.attrib and characters not in indent.attrib:
            if int(source.get(characters, "0")) != int(target.get(characters, "0")):
                indent.set(characters, source.get(characters, "0"))
        elif (
            characters in indent.attrib
            and int(indent.get(characters, "0")) == 0
            and length not in indent.attrib
            and source.get(length, "0") != target.get(length, "0")
        ):
            indent.set(length, source.get(length, "0"))

    lengths = (qn("w:firstLine"), qn("w:hanging"))
    characters = (qn("w:firstLineChars"), qn("w:hangingChars"))

    def character_value(attributes: Any) -> tuple[str | None, int]:
        for name in reversed(characters):
            value = int(attributes.get(name, "0"))
            if value:
                return name, value
        return None, 0

    if any(name in indent.attrib for name in lengths) and not any(name in indent.attrib for name in characters):
        source_name, source_value = character_value(source)
        if (source_name, source_value) != character_value(target):
            if source_name is not None:
                indent.set(source_name, source[source_name])
            else:
                cancellation = characters[1] if lengths[1] in indent.attrib else characters[0]
                indent.set(cancellation, "0")
    elif (
        any(name in indent.attrib for name in characters)
        and character_value(indent.attrib)[1] == 0
        and not any(name in indent.attrib for name in lengths)
    ):
        # A character zero is a cancellation, not an unconditional zero length.
        # Preserve the original length fallback if the mapped style changes it.
        source_length = lengths[1] if lengths[1] in source else lengths[0]
        target_length = lengths[1] if lengths[1] in target else lengths[0]
        if (source_length, source.get(source_length, "0")) != (target_length, target.get(target_length, "0")):
            indent.set(source_length, source.get(source_length, "0"))


# Revision history is attached to the source text, not to newly generated text.
# Everything else in a property container remains eligible, including future
# extension properties that python-docx does not expose through its public API.
_REVISION_PROPERTY_TAGS = frozenset(
    qn(f"w:{name}") for name in ("pPrChange", "rPrChange", "ins", "del", "moveFrom", "moveTo")
)
_PARAGRAPH_OWNERSHIP_TAGS = _REVISION_PROPERTY_TAGS | frozenset(
    qn(f"w:{name}") for name in ("pStyle", "numPr", "sectPr", "divId", "cnfStyle")
)
_IDENTITY_ATTRIBUTES = frozenset(
    qn(name) for name in ("w:rsidR", "w:rsidRPr", "w:rsidP", "w:rsidDel", "w14:paraId", "w14:textId")
)
_PARAGRAPH_PROPERTY_ORDER = (
    "pStyle",
    "keepNext",
    "keepLines",
    "pageBreakBefore",
    "framePr",
    "widowControl",
    "numPr",
    "suppressLineNumbers",
    "pBdr",
    "shd",
    "tabs",
    "suppressAutoHyphens",
    "kinsoku",
    "wordWrap",
    "overflowPunct",
    "topLinePunct",
    "autoSpaceDE",
    "autoSpaceDN",
    "bidi",
    "adjustRightInd",
    "snapToGrid",
    "spacing",
    "ind",
    "contextualSpacing",
    "mirrorIndents",
    "suppressOverlap",
    "jc",
    "textDirection",
    "textAlignment",
    "textboxTightWrap",
    "outlineLvl",
    "divId",
    "cnfStyle",
    "rPr",
    "sectPr",
    "pPrChange",
)
# CT_RPr successor order: insert copied children before their next known peer
# to preserve WordprocessingML schema order even for mixed Markdown runs.
_RUN_PROPERTY_ORDER = (
    "rStyle",
    "rFonts",
    "b",
    "bCs",
    "i",
    "iCs",
    "caps",
    "smallCaps",
    "strike",
    "dstrike",
    "outline",
    "shadow",
    "emboss",
    "imprint",
    "noProof",
    "snapToGrid",
    "vanish",
    "webHidden",
    "color",
    "spacing",
    "w",
    "kern",
    "position",
    "sz",
    "szCs",
    "highlight",
    "u",
    "effect",
    "bdr",
    "shd",
    "fitText",
    "vertAlign",
    "rtl",
    "cs",
    "em",
    "lang",
    "eastAsianLayout",
    "specVanish",
    "oMath",
    "rPrChange",
)


def _format_snapshot(source: etree._Element, *, paragraph: bool = False) -> etree._Element:
    """Copy a format container, removing only source-owned history/identity."""
    snapshot = deepcopy(source)
    excluded = _PARAGRAPH_OWNERSHIP_TAGS if paragraph else _REVISION_PROPERTY_TAGS
    for child in list(snapshot):
        if child.tag in excluded:
            snapshot.remove(child)
    for element in snapshot.iter():
        for name in _IDENTITY_ATTRIBUTES:
            if name in element.attrib:
                del element.attrib[name]
        for child in list(element):
            if child.tag in _REVISION_PROPERTY_TAGS:
                element.remove(child)
    return snapshot


def _property_attributes(element: etree._Element) -> tuple[tuple[str, str], ...]:
    return tuple((str(name), str(value)) for name, value in element.attrib.items())


def _preserve_property_namespaces(target: etree._Element, namespaces: tuple[tuple[str | None, str], ...]) -> None:
    """Retain prefix context used by opaque values such as mc:Ignorable.

    Copying an attribute only retains its QName, not prefixes appearing in its
    value. Declare the source context on the existing property container; keep
    the paragraph/run and its property element identities intact.
    """
    source = dict(namespaces)
    if any(target.nsmap.get(prefix) != uri for prefix, uri in source.items()):
        prefixes = {prefix for prefix in (*target.nsmap, *source) if prefix is not None}
        # lxml accepts None for a default namespace; its installed stub only
        # declares named prefixes for this argument.
        etree.cleanup_namespaces(target, top_nsmap=cast(dict[str, str], source), keep_ns_prefixes=sorted(prefixes))


def _insert_property(parent: etree._Element, source: etree._Element, order: tuple[str, ...]) -> None:
    """Keep known schema order while retaining opaque extension properties."""
    tags = tuple(qn(f"w:{name}") for name in order)
    successor_tags = set(tags[tags.index(source.tag) + 1 :]) if source.tag in tags else set()
    successor = next((child for child in parent if child.tag in successor_tags), None)
    copied = deepcopy(source)
    if successor is None:
        parent.append(copied)
    else:
        successor.addprevious(copied)


_FONT_ATTRIBUTE_PEERS = {
    qn(f"w:{name}"): qn(f"w:{peer}")
    for explicit, theme in (
        ("ascii", "asciiTheme"),
        ("hAnsi", "hAnsiTheme"),
        ("eastAsia", "eastAsiaTheme"),
        ("cs", "cstheme"),
    )
    for name, peer in ((explicit, theme), (theme, explicit))
}
_BODY_RUN_TOGGLE_TAGS = frozenset(
    qn(f"w:{name}")
    for name in (
        "b",
        "bCs",
        "i",
        "iCs",
        "caps",
        "smallCaps",
        "strike",
        "dstrike",
        "outline",
        "shadow",
        "emboss",
        "imprint",
        "vanish",
    )
)


def _on_off_value(element: Any) -> bool:
    return element is not None and str(element.get(qn("w:val"), "1")).lower() not in {"0", "false", "off", "no"}


def _apply_style_toggles(style: Any, tag: str, value: bool) -> bool:
    seen: set[str] = set()
    while style is not None and style.style_id not in seen:
        seen.add(style.style_id)
        r_pr = style.element.find(qn("w:rPr"))
        if r_pr is not None and _on_off_value(r_pr.find(tag)):
            value = not value
        style = style.base_style
    return value


def _merge_missing_font_attributes(target: Any, source: Any, *protected: Any) -> None:
    """Treat a script's explicit font and theme font as one property."""
    owned_names = set().union(*(set(part.attrib) for part in (target, *protected)))
    for name, value in source.attrib.items():
        names = (name, _FONT_ATTRIBUTE_PEERS.get(name, name))
        if not any(key in owned_names for key in names):
            target.set(name, value)


def _style_run_properties(style: Any) -> dict[str, Any]:
    properties: dict[str, Any] = {}
    seen: set[str] = set()
    while style is not None and style.style_id not in seen:
        seen.add(style.style_id)
        r_pr = style.element.find(qn("w:rPr"))
        for source in _format_snapshot(r_pr) if r_pr is not None else ():
            if not isinstance(getattr(source, "tag", None), str):
                continue
            existing = properties.get(source.tag)
            if existing is None:
                properties[source.tag] = deepcopy(source)
            elif source.tag == qn("w:rFonts"):
                _merge_missing_font_attributes(existing, source)
        style = style.base_style
    return properties


@dataclass(frozen=True, slots=True)
class BodyRunFormat:
    """Authored direct character properties of the body anchor's first run."""

    properties: tuple[etree._Element, ...]
    attributes: tuple[tuple[str, str], ...] = ()
    namespaces: tuple[tuple[str | None, str], ...] = ()

    @classmethod
    def from_style(cls, style: Any, *, document: Any, fallback_font: dict[str, Any]) -> BodyRunFormat:
        """Materialize the ordinary body's effective style for mixed Headings."""
        properties = _style_run_properties(style)
        defaults = document.styles.element.find(f"{qn('w:docDefaults')}/{qn('w:rPrDefault')}/{qn('w:rPr')}")
        for source in _format_snapshot(defaults) if defaults is not None else ():
            if not isinstance(getattr(source, "tag", None), str):
                continue
            existing = properties.get(source.tag)
            if existing is None:
                properties[source.tag] = deepcopy(source)
            elif source.tag == qn("w:rFonts"):
                _merge_missing_font_attributes(existing, source)
        fallback = Run(cast(Any, OxmlElement("w:r")), document._body)
        fallback.font.name = fallback_font.get("name")
        fallback.font.size = fallback_font.get("size")
        fallback.bold = fallback_font.get("bold", False)
        fallback.italic = fallback_font.get("italic", False)
        fallback.underline = fallback_font.get("underline", False)
        fallback_properties = fallback._r.get_or_add_rPr()
        fonts = fallback_properties.get_or_add_rFonts()
        if fallback_font.get("east_asia"):
            fonts.set(qn("w:eastAsia"), fallback_font["east_asia"])
        for source in fallback_properties:
            existing = properties.get(source.tag)
            if existing is None:
                properties[source.tag] = deepcopy(source)
            elif source.tag == qn("w:rFonts"):
                _merge_missing_font_attributes(existing, source)
        if qn("w:color") not in properties:
            color = OxmlElement("w:color")
            color.set(qn("w:val"), "auto")
            properties[color.tag] = color
        # In styles, true toggles the inherited value and false leaves it alone.
        # Materialize the effective body result, including explicit resets, so
        # the appended text cannot inherit the Heading's emphasis instead.
        for tag in _BODY_RUN_TOGGLE_TAGS:
            value = _on_off_value(defaults.find(tag)) if defaults is not None else False
            value = _apply_style_toggles(style, tag, value)
            toggle = OxmlElement(f"w:{tag.rsplit('}', 1)[-1]}")
            if not value:
                toggle.set(qn("w:val"), "0")
            properties[tag] = toggle
        return cls(tuple(properties.values()))

    def apply_to(self, paragraph: Any, *, start_index: int = 0, resolve_character_toggles: bool = False) -> None:
        # Include child runs in hyperlinks and semantic wrappers. Each direct
        # or character-style property keeps its own authority; an rStyle alone
        # must not prevent unrelated body font attributes from being inherited.
        for element in list(paragraph._p.iter(qn("w:r")))[start_index:]:
            run = Run(element, paragraph)
            if not run.text:
                continue
            r_pr = run._r.get_or_add_rPr()
            _preserve_property_namespaces(r_pr, self.namespaces)
            for name, value in self.attributes:
                if name not in r_pr.attrib:
                    r_pr.set(name, value)
            # Read the target's own character style before adding the anchor's
            # style. Authored direct properties override an inherited anchor
            # style, but may not override generated code/link character styles.
            styled = _style_run_properties(run.style) if r_pr.find(qn("w:rStyle")) is not None else {}
            owned = {part.tag: part for part in r_pr}
            for source in self.properties:
                if not isinstance(getattr(source, "tag", None), str):
                    r_pr.append(deepcopy(source))
                    continue
                existing = owned.get(source.tag)
                protected = styled.get(source.tag)
                if source.tag == qn("w:rFonts"):
                    merged = deepcopy(existing) if existing is not None else OxmlElement("w:rFonts")
                    _merge_missing_font_attributes(merged, source, *((protected,) if protected is not None else ()))
                    if not merged.attrib:
                        continue
                    if existing is not None:
                        existing.attrib.update(merged.attrib)
                        continue
                    source = merged
                elif existing is not None:
                    continue
                elif protected is not None:
                    if not resolve_character_toggles or source.tag not in _BODY_RUN_TOGGLE_TAGS:
                        continue
                    value = _apply_style_toggles(run.style, source.tag, _on_off_value(source))
                    source = OxmlElement(f"w:{source.tag.rsplit('}', 1)[-1]}")
                    if not value:
                        source.set(qn("w:val"), "0")
                _insert_property(r_pr, source, _RUN_PROPERTY_ORDER)


def extract_body_run_format(doc: Document) -> BodyRunFormat | None:
    """Capture direct run formatting only, never inherited style defaults."""
    placeholder = find_body_placeholder(doc)
    if placeholder is None:
        return None
    for element in placeholder._p.iter(qn("w:r")):
        run = Run(element, placeholder)
        if not run.text.strip():
            continue
        r_pr = run._r.rPr
        if r_pr is None:
            return None
        snapshot = _format_snapshot(r_pr)
        properties = tuple(snapshot)
        attributes = _property_attributes(snapshot)
        return (
            BodyRunFormat(properties, attributes, tuple(snapshot.nsmap.items())) if properties or attributes else None
        )
    return None


def _create_builtin_template() -> Document:
    """Create a minimal A4 document with ``{{正文}}`` as body placeholder.

    The built-in template:
    - A4 page (21 cm × 29.7 cm)
    - 2.54 cm margins on all sides
    - A single paragraph containing ``{{正文}}``
    """
    doc = Document()

    # ── Page setup ──────────────────────────────────────────────────
    section = doc.sections[0]
    section.page_width = Mm(210)
    section.page_height = Mm(297)

    # Margins: 2.54 cm = 1 inch = 1440 twips ≈ 914400 EMUs
    margin = Emu(914400)
    section.top_margin = margin
    section.bottom_margin = margin
    section.left_margin = margin
    section.right_margin = margin

    # ── Placeholder paragraph ───────────────────────────────────────
    p = doc.add_paragraph()
    # Locale-aware managed body defaults own the unselected-template case.
    # Generated direct overrides would incorrectly mask those style defaults.
    p.add_run(_BUILTIN_PLACEHOLDER)

    return doc


def _set_run_east_asian_font(run, font_name: str) -> None:
    """Set the East Asian font on a run via ``<w:rFonts>``."""
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = rPr.makeelement(qn("w:rFonts"), {})
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:eastAsia"), font_name)


# ═══════════════════════════════════════════════════════════════════════════
# Template loading
# ═══════════════════════════════════════════════════════════════════════════


def load_template(template_path: str | Path | None = None) -> Document:
    """Load a template file or return the built-in fallback.

    Args:
        template_path: Path to a ``.docx`` file to use as template.
            ``None`` returns the built-in A4 document.

    Returns:
        A python-docx ``Document`` opened from *template_path* or
        the built-in fallback.
    """
    if template_path is not None:
        path = Path(template_path)
        try:
            with ZipFile(path, "r") as archive:
                names = archive.namelist()
                if len(names) != len(set(names)):
                    raise TemplatePackageError("DOCX template contains duplicate ZIP members.")
        except TemplatePackageError:
            raise
        except (BadZipFile, OSError) as exc:
            raise TemplatePackageError(f"DOCX template package is unreadable: {exc}") from exc
        with path.open("rb") as stream:
            return Document(stream)
    return _create_builtin_template()


def resolve_template(
    template_name: str | Path | None = None,
) -> Document:
    """Load a runtime-resolved template path or the built-in fallback."""
    if template_name is None or not str(template_name).strip():
        return _create_builtin_template()
    return load_template(Path(template_name).expanduser().resolve(strict=True))


# ═══════════════════════════════════════════════════════════════════════════
# Placeholder scanning
# ═══════════════════════════════════════════════════════════════════════════


def find_body_placeholder(doc: Document) -> Any | None:
    """Locate the ``{{正文}}`` or ``{{body}}`` paragraph in *doc*.

    Returns the python-docx ``Paragraph`` containing the placeholder,
    or ``None`` if no placeholder is found.
    """
    for p in doc.paragraphs:
        key = _whole_paragraph_placeholder_key(p.text)
        if key in BODY_PLACEHOLDER_ALIASES:
            return p
    return None


def validate_body_placeholder_placement(doc: Document) -> None:
    """Require one standalone main-body anchor, or no authored body marker."""
    markers = []
    # XML scanning also sees text in unsupported wrappers and across breaks.
    # Every accepted marker must be discoverable by the actual body finder.
    visible_anchors = {
        paragraph._p
        for paragraph in doc.paragraphs
        if _whole_paragraph_placeholder_key(paragraph.text) in BODY_PLACEHOLDER_ALIASES
    }
    pattern = re.compile(r"\{\{\s*([^{}\r\n]+?)\s*\}\}")
    roots = [doc.element]
    roots.extend(
        part.element
        for part in doc.part.related_parts.values()
        if hasattr(part, "element") and str(part.partname).startswith(("/word/header", "/word/footer"))
    )
    for root in roots:
        for paragraph in root.iter(qn("w:p")):
            texts = []
            for node in paragraph.iter(qn("w:t")):
                owner = node.getparent()
                while owner is not None and owner.tag != qn("w:p"):
                    owner = owner.getparent()
                if owner is paragraph:
                    texts.append(node.text or "")
            text = "".join(texts)
            for match in pattern.finditer(text):
                if match.group(1).strip() not in BODY_PLACEHOLDER_ALIASES:
                    continue
                if (
                    paragraph.getparent() is not doc.element.body
                    or text.strip() != match.group(0)
                    or paragraph not in visible_anchors
                ):
                    raise BodyPlaceholderPlacementError(
                        "A body placeholder must occupy a standalone main-document paragraph."
                    )
                markers.append(paragraph)
    if len(markers) > 1:
        raise BodyPlaceholderPlacementError("A template may contain only one body placeholder.")


def extract_body_font(doc: Document) -> dict[str, Any]:
    """Extract effective font settings from the body placeholder.

    Returns effective ``name``, ``east_asia``, ``size`` (a ``Pt`` value),
    ``bold``, ``italic``, and ``underline`` values.  A distributed template
    normally keeps these values on
    the placeholder's paragraph style rather than on the literal
    ``{{body}}`` run, so direct run properties take precedence and the
    paragraph style supplies the inherited values.  Falls back to the
    default theme font (宋体/Calibri/10.5 pt) when neither source defines a
    value.
    """
    placeholder = find_body_placeholder(doc)
    if placeholder is None:
        return dict(_DEFAULT_BODY_FONT)

    for element in placeholder._p.iter(qn("w:r")):
        run = Run(element, placeholder)
        if run.text and run.text.strip():
            font = run.font
            style = placeholder.style
            return {
                "name": font.name or _style_font_name(style) or _DEFAULT_BODY_FONT["name"],
                "east_asia": _get_east_asian_font(run, style),
                "size": font.size or _style_font_size(style) or _DEFAULT_BODY_FONT["size"],
                "bold": _effective_font_property(font, style, "bold", False),
                "italic": _effective_font_property(font, style, "italic", False),
                "underline": _effective_font_property(font, style, "underline", False),
            }

    return dict(_DEFAULT_BODY_FONT)


def extract_body_style(doc: Document) -> Any | None:
    """Return the paragraph style owned by the body placeholder.

    The placeholder is the template's authoritative body-format contract;
    using its style object avoids duplicating localized style names or
    unstable style IDs in plugin code.
    """
    placeholder = find_body_placeholder(doc)
    return placeholder.style if placeholder is not None else None


def extract_body_paragraph_format(doc: Document) -> BodyParagraphFormat | None:
    """Snapshot authored formatting, leaving structure to the body splice."""
    placeholder = find_body_placeholder(doc)
    if placeholder is None:
        return None
    p_pr = placeholder._p.pPr
    snapshot = _format_snapshot(p_pr, paragraph=True) if p_pr is not None else None
    return BodyParagraphFormat(
        properties=tuple(snapshot) if snapshot is not None else (),
        attributes=_property_attributes(snapshot) if snapshot is not None else (),
        namespaces=tuple(snapshot.nsmap.items()) if snapshot is not None else (),
        source_style_id=placeholder.style.style_id if placeholder.style is not None else None,
        source_indent_attributes=tuple(inherited_indent_attributes(placeholder).items()),
    )


def _style_font_name(style: Any) -> str | None:
    """Return the first explicitly defined Latin font in a style chain."""
    current = style
    while current is not None:
        name = current.font.name
        if name:
            return name
        current = current.base_style
    return None


def _style_font_size(style: Any):
    """Return the first explicitly defined font size in a style chain."""
    current = style
    while current is not None:
        size = current.font.size
        if size is not None:
            return size
        current = current.base_style
    return None


def _effective_font_property(font: Any, style: Any, property_name: str, default: Any) -> Any:
    """Resolve a boolean-like run font property through its style chain."""

    direct = getattr(font, property_name, None)
    if direct is not None:
        return direct
    current = style
    while current is not None:
        value = getattr(current.font, property_name, None)
        if value is not None:
            return value
        current = current.base_style
    return default


def _get_east_asian_font(run, style: Any = None) -> str:
    """Read ``w:eastAsia`` from a run, then its paragraph style chain."""
    try:
        rPr = run._r.find(qn("w:rPr"))
        if rPr is not None:
            rFonts = rPr.find(qn("w:rFonts"))
            if rFonts is not None:
                ea = rFonts.get(qn("w:eastAsia"))
                if ea:
                    return ea
    except Exception:
        pass

    current = style
    while current is not None:
        try:
            rPr = current._element.find(qn("w:rPr"))
            rFonts = rPr.find(qn("w:rFonts")) if rPr is not None else None
            if rFonts is not None:
                ea = rFonts.get(qn("w:eastAsia"))
                if ea:
                    return ea
        except Exception:
            pass
        current = current.base_style
    return _DEFAULT_BODY_FONT["east_asia"]


def scan_placeholders(doc: Document) -> dict[str, list[Any]]:
    """Walk *doc* and find all ``{{key}}`` occurrences.

    Returns:
        ``{key: [paragraph, ...]}`` mapping for every placeholder
        occurrence (without braces). A paragraph is listed once per key even
        when that key occurs more than once in the paragraph.
    """
    from docwen_plugin_markdown.template_ownership import iter_paragraph_placeholder_spans

    placeholders: dict[str, list[Any]] = {}
    roots = [doc.element]
    roots.extend(
        part.element
        for part in doc.part.related_parts.values()
        if hasattr(part, "element") and str(part.partname).startswith(("/word/header", "/word/footer"))
    )
    for root in roots:
        for element in root.iter(qn("w:p")):
            # Scan original XML as well as supported body/cell anchors so an
            # unsupported attachment location cannot silently go unnoticed.
            # Text in a nested textbox belongs only to its own paragraph.
            keys = dict.fromkeys(span.key for span in iter_paragraph_placeholder_spans(element))
            paragraph = Paragraph(element, doc)
            for key in keys:
                placeholders.setdefault(key, []).append(paragraph)
    return placeholders


def _whole_paragraph_placeholder_key(text: str) -> str | None:
    match = re.fullmatch(r"\s*\{\{\s*([^{}\r\n]+?)\s*\}\}\s*", text or "")
    if match is None:
        return None
    return match.group(1).strip()

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
from docx.text.run import Run

from docwen_plugin_markdown.yaml_processor import BODY_PLACEHOLDER_ALIASES

# ── Built-in fallback template ──────────────────────────────────────────────

_BUILTIN_PLACEHOLDER = "{{正文}}"


def body_content_extent(doc: Any, placeholder: Any = None) -> tuple[int, int]:
    """Return usable column width/page height at the template body location."""
    section = doc.sections[-1]
    if placeholder is not None:
        for element in (placeholder._p, *placeholder._p.itersiblings()):
            properties = element if element.tag == qn("w:sectPr") else element.find(qn("w:pPr") + "/" + qn("w:sectPr"))
            if properties is not None:
                section = Section(properties, doc.part)
                break

    def length(name: str, default_mm: float) -> int:
        value = getattr(section, name)
        return int(Mm(default_mm) if value is None else value)

    width = length("page_width", 210) - length("left_margin", 25.4) - length("right_margin", 25.4) - length("gutter", 0)
    height = length("page_height", 297) - length("top_margin", 25.4) - length("bottom_margin", 25.4)
    columns = section._sectPr.find(qn("w:cols"))
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


class TemplatePackageError(ValueError):
    """A stable invalid-input failure for a malformed template package."""

    diagnostic_code = "MD2DOCX-TEMPLATE-PACKAGE-INVALID"


class BodyPlaceholderPlacementError(TemplatePackageError):
    """A body marker exists outside the supported single-paragraph anchor."""

    diagnostic_code = "MD2DOCX-TEMPLATE-BODY-PLACEMENT"


@dataclass(frozen=True, slots=True)
class BodyParagraphFormat:
    """Direct paragraph-format contract projected from ``{{body}}``."""

    alignment: Any
    left_indent: Any
    right_indent: Any
    first_line_indent: Any
    space_before: Any
    space_after: Any
    line_spacing: Any
    line_spacing_rule: Any
    keep_together: bool | None
    keep_with_next: bool | None
    page_break_before: bool | None
    widow_control: bool | None
    # Preserve the authored w:ind payload, including character-width units
    # such as hangingChars/firstLineChars that python-docx's Length API drops.
    indent_attributes: tuple[tuple[str, str], ...] = ()
    # A style mapping can introduce competing indentation that the anchor did
    # not inherit. Retain its original style context only for those direct axes.
    source_style_id: str | None = None
    source_indent_attributes: tuple[tuple[str, str], ...] = ()

    def apply_to(self, paragraph: Any) -> None:
        paragraph_format = paragraph.paragraph_format
        for name in (
            "alignment",
            "left_indent",
            "right_indent",
            "first_line_indent",
            "space_before",
            "space_after",
            "line_spacing",
            "line_spacing_rule",
            "keep_together",
            "keep_with_next",
            "page_break_before",
            "widow_control",
        ):
            # A direct w:ind snapshot is the authoritative source for all
            # indentation attributes. Do not first round character units to
            # lengths or synthesize an incompatible firstLine/hanging pair.
            if self.indent_attributes and name in {"left_indent", "right_indent", "first_line_indent"}:
                continue
            value = getattr(self, name)
            if value is not None:
                setattr(paragraph_format, name, value)
        if self.indent_attributes:
            indent = paragraph._p.get_or_add_pPr().get_or_add_ind()
            indent.attrib.clear()
            for name, value in self.indent_attributes:
                indent.set(name, value)
        target_style = paragraph.style
        if (
            self.source_style_id is not None
            and target_style is not None
            and target_style.style_id != self.source_style_id
        ):
            indent = paragraph._p.pPr.ind if paragraph._p.pPr is not None else None
            if indent is not None:
                _preserve_mapped_indentation(paragraph, indent, dict(self.source_indent_attributes))


def _inherited_indent_attributes(paragraph: Any) -> dict[str, str]:
    """Read indentation only, without flattening any other style properties."""
    inherited: dict[str, str] = {}
    first_line_characters = (qn("w:firstLineChars"), qn("w:hangingChars"))

    def inherit(source: Any) -> None:
        # A nearer character declaration selects/cancels the related group;
        # an older opposite-direction character value cannot revive it.
        owns_first_line = any(name in inherited for name in first_line_characters)
        for name, value in source.attrib.items():
            if name not in first_line_characters or not owns_first_line:
                inherited.setdefault(name, value)

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
    target = _inherited_indent_attributes(paragraph)
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


# Only user-authored direct text properties are eligible for projection into
# regular Markdown body runs. In particular never project rStyle or structural
# field/anchor properties from the placeholder.
_BODY_RUN_DIRECT_TAGS = frozenset(
    qn(f"w:{name}") for name in ("rFonts", "b", "bCs", "i", "iCs", "strike", "dstrike", "color", "sz", "szCs", "u")
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
)

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
_BODY_RUN_TOGGLE_TAGS = frozenset(qn(f"w:{name}") for name in ("b", "bCs", "i", "iCs", "strike", "dstrike"))


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
        for source in r_pr if r_pr is not None else ():
            if source.tag not in _BODY_RUN_DIRECT_TAGS:
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

    properties: tuple[Any, ...]

    @classmethod
    def from_style(cls, style: Any, *, document: Any, fallback_font: dict[str, Any]) -> BodyRunFormat:
        """Materialize the ordinary body's effective style for mixed Headings."""
        properties = _style_run_properties(style)
        defaults = document.styles.element.find(f"{qn('w:docDefaults')}/{qn('w:rPrDefault')}/{qn('w:rPr')}")
        for source in defaults if defaults is not None else ():
            if source.tag not in _BODY_RUN_DIRECT_TAGS:
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
            styled = _style_run_properties(run.style) if r_pr.find(qn("w:rStyle")) is not None else {}
            for source in self.properties:
                existing = r_pr.find(source.tag)
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
                name = source.tag.rsplit("}", 1)[-1]
                successors = tuple(f"w:{part}" for part in _RUN_PROPERTY_ORDER[_RUN_PROPERTY_ORDER.index(name) + 1 :])
                r_pr.insert_element_before(deepcopy(cast(Any, source)), *successors)


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
        properties = tuple(deepcopy(part) for part in r_pr if part.tag in _BODY_RUN_DIRECT_TAGS)
        return BodyRunFormat(properties) if properties else None
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
            text = "".join(node.text or "" for node in paragraph.iter(qn("w:t")))
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
    """Snapshot direct body formatting without copying unrelated OOXML."""
    placeholder = find_body_placeholder(doc)
    if placeholder is None:
        return None
    paragraph_format = placeholder.paragraph_format
    p_pr = placeholder._p.pPr
    source_indent = p_pr.find(qn("w:ind")) if p_pr is not None else None
    authored_indent = tuple(source_indent.attrib.items()) if source_indent is not None else ()
    return BodyParagraphFormat(
        alignment=paragraph_format.alignment,
        left_indent=paragraph_format.left_indent,
        right_indent=paragraph_format.right_indent,
        first_line_indent=paragraph_format.first_line_indent,
        space_before=paragraph_format.space_before,
        space_after=paragraph_format.space_after,
        line_spacing=paragraph_format.line_spacing,
        line_spacing_rule=paragraph_format.line_spacing_rule,
        keep_together=paragraph_format.keep_together,
        keep_with_next=paragraph_format.keep_with_next,
        page_break_before=paragraph_format.page_break_before,
        widow_control=paragraph_format.widow_control,
        indent_attributes=authored_indent,
        source_style_id=placeholder.style.style_id if placeholder.style is not None else None,
        source_indent_attributes=tuple(_inherited_indent_attributes(placeholder).items()),
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
    pattern = re.compile(r"\{\{\s*([^{}\r\n]+?)\s*\}\}")
    placeholders: dict[str, list[Any]] = {}
    for p in _iter_document_paragraphs(doc):
        keys = dict.fromkeys(match.group(1).strip() for match in pattern.finditer(p.text))
        for key in keys:
            placeholders.setdefault(key, []).append(p)
    return placeholders


def _whole_paragraph_placeholder_key(text: str) -> str | None:
    match = re.fullmatch(r"\s*\{\{\s*([^{}\r\n]+?)\s*\}\}\s*", text or "")
    if match is None:
        return None
    return match.group(1).strip()


def _iter_document_paragraphs(doc: Document):
    # Keep the element wrappers themselves alive while traversing. Using
    # ``id(paragraph._element)`` is unsafe here because lxml may release a
    # wrapper between cells and reuse its Python object id, which silently
    # drops a later, distinct table paragraph from placeholder scanning.
    seen: set[Any] = set()

    for paragraph in doc.paragraphs:
        seen.add(paragraph._element)
        yield paragraph

    for table in doc.tables:
        yield from _iter_table_paragraphs(table, seen)


def _iter_table_paragraphs(table, seen: set[Any]):
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                element = paragraph._element
                if element in seen:
                    continue
                seen.add(element)
                yield paragraph
            for nested_table in cell.tables:
                yield from _iter_table_paragraphs(nested_table, seen)

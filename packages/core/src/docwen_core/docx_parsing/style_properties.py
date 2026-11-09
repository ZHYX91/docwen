"""Resolve native character formatting and paragraph properties from styles.

Paragraph styles and document defaults describe a block's presentation. Inline
Markdown records character-style and direct run formatting relative to that
baseline, so a bold Heading style does not invent authored strong emphasis.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from docwen_core.docx_parsing.xml_ns import NS_W

_W = f"{{{NS_W}}}"
_OFF_VALUES = frozenset({"0", "false", "off", "no", "none", "nil"})
_TOGGLES = {"b", "i", "strike", "vanish"}


def on_off_enabled(element: Any) -> bool:
    value = element.get(_W + "val")
    return value is None or value.strip().casefold() not in _OFF_VALUES


def _document_styles(parent: Any) -> Any | None:
    try:
        return parent.part.document.styles.element
    except (AttributeError, KeyError):
        return None


def _style_chain(styles: Any, style_id: str | None, kind: str) -> list[Any]:
    """Return base-first style elements, without trusting cyclic basedOn IDs."""
    if styles is None:
        return []
    by_id = {
        style.get(_W + "styleId"): style for style in styles.findall(_W + "style") if style.get(_W + "type") == kind
    }
    default = next((key for key, style in by_id.items() if style.get(_W + "default") in {"1", "true"}), None)
    seen: set[str] = set()
    chain: list[Any] = []
    key = default if style_id is None else style_id
    while key is not None and key not in seen and key in by_id:
        seen.add(key)
        style = by_id[key]
        chain.append(style)
        base = style.find(_W + "basedOn")
        key = base.get(_W + "val") if base is not None else None
    return list(reversed(chain))


def _paragraph_style_sources(paragraph: Any, styles: Any, property_name: str) -> list[Any]:
    style_ref = paragraph.find(f"{_W}pPr/{_W}pStyle") if paragraph is not None else None
    style_id = style_ref.get(_W + "val", "") if style_ref is not None else None
    sources: list[Any] = []
    if styles is not None:
        defaults = styles.find(f"{_W}docDefaults/{_W}{property_name}Default/{_W}{property_name}")
        if defaults is not None:
            sources.append(defaults)
    for style in _style_chain(styles, style_id, "paragraph"):
        properties = style.find(_W + property_name)
        if properties is not None:
            sources.append(properties)
    return sources


def paragraph_property_sources(paragraph: Any) -> list[Any]:
    """Return paragraph properties in descending precedence, including styles."""
    element = getattr(paragraph, "_p", paragraph)
    sources = _paragraph_style_sources(element, _document_styles(paragraph), "pPr")
    direct = element.find(_W + "pPr")
    if direct is not None:
        sources.append(direct)
    return list(reversed(sources))


@dataclass(frozen=True)
class RunFormatting:
    bold: bool = False
    italic: bool = False
    underline: bool = False
    strike: bool = False
    superscript: bool = False
    subscript: bool = False
    highlight: bool = False
    shading_fill: str = ""
    shading_value: str = ""
    hidden: bool = False


def resolve_run_formatting(run: Any, parent: Any = None, *, structural_character_style: bool = False) -> RunFormatting:
    """Resolve inline properties, respecting style toggles and explicit off.

    ``run`` can be typed python-docx XML or a ZIP-loaded note element. The
    document-owning ``parent`` supplies styles in both cases. Paragraph/default
    formatting is resolved as a baseline; only authored inline contributions
    produce Markdown formatting, while hidden text uses the full cascade.
    """
    styles = _document_styles(parent)
    paragraph = run.getparent() if hasattr(run, "getparent") else None
    while paragraph is not None and paragraph.tag != _W + "p":
        paragraph = paragraph.getparent()

    values: dict[str, Any] = dict.fromkeys(_TOGGLES, False)
    values.update(u=False, vertAlign="baseline", highlight=False, shd=("", ""))
    authored: set[str] = set()

    def apply(properties: Any, *, style: bool, inline: bool) -> None:
        if properties is None:
            return
        for name in _TOGGLES:
            prop = properties.find(_W + name)
            if prop is None:
                continue
            enabled = on_off_enabled(prop)
            if style:
                if enabled:
                    values[name] = not values[name]
            else:
                values[name] = enabled
            if inline and (not style or enabled):
                authored.add(name)
        for name in ("u", "vertAlign", "highlight", "shd"):
            prop = properties.find(_W + name)
            if prop is None:
                continue
            if name == "u":
                values[name] = on_off_enabled(prop)
            elif name == "vertAlign":
                values[name] = prop.get(_W + "val", "baseline")
            elif name == "highlight":
                values[name] = (prop.get(_W + "val") or "none").casefold() not in {"none", ""}
            else:
                values[name] = ((prop.get(_W + "fill") or "").upper(), (prop.get(_W + "val") or "").lower())
            if inline:
                authored.add(name)

    baseline_sources = _paragraph_style_sources(paragraph, styles, "rPr")
    defaults = styles.find(f"{_W}docDefaults/{_W}rPrDefault/{_W}rPr") if styles is not None else None
    for properties in baseline_sources:
        apply(properties, style=properties is not defaults, inline=False)
    baseline = values.copy()
    direct = run.find(_W + "rPr")
    character_ref = direct.find(_W + "rStyle") if direct is not None else None
    if character_ref is not None:
        for style in _style_chain(styles, character_ref.get(_W + "val", ""), "character"):
            apply(style.find(_W + "rPr"), style=True, inline=not structural_character_style)
    if structural_character_style:
        baseline = values.copy()
    apply(direct, style=False, inline=True)

    def enabled(name: str) -> bool:
        return bool(values[name]) and (name in authored or values[name] != baseline[name])

    vertical = values["vertAlign"] if "vertAlign" in authored else "baseline"
    fill, shading = values["shd"] if "shd" in authored else ("", "")
    return RunFormatting(
        bold=enabled("b"),
        italic=enabled("i"),
        underline=enabled("u"),
        strike=enabled("strike"),
        superscript=vertical == "superscript",
        subscript=vertical == "subscript",
        highlight=enabled("highlight"),
        shading_fill=fill,
        shading_value=shading,
        hidden=bool(values["vanish"]),
    )

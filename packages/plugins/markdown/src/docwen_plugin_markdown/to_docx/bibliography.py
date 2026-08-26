"""Typed bibliography resource and template ownership for Markdown to DOCX."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

from docwen_core.models.semantic_document import SemanticBibliographyFragment
from docwen_core.semantic_bibliography import (
    SEMANTIC_BIBLIOGRAPHY_MEDIA_TYPE,
    SemanticBibliographyResourceError,
    parse_semantic_bibliography,
)
from docwen_plugin_markdown.yaml_processor import BODY_PLACEHOLDER_ALIASES

_BIBLIOGRAPHY_PLACEHOLDER = "bibliography"


@dataclass(frozen=True, slots=True)
class BibliographyConversionError(ValueError):
    """Stable plugin-boundary bibliography failure."""

    diagnostic_code: str
    message: str
    error_type: str = "invalid_input"

    def __str__(self) -> str:
        return self.message


def load_bibliography_resource(declared_inputs: tuple[Any, ...]) -> SemanticBibliographyFragment | None:
    """Load zero or one exact typed bibliography input."""

    resources = tuple(item for item in declared_inputs if item.input_role == "bibliography")
    if len(resources) > 1:
        raise BibliographyConversionError(
            "MD2DOCX-BIBLIOGRAPHY-RESOURCE-INVALID",
            "Markdown to DOCX accepts at most one bibliography resource.",
        )
    if not resources:
        return None
    resource = resources[0]
    if resource.input_kind != "resource" or resource.media_type != SEMANTIC_BIBLIOGRAPHY_MEDIA_TYPE:
        raise BibliographyConversionError(
            "MD2DOCX-BIBLIOGRAPHY-RESOURCE-INVALID",
            "Bibliography input must be a resource with the semantic bibliography v1 media type.",
        )
    try:
        payload = Path(resource.path).read_bytes()
    except OSError as exc:
        raise BibliographyConversionError(
            "MD2DOCX-BIBLIOGRAPHY-RESOURCE-INVALID",
            "Bibliography resource could not be read safely.",
        ) from exc
    try:
        return parse_semantic_bibliography(payload)
    except SemanticBibliographyResourceError as exc:
        raise BibliographyConversionError(
            "MD2DOCX-BIBLIOGRAPHY-RESOURCE-INVALID",
            str(exc),
        ) from exc


def validate_bibliography_placement(
    document: Any,
    fragment: SemanticBibliographyFragment | None,
) -> None:
    """Validate explicit placement or the unique safe synthesis anchor."""

    occurrences, direct_anchors = _bibliography_occurrences(document)
    if occurrences:
        if len(occurrences) != 1 or len(direct_anchors) != 1:
            raise BibliographyConversionError(
                "MD2DOCX-BIBLIOGRAPHY-PLACEHOLDER-INVALID",
                "The bibliography placeholder must occur exactly once as the sole visible content of a direct body paragraph.",
            )
        return
    if fragment is None or not fragment.entries:
        return
    if len(_direct_body_markers(document)) != 1:
        raise BibliographyConversionError(
            "MD2DOCX-BIBLIOGRAPHY-PLACEHOLDER-INVALID",
            "A non-empty bibliography requires one explicit marker or one unique direct-body body marker.",
        )


def prepare_bibliography_anchor(
    document: Any,
    fragment: SemanticBibliographyFragment | None,
    *,
    bibliography_style_id: str,
) -> Any | None:
    """Rebind the proven anchor after style completion, synthesizing only if allowed."""

    validate_bibliography_placement(document, fragment)
    occurrences, direct_anchors = _bibliography_occurrences(document)
    if occurrences:
        return direct_anchors[0]
    if fragment is None or not fragment.entries:
        return None

    from docx.oxml import OxmlElement
    from docx.text.paragraph import Paragraph

    body_marker = _direct_body_markers(document)[0]
    paragraph_element: Any = OxmlElement("w:p")
    paragraph = Paragraph(paragraph_element, document)
    paragraph.style = bibliography_style_id
    paragraph.add_run("{{ bibliography }}")
    body_marker._p.addnext(paragraph_element)
    return paragraph


def without_reserved_bibliography_placeholder(placeholder_map: dict[str, list[Any]]) -> dict[str, list[Any]]:
    """Prevent YAML processors and cleanup rules from consuming the reserved anchor."""

    return {key: paragraphs for key, paragraphs in placeholder_map.items() if key != _BIBLIOGRAPHY_PLACEHOLDER}


def _bibliography_occurrences(document: Any) -> tuple[list[Any], list[Any]]:
    occurrences: list[Any] = []
    direct: list[Any] = []
    body = document.element.body
    for part in document.part.package.parts:
        root = getattr(part, "element", None)
        if root is None:
            continue
        for paragraph_element in root.iter(_qn("w:p")):
            visible = _visible_text(paragraph_element)
            if "{{" not in visible or "}}" not in visible:
                continue
            if not _contains_bibliography_placeholder(visible):
                continue
            occurrences.append(paragraph_element)
            if (
                part is document.part
                and paragraph_element.getparent() is body
                and _whole_placeholder_key(visible) == _BIBLIOGRAPHY_PLACEHOLDER
            ):
                from docx.text.paragraph import Paragraph

                direct.append(Paragraph(paragraph_element, document))
    return occurrences, direct


def _direct_body_markers(document: Any) -> list[Any]:
    from docx.text.paragraph import Paragraph

    output: list[Any] = []
    for element in document.element.body:
        if element.tag != _qn("w:p"):
            continue
        paragraph = Paragraph(element, document)
        if _whole_placeholder_key(_visible_text(element)) in BODY_PLACEHOLDER_ALIASES:
            output.append(paragraph)
    return output


def _contains_bibliography_placeholder(text: str) -> bool:
    import re

    return re.search(r"\{\{\s*bibliography\s*\}\}", text) is not None


def _whole_placeholder_key(text: str) -> str | None:
    import re

    match = re.fullmatch(r"\s*\{\{\s*([^{}\r\n]+?)\s*\}\}\s*", text or "")
    return None if match is None else match.group(1).strip()


def _visible_text(paragraph_element: Any) -> str:
    output: list[str] = []
    for element in paragraph_element.iter():
        if element.tag == _qn("w:t"):
            output.append(element.text or "")
        elif element.tag == _qn("w:tab"):
            output.append("\t")
        elif element.tag in {_qn("w:br"), _qn("w:cr")}:
            output.append("\n")
    return "".join(output)


def _qn(tag: str) -> str:
    from docx.oxml.ns import qn

    return qn(tag)


__all__ = [
    "BibliographyConversionError",
    "load_bibliography_resource",
    "prepare_bibliography_anchor",
    "validate_bibliography_placement",
    "without_reserved_bibliography_placeholder",
]

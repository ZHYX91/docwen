"""Proofread paragraph skip policy."""

from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass

from docwen_core.docx_parsing.format_features import (
    StyleDetectorConfig,
    detect_paragraph_style_type,
    style_detector_config_from_document_config,
)


@dataclass(frozen=True)
class ProofreadSkipOptions:
    """Options controlling paragraphs skipped during proofreading."""

    code_blocks: bool = True
    quote_blocks: bool = False
    style_detector_config: StyleDetectorConfig | None = None


def _mapping_section(config: object, key: str) -> Mapping[str, object]:
    if isinstance(config, Mapping):
        value = config.get(key, {})
    else:
        getter = getattr(config, "get", None)
        value = getter(key, {}) if callable(getter) else {}
    return value if isinstance(value, Mapping) else {}


def resolve_skip_options(context: object) -> ProofreadSkipOptions:
    """Resolve skip options from runtime config and request options."""
    config = getattr(context, "config", {}) or {}
    proofread = _mapping_section(config, "proofread")
    skip = _mapping_section(proofread, "skip")

    code_blocks = bool(skip.get("code_blocks", True))
    quote_blocks = bool(skip.get("quote_blocks", False))

    request = getattr(context, "request", None)
    request_options = getattr(request, "options", {}) if request is not None else {}
    if isinstance(request_options, dict):
        if "skip_code_blocks" in request_options:
            code_blocks = bool(request_options["skip_code_blocks"])
        if "skip_quote_blocks" in request_options:
            quote_blocks = bool(request_options["skip_quote_blocks"])

    document = _mapping_section(config, "document")
    style_detector_config = (
        style_detector_config_from_document_config(document) if isinstance(document.get("style"), Mapping) else None
    )
    return ProofreadSkipOptions(
        code_blocks=code_blocks,
        quote_blocks=quote_blocks,
        style_detector_config=style_detector_config,
    )


def should_skip_docx_paragraph(paragraph: object, options: ProofreadSkipOptions) -> bool:
    """Return True when a DOCX paragraph should not be proofread."""
    if options.code_blocks and _is_code_paragraph(paragraph, options.style_detector_config):
        return True
    # Mixed text/non-text paragraphs remain eligible.  The validator edits
    # text-run boundaries in place, so drawings, formulas, and shapes do not
    # need the old paragraph-rebuild exclusion.  Pure non-text paragraphs are
    # filtered earlier by their empty text projection.
    return bool(options.quote_blocks and _is_quote_paragraph(paragraph, options.style_detector_config))


def _is_code_paragraph(paragraph: object, config: StyleDetectorConfig | None) -> bool:
    style_type, _level = detect_paragraph_style_type(paragraph, config=config)
    return style_type == "code_block"


def _is_quote_paragraph(paragraph: object, config: StyleDetectorConfig | None) -> bool:
    style_type, _level = detect_paragraph_style_type(paragraph, config=config)
    if style_type == "quote":
        return True

    text = getattr(paragraph, "text", "")
    return isinstance(text, str) and text.lstrip().startswith(">")


def _has_non_text_content(paragraph: object) -> bool:
    """Check for images, formulas, OLE objects, or shapes in paragraph XML.

    Inspects the raw XML of a python-docx paragraph element for known
    OpenXML non-text content tags.  This is a conservative heuristic —
    the presence of a ``w:drawing`` or ``m:oMath`` element indicates that
    the paragraph contains non-textual content that should be skipped.
    """
    # Guard: paragraph must have a _element with an .xml property
    elem = getattr(paragraph, "_element", None)
    if elem is None:
        return False
    xml: str | None = getattr(elem, "xml", None)
    if not xml:
        return False

    # Check for known non-text content tags (case-insensitive substring match)
    lower_xml = xml.lower()
    return any(tag in lower_xml for tag in ("w:drawing", "m:omath", "w:object", "v:shape"))

"""Request-owned export policy for PPTX to Markdown conversion."""

from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass, replace
from typing import TYPE_CHECKING, Any

from docwen_core.export_semantics import MarkdownExportSemantics

if TYPE_CHECKING:
    from docwen_core.protocols.execution_context import ConverterContext


@dataclass(frozen=True, slots=True)
class PresentationMarkdownRequestPolicy:
    """All mutable export inputs frozen for one presentation request."""

    export: MarkdownExportSemantics
    ocr_blockquote_title: str


def build_presentation_markdown_request_policy(
    context: ConverterContext,
    options: Mapping[str, Any],
) -> PresentationMarkdownRequestPolicy:
    """Project one authoritative request snapshot into immutable policy."""
    conversion = _section(context.config.get("conversion", {}))
    link = _section(context.config.get("link", {}))
    export_config = _section(context.config.get("export", {}))
    output = _section(context.config.get("output", {}))
    export = MarkdownExportSemantics.from_config(
        link_format=_section(link.get("format", {})),
        ocr_output=_section(conversion.get("ocr_output", {})),
        export_cfg=export_config,
        conversion_cfg=conversion,
        intermediate_files_cfg=_section(output.get("intermediate_files", {})),
        locale=_request_locale(context.config, options),
    )
    injected_title = getattr(context, "ocr_blockquote_title", None)
    if isinstance(injected_title, str):
        ocr_blockquote_title = injected_title.strip()
    elif export.ocr_blockquote_title_enabled:
        ocr_blockquote_title = export.ocr_blockquote_title_override_text.strip()
    else:
        ocr_blockquote_title = ""

    image_extraction_mode = str(options.get("image_mode") or export.image_extraction_mode)
    ocr_placement_mode = str(options.get("ocr_placement") or export.ocr_placement_mode)
    if image_extraction_mode.strip().lower() == "base64":
        ocr_placement_mode = "main_md"
    export = replace(
        export,
        image_extraction_mode=image_extraction_mode,
        ocr_placement_mode=ocr_placement_mode,
        image_link_style=str(options.get("image_link_style") or export.image_link_style),
    )
    return PresentationMarkdownRequestPolicy(
        export=export,
        ocr_blockquote_title=ocr_blockquote_title,
    )


def _section(value: Any) -> dict[str, Any]:
    if isinstance(value, Mapping):
        return {str(key): item for key, item in value.items()}
    return {}


def _request_locale(config: Any, options: Mapping[str, Any]) -> str:
    explicit = options.get("locale")
    if explicit is not None and str(explicit).strip():
        return str(explicit)
    gui = _section(config.get("gui", {}))
    language = gui.get("language", {})
    locale = language.get("locale") if isinstance(language, Mapping) else language
    return str(locale or "zh_CN")


__all__ = [
    "PresentationMarkdownRequestPolicy",
    "build_presentation_markdown_request_policy",
]

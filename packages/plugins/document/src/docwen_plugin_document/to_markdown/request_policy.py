"""Request-owned policy projection for standard DOCX -> Markdown conversion."""

from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass, replace
from typing import TYPE_CHECKING, Any

from docwen_core.docx_parsing.format_features import (
    DocxMarkdownFormattingConfig,
    DocxMarkdownSyntaxConfig,
    StyleDetectorConfig,
    docx_markdown_formatting_config_from_conversion_config,
    docx_markdown_syntax_config_from_conversion_config,
    style_detector_config_from_document_config,
)
from docwen_core.export_semantics import (
    MarkdownExportSemantics,
    normalize_markdown_break_separator,
    normalize_table_merge_export_strategy,
)

if TYPE_CHECKING:
    from docwen_core.protocols.execution_context import PluginExecutionContext


@dataclass(frozen=True, slots=True)
class DocxMarkdownRequestPolicy:
    """Effective request-owned policy for one standard DOCX conversion."""

    formatting: DocxMarkdownFormattingConfig
    syntax: DocxMarkdownSyntaxConfig
    style_detector: StyleDetectorConfig | None
    export: MarkdownExportSemantics
    ocr_blockquote_title: str

    def resolve_export_modes(self) -> dict[str, str]:
        """Return export modes already frozen into this request policy."""
        return {
            "image_extraction_mode": self.export.image_extraction_mode,
            "ocr_placement_mode": self.export.ocr_placement_mode,
            "table_merge_export_strategy": self.export.table_merge_export_strategy,
        }

    @property
    def image_link_style(self) -> str:
        """Return the image-link style frozen into this request policy."""
        return self.export.image_link_style


def build_docx_markdown_request_policy(
    context: PluginExecutionContext,
    options: Mapping[str, Any],
) -> DocxMarkdownRequestPolicy:
    """Project one request snapshot plus explicit options into DOCX policy.

    ``ConversionRequest.config_snapshot`` is authoritative, including an empty
    or partial mapping. Missing values resolve to deterministic defaults rather
    than mutable process-wide state.
    """
    conversion = _section(context.config.get("conversion", {}))
    document = _section(context.config.get("document", {}))
    link = _section(context.config.get("link", {}))
    export_config = _section(context.config.get("export", {}))
    output = _section(context.config.get("output", {}))

    link_format = _section(link.get("format", {}))
    ocr_output = _section(conversion.get("ocr_output", {}))
    intermediate_files = _section(output.get("intermediate_files", {}))
    locale = _request_locale(context.config, options)
    export = MarkdownExportSemantics.from_config(
        link_format=link_format,
        ocr_output=ocr_output,
        export_cfg=export_config,
        conversion_cfg=conversion,
        intermediate_files_cfg=intermediate_files,
        locale=locale,
    )
    export = replace(
        export,
        table_merge_export_strategy=normalize_table_merge_export_strategy(
            _first_nonblank(
                document.get("to_md_table_merge_export_strategy"),
                export.table_merge_export_strategy,
            ),
            default_strategy=export.table_merge_export_strategy,
        ),
    )
    formatting = docx_markdown_formatting_config_from_conversion_config(conversion)
    syntax = docx_markdown_syntax_config_from_conversion_config(conversion)
    style_detector = style_detector_config_from_document_config(document)
    ocr_blockquote_title = context.ocr_blockquote_title.strip()

    image_extraction_mode = str(
        _option_nonblank(
            options,
            "image_mode",
            export.image_extraction_mode,
        )
    )
    ocr_placement_mode = str(
        _option_nonblank(
            options,
            "ocr_placement",
            export.ocr_placement_mode,
        )
    )
    if image_extraction_mode.strip().lower() == "base64":
        ocr_placement_mode = "main_md"

    export = replace(
        export,
        image_extraction_mode=image_extraction_mode,
        ocr_placement_mode=ocr_placement_mode,
        table_merge_export_strategy=normalize_table_merge_export_strategy(
            _option_value(
                options,
                "table_merge_strategy",
                export.table_merge_export_strategy,
            ),
            default_strategy=export.table_merge_export_strategy,
        ),
        image_link_style=str(
            _option_value(
                options,
                "image_link_style",
                export.image_link_style,
            )
        ),
        page_break_separator=normalize_markdown_break_separator(
            _option_value(options, "page_break_separator", export.page_break_separator),
            default=export.page_break_separator,
        ),
        section_break_separator=normalize_markdown_break_separator(
            _option_value(options, "section_break_separator", export.section_break_separator),
            default=export.section_break_separator,
        ),
        horizontal_rule_separator=normalize_markdown_break_separator(
            _option_value(
                options,
                "horizontal_rule_separator",
                export.horizontal_rule_separator,
            ),
            default=export.horizontal_rule_separator,
        ),
    )
    formatting = replace(
        formatting,
        preserve_formatting=bool(_option_value(options, "preserve_formatting", formatting.preserve_formatting)),
        preserve_heading_formatting=bool(
            _option_value(
                options,
                "preserve_heading_formatting",
                formatting.preserve_heading_formatting,
            )
        ),
        preserve_table_header_formatting=bool(
            _option_value(
                options,
                "preserve_table_header_formatting",
                formatting.preserve_table_header_formatting,
            )
        ),
    )
    requested_style = _style_detector_from_options(options)
    if requested_style is not None:
        style_detector = _merge_style_detector_config(style_detector, requested_style)

    return DocxMarkdownRequestPolicy(
        formatting=formatting,
        syntax=syntax,
        style_detector=style_detector,
        export=export,
        ocr_blockquote_title=ocr_blockquote_title,
    )


def _section(value: object) -> dict[str, Any]:
    if not isinstance(value, Mapping):
        return {}
    return {str(key): item for key, item in value.items()}


def _request_locale(config: object, options: Mapping[str, Any]) -> str:
    requested = options.get("locale")
    if isinstance(requested, str) and requested.strip():
        return requested.strip()
    getter = getattr(config, "get", None)
    if callable(getter):
        configured = getter("gui.language.locale", "zh_CN")
        if isinstance(configured, str) and configured.strip():
            return configured.strip()
    return "zh_CN"


def _option_value(options: Mapping[str, Any], key: str, default: object) -> object:
    if key not in options or options[key] is None:
        return default
    return options[key]


def _option_nonblank(options: Mapping[str, Any], key: str, default: object) -> object:
    value = _option_value(options, key, default)
    if isinstance(value, str) and not value.strip():
        return default
    return value


def _first_nonblank(*values: object) -> object:
    for value in values:
        if value is None:
            continue
        if isinstance(value, str) and not value.strip():
            continue
        return value
    return ""


def _option_strings(options: Mapping[str, Any], key: str) -> list[str]:
    value = options.get(key, ())
    if not isinstance(value, (list, tuple, set, frozenset)):
        return []
    items = (
        sorted(value, key=lambda item: (str(item).casefold(), str(item)))
        if isinstance(value, (set, frozenset))
        else value
    )
    return [str(item).strip() for item in items if str(item).strip()]


def _style_detector_from_options(options: Mapping[str, Any]) -> StyleDetectorConfig | None:
    code_aliases = _option_strings(options, "code_block_style_aliases")
    quote_aliases = _option_strings(options, "quote_style_aliases")
    quote_generic_by_key: dict[str, str] = {}
    for name in _option_strings(options, "quote_generic_names"):
        quote_generic_by_key.setdefault(name.casefold(), name)
    quote_generic = frozenset(quote_generic_by_key.values())
    if not code_aliases and not quote_aliases and not quote_generic:
        return None
    return StyleDetectorConfig(
        code_block_style_fragments=tuple(code_aliases),
        quote_style_patterns=tuple((alias, 1) for alias in quote_aliases),
        quote_generic_names=quote_generic,
    )


def _merge_style_detector_config(
    base: StyleDetectorConfig | None,
    requested: StyleDetectorConfig,
) -> StyleDetectorConfig:
    """Apply request style fields without resetting document-owned policy."""
    effective = base or StyleDetectorConfig()
    code_fragments_by_key: dict[str, str] = {}
    for fragment in (*effective.code_block_style_fragments, *requested.code_block_style_fragments):
        stripped = fragment.strip()
        if stripped:
            code_fragments_by_key.setdefault(stripped.casefold(), stripped)

    quote_patterns_by_key: dict[tuple[str, int], tuple[str, int]] = {}
    for fragment, level in (*effective.quote_style_patterns, *requested.quote_style_patterns):
        stripped = fragment.strip()
        if stripped:
            quote_patterns_by_key.setdefault((stripped.casefold(), level), (stripped, level))

    quote_generic_by_key: dict[str, str] = {}
    effective_generic = sorted(effective.quote_generic_names, key=lambda name: (name.casefold(), name))
    requested_generic = sorted(requested.quote_generic_names, key=lambda name: (name.casefold(), name))
    for name in (*effective_generic, *requested_generic):
        stripped = name.strip()
        if stripped:
            quote_generic_by_key.setdefault(stripped.casefold(), stripped)

    return replace(
        effective,
        code_block_style_fragments=tuple(code_fragments_by_key.values()),
        quote_style_patterns=tuple(quote_patterns_by_key.values()),
        quote_generic_names=frozenset(quote_generic_by_key.values()),
    )


__all__ = ["DocxMarkdownRequestPolicy", "build_docx_markdown_request_policy"]

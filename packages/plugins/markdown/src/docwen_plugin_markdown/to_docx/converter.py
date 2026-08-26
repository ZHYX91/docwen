"""MdToDocxConverter — converts Markdown to DOCX using template-driven pipeline.

Five-stage pipeline:
1. **YAML** — extract YAML front matter from markdown content
2. **Parse** — parse cleaned body with extended mistune parser
3. **Template** — resolve/load template, locate placeholders, extract body font
4. **Render & Fill** — render AST to paragraphs via ``MdToDocxRenderer``,
   then inject into template via ``fill_template``
5. **Post** — save DOCX, write notes + numbering parts
"""

from __future__ import annotations

import re
import secrets
import time
from collections.abc import Mapping
from pathlib import Path
from typing import TYPE_CHECKING, Any

from docwen_core.docx_semantics import DocxSemanticRenderer
from docwen_core.docx_semantics_v3 import (
    CaptionStyleBindingV3,
    DocxSemanticsV3Error,
    DocxSemanticsV3Session,
)
from docwen_core.export_semantics import LinkRuntimeConfig
from docwen_core.links import (
    DeclaredResourceResolver,
    bind_declared_markdown_images,
    process_markdown_links,
    reject_declared_input_link_lookups,
)
from docwen_core.models.artifact import (
    ARTIFACT_KIND_PRIMARY,
    ArtifactManifest,
)
from docwen_core.models.result import (
    ConversionDiagnostic,
    ConversionErrorInfo,
    ConversionMetrics,
    ConversionResult,
)
from docwen_core.models.semantic_document import (
    SemanticBibliographyFragment,
    SemanticDocumentValidationError,
)
from docwen_core.text.heading_merge import (
    DEFAULT_HEADING_MERGE_PUNCTUATION,
    normalize_heading_merge_punctuation,
)
from docwen_core.text.heading_numbering import (
    NumberingSchemeResolutionError,
    resolve_heading_numbering_scheme,
)
from docwen_plugin_markdown.ast_transforms import (
    annotate_ast_with_hr_attachments,
    annotate_ast_with_merges,
)
from docwen_plugin_markdown.common_utils import (
    add_md_numbering,
    read_input_markdown,
    remove_md_numbering,
)
from docwen_plugin_markdown.document_semantics import analyze_document_semantics
from docwen_plugin_markdown.field_registry import (
    collect_placeholder_rules,
    collect_special_placeholder_handlers,
    run_yaml_processors,
)
from docwen_plugin_markdown.mistune_extensions import parse_markdown_text
from docwen_plugin_markdown.preprocessor import (
    detect_heading_merges,
    detect_hr_attachments,
    handle_setext_headings,
    materialize_image_placeholders,
    normalize_html_tags,
)
from docwen_plugin_markdown.renderer import MdToDocxRenderer
from docwen_plugin_markdown.resolved_conversion_v4 import claims_resolved_v4_inputs
from docwen_plugin_markdown.runtime_semantics_v3 import (
    RuntimeSemanticsV3Unsupported,
    apply_runtime_semantics_v3,
    prepare_runtime_semantics_v3,
)
from docwen_plugin_markdown.template_filler import fill_template
from docwen_plugin_markdown.template_utils import (
    TemplatePackageError,
    extract_body_font,
    extract_body_paragraph_format,
    extract_body_style,
    find_body_placeholder,
    resolve_template,
    scan_placeholders,
)
from docwen_plugin_markdown.to_docx.bibliography import (
    BibliographyConversionError,
    load_bibliography_resource,
    prepare_bibliography_anchor,
    validate_bibliography_placement,
    without_reserved_bibliography_placeholder,
)
from docwen_plugin_markdown.to_docx.managed_styles import (
    ManagedStyleCompletionError,
    complete_managed_styles,
    validate_managed_style_package,
)
from docwen_plugin_markdown.to_docx.notes import (
    NoteWritebackError,
    extract_notes_from_ast,
    normalize_note_syntax,
    prepare_note_context_for_document,
    write_notes_to_docx,
)
from docwen_plugin_markdown.yaml_processor import (
    TITLE_PLACEHOLDER_ALIASES,
    ensure_title_fallback,
    extract_yaml_front_matter,
)

if TYPE_CHECKING:
    from docx.document import Document

    from docwen_core.protocols.execution_context import DocumentStyleConverterContext

# ── Media types ─────────────────────────────────────────────────────────
MEDIA_TYPE_DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

# ── Built-in table style names ─────────────────────────────────────────
# These are the Word built-in style IDs. Localised display names in the
# Word UI may differ, but python-docx resolves style IDs case-insensitively.
_BUILTIN_STYLE_NAME = "Table Grid"
_MISSING = object()


def _config_value(config: object, key: str, default: object = None) -> object:
    try:
        return config.get(key, default)  # type: ignore[attr-defined]
    except Exception:
        return default


def _request_link_config(config: object) -> LinkRuntimeConfig:
    """Project this conversion request's immutable Link settings."""
    raw = _config_value(config, "link", {})
    if not isinstance(raw, Mapping):
        return LinkRuntimeConfig()
    return LinkRuntimeConfig.from_config(dict(raw))


def _request_yaml_list_separator(config: object) -> str:
    """Resolve the exact YAML list separator from this request snapshot."""
    raw = _config_value(config, "conversion.md_to_docx.list_separator", _MISSING)
    if raw is _MISSING or raw is None:
        return "、"
    return str(raw)


def _request_heading_merge_punctuation(options: dict[str, object], config: object) -> frozenset[str]:
    """Resolve the editable punctuation string from this request snapshot."""

    raw = options.get("heading_merge_punctuation", _MISSING)
    if raw is _MISSING or raw is None:
        raw = _config_value(
            config,
            "conversion.md_to_docx.heading_merge_punctuation",
            DEFAULT_HEADING_MERGE_PUNCTUATION,
        )
    return normalize_heading_merge_punctuation(raw)


def _option_or_config(
    options: dict[str, object],
    option_key: str,
    config: object,
    config_key: str,
    default: str,
    *,
    allowed: set[str] | None = None,
) -> str:
    raw = options.get(option_key, _MISSING)
    if raw is _MISSING or raw is None or str(raw).strip() == "":
        raw = _config_value(config, config_key, default)
    value = str(raw or default).strip().lower()
    if allowed is not None and value not in allowed:
        return default
    return value


def _resolve_body_formatting_mode(options: dict[str, object], config: object) -> str:
    raw = options.get("formatting_mode", _MISSING)
    if raw is _MISSING or raw is None or str(raw).strip() == "":
        raw = _config_value(config, "conversion.md_to_docx.formatting_mode", "full")
    value = str(raw or "full").strip().lower()
    if value in {"full", "minimal", "keep"}:
        return value
    if value == "remove":
        return "minimal"
    if value == "apply":
        return "full"
    return "full"


def _resolve_horizontal_rule_actions(config: object) -> dict[str, str]:
    enabled = _config_value(config, "conversion.horizontal_rule.enabled", True)
    if enabled is False or str(enabled).strip().lower() == "false":
        return {
            "dash": "ignore",
            "asterisk": "ignore",
            "underscore": "ignore",
        }

    allowed = {
        "page_break",
        "section_break",
        "horizontal_rule_1",
        "horizontal_rule_2",
        "horizontal_rule_3",
        "ignore",
    }

    def read(marker_key: str, default: str) -> str:
        raw = _config_value(config, f"conversion.horizontal_rule.md_to_docx.{marker_key}", default)
        value = str(raw or default).strip()
        return value if value in allowed else default

    return {
        "dash": read("dash", "page_break"),
        "asterisk": read("asterisk", "section_break"),
        "underscore": read("underscore", "horizontal_rule_1"),
    }


def _resolve_locale(gui_config: object) -> str:
    if isinstance(gui_config, dict):
        language = gui_config.get("language", {})
        if isinstance(language, dict):
            locale = language.get("locale")
            if isinstance(locale, str) and locale:
                return locale
    return "zh_CN"


def _resolve_table_style_name(
    mode: str,
    builtin_key: str,
    custom_name: str,
) -> str:
    """Resolve a table style name from config-driven options.

    Args:
        mode: ``"builtin"`` or ``"custom"``.
        builtin_key: ``"three_line_table"`` or ``"table_grid"``.
        custom_name: User-specified style name (used when mode is ``"custom"``).

    Returns:
        A Word style name string.
    """
    if mode == "custom" and custom_name:
        return custom_name
    if _normalize_table_style_key(builtin_key) == "three_line_table":
        return "Three Line Table"
    # Default / fallback to builtin
    return _BUILTIN_STYLE_NAME


def _normalize_table_style_key(builtin_key: str) -> str:
    value = str(builtin_key or "").strip().lower()
    if value in {"three_line_table", "table_grid"}:
        return value
    return "three_line_table"


def _markdown_image_alt_texts(markdown: str) -> list[str]:
    """Capture authored image alt text before link materialization."""

    return [
        "".join(str(child.get("raw", "") or child.get("text", "")) for child in image.get("children", []))
        for image in _iter_ast_images(parse_markdown_text(markdown, auto_link_bare_url=False))
    ]


def _iter_ast_images(nodes: list[dict[str, Any]]):
    for node in nodes:
        if node.get("type") == "image":
            yield node
        children = node.get("children")
        if isinstance(children, list):
            yield from _iter_ast_images(children)


def _restore_markdown_image_alt_texts(ast: list[dict[str, Any]], alt_texts: list[str]) -> None:
    """Restore alts lost by path-only image placeholder materialization."""

    for image, alt_text in zip(_iter_ast_images(ast), alt_texts, strict=False):
        image["children"] = [{"type": "text", "raw": alt_text}]


def _resolve_quote_style_levels(config: object) -> dict[int, int]:
    """Resolve configured ``quote_N`` style keys to semantic levels.

    Display names stay owned by the localized DOCX template.  The document
    config only maps Markdown depth to stable semantic keys, so the plugin can
    consume that mapping without importing runtime i18n resources.
    """
    result: dict[int, int] = {}
    for depth in range(1, 10):
        default = f"quote_{depth}"
        raw = _config_value(
            config,
            f"document.style.quote.md_to_docx.level_{depth}_style_key",
            default,
        )
        match = re.fullmatch(r"quote_([1-9])", str(raw or "").strip().lower())
        result[depth] = int(match.group(1)) if match else depth
    return result


def _resolve_template_style_keys(config: object) -> dict[str, str]:
    """Consume stable code/formula style keys without localized names."""
    paths = {
        "inline_code": "document.style.code.md_to_docx.inline_code_style_key",
        "code_block": "document.style.code.md_to_docx.code_block_style_key",
        "inline_formula": "document.style.formula.md_to_docx.inline_formula_style_key",
        "formula_block": "document.style.formula.md_to_docx.formula_block_style_key",
    }
    return {
        semantic: str(_config_value(config, path, semantic) or semantic).strip().lower()
        for semantic, path in paths.items()
    }


def _inject_docx_metadata(doc: Document, yaml_dict: dict[str, object]) -> None:
    """Write YAML title/subject through python-docx core properties.

    Mutating the loaded template before its first save preserves the template's
    namespace map.  Rewriting ``docProps/core.xml`` with ``ElementTree`` after
    save can rename the ``dcterms`` element prefix while leaving lexical QName
    values such as ``xsi:type=\"dcterms:W3CDTF\"`` unchanged.  Microsoft Word
    treats that undeclared QName prefix as a corrupt DOCX even though WPS opens
    it permissively.
    """
    title = (
        yaml_dict.get("标题")
        or yaml_dict.get("title")
        or next(
            (yaml_dict.get(key) for key in sorted(TITLE_PLACEHOLDER_ALIASES) if yaml_dict.get(key)),
            "",
        )
    )
    if isinstance(title, list):
        title = title[0] if title else ""
    title = str(title).strip()

    subtitle = yaml_dict.get("副标题") or yaml_dict.get("subtitle") or ""
    if isinstance(subtitle, list):
        subtitle = subtitle[0] if subtitle else ""
    subtitle = str(subtitle).strip()

    if not title and not subtitle:
        return

    if title:
        doc.core_properties.title = title
    if subtitle:
        doc.core_properties.subject = subtitle


def _numbering_resolution_failure(
    task_id: str,
    started_at: float,
    error: NumberingSchemeResolutionError,
) -> ConversionResult:
    return ConversionResult(
        task_id=task_id,
        success=False,
        error=ConversionErrorInfo(
            error_type=error.error_type,
            message=str(error),
            diagnostic_code=error.diagnostic_code,
        ),
        diagnostics=[
            ConversionDiagnostic(
                level="error",
                message=str(error),
                code=error.diagnostic_code,
            )
        ],
        metrics=ConversionMetrics(duration_ms=(time.monotonic() - started_at) * 1000.0),
    )


def _bibliography_failure(
    task_id: str,
    started_at: float,
    error: BibliographyConversionError,
) -> ConversionResult:
    return ConversionResult(
        task_id=task_id,
        success=False,
        error=ConversionErrorInfo(
            error_type=error.error_type,
            message=str(error),
            diagnostic_code=error.diagnostic_code,
        ),
        diagnostics=[
            ConversionDiagnostic(
                level="error",
                message=str(error),
                code=error.diagnostic_code,
            )
        ],
        metrics=ConversionMetrics(duration_ms=(time.monotonic() - started_at) * 1000.0),
    )


def _note_failure(
    task_id: str,
    started_at: float,
    error: NoteWritebackError,
    *,
    output_path: str | None = None,
) -> ConversionResult:
    if output_path is not None:
        Path(output_path).unlink(missing_ok=True)
    return ConversionResult(
        task_id=task_id,
        success=False,
        error=ConversionErrorInfo(
            error_type=error.error_type,
            message=str(error),
            diagnostic_code=error.diagnostic_code,
        ),
        diagnostics=[
            ConversionDiagnostic(
                level="error",
                message=str(error),
                code=error.diagnostic_code,
            )
        ],
        metrics=ConversionMetrics(duration_ms=(time.monotonic() - started_at) * 1000.0),
    )


def _semantic_v3_diagnostic(item: dict[str, Any]) -> ConversionDiagnostic:
    payload = dict(item)
    payload["level"] = payload.pop("severity")
    return ConversionDiagnostic.from_dict(payload)


def _semantic_v3_failure(
    task_id: str,
    started_at: float,
    message: str,
    diagnostics: list[ConversionDiagnostic],
    *,
    input_bytes: int = 0,
    output_path: str | None = None,
) -> ConversionResult:
    if output_path is not None:
        Path(output_path).unlink(missing_ok=True)
    code = diagnostics[0].code if diagnostics else "MD2DOCX-SEMANTICS-V3-UNSUPPORTED"
    return ConversionResult(
        task_id=task_id,
        success=False,
        error=ConversionErrorInfo(
            error_type="invalid_document_semantics",
            message=message,
            diagnostic_code=code,
        ),
        diagnostics=diagnostics
        or [
            ConversionDiagnostic(
                level="error",
                message=message,
                code=code,
            )
        ],
        metrics=ConversionMetrics(
            duration_ms=(time.monotonic() - started_at) * 1000.0,
            input_bytes=input_bytes,
        ),
    )


def _semantic_v3_package_diagnostic(message: str) -> ConversionDiagnostic:
    return ConversionDiagnostic(
        level="error",
        message=message,
        code="MD2DOCX-SEMANTICS-V3-PACKAGE-PROOF",
    )


class MdToDocxConverter:
    """Convert Markdown to DOCX using the 5-stage template-driven pipeline.

    Stage 1 (YAML): Extract front matter via ``extract_yaml_front_matter()``.
    Stage 2 (Parse): Parse cleaned Markdown via ``parse_markdown_text()``.
    Stage 3 (Template): Load template, find placeholders, extract body font.
    Stage 4 (Render & Fill): Render AST to paragraphs, inject into template.
    Stage 5 (Post): Save DOCX, write notes + numbering parts, inject metadata.
    """

    def convert(self, context: DocumentStyleConverterContext) -> ConversionResult:
        t_start = time.monotonic()
        task_id = context.request.request_id
        cancellable = context.cancellation
        logger = context.logger
        progress = context.progress
        workspace = context.workspace
        options = context.request.options

        try:
            # ── 1. Cancellation guard ──────────────────────────────────
            cancellable.check()

            # ── 2. Read input ──────────────────────────────────────────
            input_path = workspace.input_path
            declared_inputs = workspace.input_resources()
            if claims_resolved_v4_inputs(declared_inputs):
                # The resolved-document port is an exact-two capability, not
                # a source-file compatibility mode. A partial role claim is
                # routed here as well and fails closed before the historical
                # reader, v3 parser, link resolver, or source-number controls
                # can observe either JSON resource.
                from docwen_plugin_markdown.to_docx.resolved_v4_route import (
                    convert_resolved_v4_to_docx,
                )

                return convert_resolved_v4_to_docx(context, started_at=t_start)
            try:
                bibliography_fragment = load_bibliography_resource(declared_inputs)
            except BibliographyConversionError as exc:
                return _bibliography_failure(task_id, t_start, exc)
            source_input = next((item for item in declared_inputs if item.input_role == "source"), None)
            declared_resource_resolver = None
            if source_input is not None and source_input.logical_path:
                declared_resource_resolver = DeclaredResourceResolver(
                    source_logical_path=source_input.logical_path,
                    resources={
                        item.logical_path: item.path for item in declared_inputs if item.input_role == "linked_resource"
                    },
                )
            progress.report_progress(5.0, "Reading Markdown input")
            content, input_bytes = read_input_markdown(input_path)

            # Analyze the exact accepted input before numbering, YAML
            # extraction, or generic link preprocessing can change source
            # coordinates. The inert projection is length preserving, so the
            # Machine evidence remains bound to the full authored input.
            semantic_input_id = str(
                source_input.metadata.get("machine_input_id", "source") if source_input is not None else "source"
            )
            try:
                semantic_v3_plan = prepare_runtime_semantics_v3(
                    content,
                    input_id=semantic_input_id,
                )
            except RuntimeSemanticsV3Unsupported as exc:
                return _semantic_v3_failure(
                    task_id,
                    t_start,
                    str(exc),
                    [],
                    input_bytes=input_bytes,
                )
            if semantic_v3_plan.analysis.has_errors:
                return _semantic_v3_failure(
                    task_id,
                    t_start,
                    "Markdown v3 source semantics are invalid.",
                    [_semantic_v3_diagnostic(item) for item in semantic_v3_plan.analysis.diagnostics],
                    input_bytes=input_bytes,
                )
            content = semantic_v3_plan.shielded_source

            # ── 3. Optionally remove/add heading numbering ────────────
            remove_num: bool = options.get("remove_numbering", False)
            cleanup_rules = getattr(context, "heading_cleanup_rules", ()) or ()
            if remove_num:
                progress.report_progress(10.0, "Removing heading numbering")
                content = remove_md_numbering(content, rules=cleanup_rules)

            add_num: bool = options.get("add_numbering", False)
            render_mode: str = options.get("heading_numbering_render_mode", "text")
            word_native_translation = None  # set if word_native mode is used
            approximate_warning: str | None = None

            if add_num:
                scheme: str = options.get("numbering_scheme", "")
                try:
                    if render_mode == "word_native":
                        # Word adds the resolved scheme; source text remains clean.
                        if not remove_num:
                            content = remove_md_numbering(content, rules=cleanup_rules)
                        from docwen_core.text.numbering_word_adapter import (
                            translate_scheme,
                        )

                        scheme_config = resolve_heading_numbering_scheme(
                            scheme,
                            context.numbering_registry,
                        )
                        translation = translate_scheme(scheme_config)

                        if translation.verdict == "unsupported":
                            # Return error — do NOT silently fall back to text
                            return ConversionResult(
                                task_id=task_id,
                                success=False,
                                error=ConversionErrorInfo(
                                    error_type="unsupported_numbering",
                                    message=(
                                        f"Scheme '{scheme}' is not compatible "
                                        f"with word_native output: "
                                        f"{translation.reason}"
                                    ),
                                    diagnostic_code="MD2DOCX-NUMBERING-UNSUPPORTED",
                                ),
                                diagnostics=[
                                    ConversionDiagnostic(
                                        level="error",
                                        message=(
                                            f"Scheme '{scheme}' cannot be "
                                            f"translated to Word native "
                                            f"numbering: {translation.reason}"
                                        ),
                                        code="MD2DOCX-NUMBERING-UNSUPPORTED",
                                    )
                                ],
                                metrics=ConversionMetrics(
                                    duration_ms=(time.monotonic() - t_start) * 1000.0,
                                ),
                            )
                        else:
                            # word_native mode: skip text concatenation,
                            # inject after save
                            word_native_translation = translation
                            if translation.verdict == "approximate":
                                approximate_warning = (
                                    f"Scheme '{scheme}' uses numbering styles "
                                    f"that Word cannot render natively for "
                                    f"levels 6-9. Those headings will appear "
                                    f"without numbering. "
                                    f"Details: {translation.reason}"
                                )
                            # Still strip existing numbering (already done
                            # above if remove_num). Do NOT call
                            # add_md_numbering — headings stay as pure text.
                    else:
                        content = add_md_numbering(
                            content,
                            scheme=scheme,
                            registry=context.numbering_registry,
                        )
                except NumberingSchemeResolutionError as exc:
                    return _numbering_resolution_failure(task_id, t_start, exc)

            # ── Stage 1: YAML extraction ───────────────────────────────
            progress.report_progress(15.0, "Extracting YAML front matter")
            yaml_dict, md_body = extract_yaml_front_matter(content)
            field_processors_config = context.config.get("field_processors", {})
            current_locale = _resolve_locale(context.config.get("gui", {}))
            run_yaml_processors(yaml_dict, field_processors_config, current_locale=current_locale)
            placeholder_rules = collect_placeholder_rules(field_processors_config, current_locale=current_locale)
            special_placeholder_handlers = collect_special_placeholder_handlers(
                field_processors_config,
                current_locale=current_locale,
            )

            # ── Stage 2: Preprocess & Parse ────────────────────────────
            progress.report_progress(20.0, "Preprocessing Markdown")

            source_image_alt_texts = _markdown_image_alt_texts(md_body)

            # Numbering may have changed the full shielded source. Use the
            # freshly extracted body rather than the plan's old YAML offset.
            link_source = md_body
            if declared_resource_resolver is not None:
                # Semantic references are inert in ``link_source`` and cannot
                # trigger a filesystem lookup. Ordinary WikiLinks remain
                # visible here and are rejected at the declared-input boundary.
                reject_declared_input_link_lookups(link_source)
                # Bind request-declared resources in the already shielded
                # projection.  Replacing the authored body here would discard
                # semantic markers and, conversely, processing the pre-bind
                # snapshot would fall back to the physical source directory.
                link_source = bind_declared_markdown_images(
                    link_source,
                    declared_resource_resolver,
                )

            # 2a. Apply the request-scoped link policy after YAML extraction.
            link_config = _request_link_config(context.config)
            image_scope = secrets.token_urlsafe(24)
            md_body = process_markdown_links(
                link_source,
                input_path,
                link_config=link_config,
                target_format="docx",
                temp_dir=str(workspace.staging_dir),
                image_scope=image_scope,
            )
            md_body = materialize_image_placeholders(
                md_body,
                image_scope=image_scope,
            )

            # 2b. Convert Setext headings to ATX format after embed expansion.
            md_body = handle_setext_headings(md_body)

            # 2c. Normalize HTML tags (<br> → hard break, etc.)
            md_body = normalize_html_tags(md_body)

            # 2d. Validate the cross-project footnote/endnote contract and
            # create a request-local parser projection.  The source Markdown
            # is never rewritten.
            try:
                md_body = normalize_note_syntax(md_body)
            except NoteWritebackError as exc:
                return _note_failure(task_id, t_start, exc)

            # 2e. Detect heading merge boundaries
            heading_merge_mode = _option_or_config(
                options,
                "heading_merge_mode",
                context.config,
                "conversion.md_to_docx.heading_merge_mode",
                "punct_required",
                allowed={"punct_required", "always", "never"},
            )
            merge_indices = detect_heading_merges(
                md_body,
                mode=heading_merge_mode,
                punctuation=_request_heading_merge_punctuation(options, context.config),
            )

            # 2f. Detect HR attachments
            hr_attachments = detect_hr_attachments(md_body)

            # 2g. Parse with extended mistune
            progress.report_progress(30.0, "Parsing Markdown")
            raw_ast = parse_markdown_text(md_body, auto_link_bare_url=False)
            _restore_markdown_image_alt_texts(raw_ast, source_image_alt_texts)
            try:
                raw_ast = apply_runtime_semantics_v3(raw_ast, semantic_v3_plan)
            except RuntimeSemanticsV3Unsupported as exc:
                return _semantic_v3_failure(
                    task_id,
                    t_start,
                    str(exc),
                    [],
                    input_bytes=input_bytes,
                )

            # 2h. Annotate AST with merge info
            annotate_ast_with_merges(raw_ast, merge_indices)

            # 2i. Annotate AST with HR attachment info
            annotate_ast_with_hr_attachments(raw_ast, hr_attachments, md_body)

            # 2j. Extract notes from AST
            cleaned_ast, note_ctx = extract_notes_from_ast(raw_ast)

            # 2k. Recognize the frozen document-semantics v1 slice.  Errors
            # are rejected before template resolution so an invalid semantic
            # document never produces a seemingly successful DOCX artifact.
            semantic_analysis = analyze_document_semantics(cleaned_ast, current_v3=True)
            if semantic_analysis.has_errors:
                return ConversionResult(
                    task_id=task_id,
                    success=False,
                    error=ConversionErrorInfo(
                        error_type="invalid_document_semantics",
                        message="Markdown document semantics are invalid.",
                        diagnostic_code="MD2DOCX-DOCUMENT-SEMANTICS-INVALID",
                    ),
                    diagnostics=[
                        ConversionDiagnostic(
                            level=item.level,
                            message=item.message,
                            code=item.code,
                        )
                        for item in semantic_analysis.diagnostics
                    ],
                    metrics=ConversionMetrics(
                        duration_ms=(time.monotonic() - t_start) * 1000.0,
                        input_bytes=input_bytes,
                    ),
                )

            # ── Stage 3: Template resolution ──────────────────────────
            progress.report_progress(40.0, "Resolving template")
            template_name: str | None = options.get("template_name")
            code_font = _option_or_config(
                options,
                "code_font",
                context.config,
                "conversion.code_detection.code_font",
                "Consolas",
            )
            code_background_color = _option_or_config(
                options,
                "code_background_color",
                context.config,
                "conversion.code_detection.code_background_color",
                "E7E6E6",
            )

            try:
                doc = resolve_template(template_name)
            except TemplatePackageError as exc:
                return ConversionResult(
                    task_id=task_id,
                    success=False,
                    error=ConversionErrorInfo(
                        error_type="invalid_input",
                        message=str(exc),
                        diagnostic_code=exc.diagnostic_code,
                    ),
                    diagnostics=[
                        ConversionDiagnostic(
                            level="error",
                            message=str(exc),
                            code=exc.diagnostic_code,
                        )
                    ],
                    metrics=ConversionMetrics(
                        duration_ms=(time.monotonic() - t_start) * 1000.0,
                        input_bytes=input_bytes,
                    ),
                )
            try:
                validate_bibliography_placement(doc, bibliography_fragment)
            except BibliographyConversionError as exc:
                return _bibliography_failure(task_id, t_start, exc)
            try:
                doc, managed_styles = complete_managed_styles(
                    doc,
                    context.document_style_catalog,
                    code_font=code_font,
                    code_background_color=code_background_color,
                )
            except ManagedStyleCompletionError as exc:
                return ConversionResult(
                    task_id=task_id,
                    success=False,
                    error=ConversionErrorInfo(
                        error_type=exc.error_type,
                        message=str(exc),
                        diagnostic_code=exc.diagnostic_code,
                    ),
                    diagnostics=[
                        ConversionDiagnostic(
                            level="error",
                            message=str(exc),
                            code=exc.diagnostic_code,
                        )
                    ],
                    metrics=ConversionMetrics(
                        duration_ms=(time.monotonic() - t_start) * 1000.0,
                        input_bytes=input_bytes,
                    ),
                )
            try:
                prepare_note_context_for_document(doc, note_ctx)
            except NoteWritebackError as exc:
                return _note_failure(task_id, t_start, exc)
            placeholder_para = find_body_placeholder(doc)
            placeholder_map = scan_placeholders(doc)
            try:
                bibliography_anchor = prepare_bibliography_anchor(
                    doc,
                    bibliography_fragment,
                    bibliography_style_id=managed_styles.style_id("bibliography"),
                )
            except BibliographyConversionError as exc:
                return _bibliography_failure(task_id, t_start, exc)
            placeholder_map = without_reserved_bibliography_placeholder(placeholder_map)
            ensure_title_fallback(
                yaml_dict,
                placeholder_names=placeholder_map,
                source_stem=Path(input_path).stem,
            )
            body_font = extract_body_font(doc)
            body_style = extract_body_style(doc)
            body_paragraph_format = extract_body_paragraph_format(doc)

            # Resolve formatting options
            heading_formatting_mode = _option_or_config(
                options,
                "heading_formatting_mode",
                context.config,
                "conversion.md_to_docx.heading_formatting_mode",
                "remove",
                allowed={"apply", "keep", "remove"},
            )
            table_header_formatting_mode = _option_or_config(
                options,
                "table_header_formatting_mode",
                context.config,
                "conversion.md_to_docx.table_header_formatting_mode",
                "remove",
                allowed={"apply", "keep", "remove"},
            )
            formatting_mode = _resolve_body_formatting_mode(options, context.config)
            table_style_mode = _option_or_config(
                options,
                "table_style_mode",
                context.config,
                "document.style.table.md_to_docx.table_style_mode",
                "builtin",
                allowed={"builtin", "custom"},
            )
            builtin_style_key = _option_or_config(
                options,
                "builtin_style_key",
                context.config,
                "document.style.table.md_to_docx.builtin_style_key",
                "three_line_table",
            )
            custom_style_name = _option_or_config(
                options,
                "custom_style_name",
                context.config,
                "document.style.table.md_to_docx.custom_style_name",
                "",
            )
            table_style_name = _resolve_table_style_name(table_style_mode, builtin_style_key, custom_style_name)
            table_style_key = _normalize_table_style_key(builtin_style_key) if table_style_mode == "builtin" else None
            quote_style_levels = _resolve_quote_style_levels(context.config)
            template_style_keys = _resolve_template_style_keys(context.config)

            # HR mapping: dash/asterisk/underscore → text separator
            hr_mapping = options.get("hr_mapping")
            hr_actions = _resolve_horizontal_rule_actions(context.config)

            # ── Stage 4: Render & Fill ────────────────────────────────
            cancellable.check()
            progress.report_progress(55.0, "Rendering AST to paragraphs")

            semantic_v3_session = DocxSemanticsV3Session(
                doc,
                source_sha256=semantic_v3_plan.source_sha256,
                caption_style_bindings=tuple(
                    CaptionStyleBindingV3(
                        semantic_key=semantic_key,
                        resolved_style_id=managed_styles.style_id(semantic_key),
                        visible_name=managed_styles.get(semantic_key).name or "",
                    )
                    for semantic_key in (
                        "figure_caption",
                        "table_caption",
                        "equation_caption",
                        "code_block_caption",
                    )
                ),
            )

            # Create renderer with explicit doc object (does NOT create its own Document)
            renderer = MdToDocxRenderer(
                doc=doc,
                body_font=body_font,
                body_style=body_style,
                body_paragraph_format=body_paragraph_format,
                code_font=code_font,
                code_bg_color=code_background_color,
                heading_formatting_mode=heading_formatting_mode,
                table_header_formatting_mode=table_header_formatting_mode,
                formatting_mode=formatting_mode,
                table_style_name=table_style_name,
                table_style_key=table_style_key,
                quote_style_levels=quote_style_levels,
                template_style_keys=template_style_keys,
                managed_styles=managed_styles,
                semantic_v3_session=semantic_v3_session,
                hr_mapping=hr_mapping,
                hr_actions=hr_actions,
                hr_attachments=hr_attachments,
                cancellation=cancellable,
                note_ctx=note_ctx,
                source_file_path=input_path,
                declared_resource_resolver=declared_resource_resolver,
            )
            try:
                paragraphs = renderer.render(semantic_analysis.ast)
            except DocxSemanticsV3Error as exc:
                return _semantic_v3_failure(
                    task_id,
                    t_start,
                    str(exc),
                    [],
                    input_bytes=input_bytes,
                )

            # Inject paragraphs into template + fill YAML placeholders
            progress.report_progress(70.0, "Filling template")
            fill_template(
                doc=doc,
                yaml_dict=yaml_dict,
                rendered_paragraphs=paragraphs,
                placeholder_para=placeholder_para,
                placeholder_map=placeholder_map,
                placeholder_rules=placeholder_rules,
                special_placeholder_handlers=special_placeholder_handlers,
                list_separator=_request_yaml_list_separator(context.config),
            )
            try:
                semantic_v3_session.finalize_document()
            except DocxSemanticsV3Error as exc:
                return _semantic_v3_failure(
                    task_id,
                    t_start,
                    str(exc),
                    [],
                    input_bytes=input_bytes,
                )
            if bibliography_anchor is not None:
                try:
                    DocxSemanticRenderer(doc).render_bibliography_fragment(
                        bibliography_fragment or SemanticBibliographyFragment(entries=()),
                        placeholder_anchor=bibliography_anchor,
                        fallback_style_id=managed_styles.style_id("bibliography"),
                        hyperlink_style_id=managed_styles.style_id("hyperlink"),
                    )
                except SemanticDocumentValidationError as exc:
                    message = exc.diagnostics[0].message if exc.diagnostics else str(exc)
                    return _bibliography_failure(
                        task_id,
                        t_start,
                        BibliographyConversionError(
                            "MD2DOCX-BIBLIOGRAPHY-OOXML-CONFLICT",
                            message,
                        ),
                    )

            # ── Stage 5: Post (save + notes + numbering) ──────────────
            cancellable.check()
            progress.report_progress(80.0, "Writing DOCX to staging")

            output_path = workspace.create_artifact_path(ARTIFACT_KIND_PRIMARY, ".docx")
            input_stem = Path(input_path).stem
            suggested_name = f"{input_stem}.docx"

            # Inject DOCX metadata (title/subject) from YAML
            if yaml_dict:
                _inject_docx_metadata(doc, yaml_dict)

            doc.save(output_path)

            # Write footnote/endnote body elements into the DOCX ZIP parts
            if note_ctx.has_notes:
                try:
                    write_notes_to_docx(output_path, note_ctx)
                except NoteWritebackError as exc:
                    return _note_failure(task_id, t_start, exc, output_path=output_path)

            # Write Word-native list numbering definitions
            if renderer.list_numbering.has_definitions:
                from docwen_plugin_markdown.to_docx.numbering import (
                    write_numbering_to_docx,
                )

                write_numbering_to_docx(output_path, renderer.list_numbering)

            # Write Word-native heading numbering definitions
            if word_native_translation is not None:
                from docwen_plugin_markdown.to_docx.heading_numbering import (
                    write_heading_numbering_to_docx,
                )

                write_heading_numbering_to_docx(
                    output_path,
                    word_native_translation,
                    heading_style_ids={level: managed_styles.style_id(f"heading_{level}") for level in range(1, 10)},
                )

            try:
                semantic_v3_session.write_package(output_path)
            except DocxSemanticsV3Error as exc:
                return _semantic_v3_failure(
                    task_id,
                    t_start,
                    str(exc),
                    [_semantic_v3_package_diagnostic(str(exc))],
                    input_bytes=input_bytes,
                    output_path=output_path,
                )

            try:
                validate_managed_style_package(
                    Path(output_path).read_bytes(),
                    context.document_style_catalog,
                    managed_styles,
                )
            except ManagedStyleCompletionError as exc:
                Path(output_path).unlink(missing_ok=True)
                return ConversionResult(
                    task_id=task_id,
                    success=False,
                    error=ConversionErrorInfo(
                        error_type=exc.error_type,
                        message=str(exc),
                        diagnostic_code=exc.diagnostic_code,
                    ),
                    diagnostics=[ConversionDiagnostic(level="error", message=str(exc), code=exc.diagnostic_code)],
                    metrics=ConversionMetrics(
                        duration_ms=(time.monotonic() - t_start) * 1000.0,
                        input_bytes=input_bytes,
                    ),
                )

            try:
                semantic_v3_session.prove_package(output_path)
            except DocxSemanticsV3Error as exc:
                return _semantic_v3_failure(
                    task_id,
                    t_start,
                    str(exc),
                    [_semantic_v3_package_diagnostic(str(exc))],
                    input_bytes=input_bytes,
                    output_path=output_path,
                )

            output_bytes = Path(output_path).stat().st_size

            # ── 6. Register artifact ───────────────────────────────────
            artifact = ArtifactManifest(
                artifact_id=f"{task_id}-docx",
                kind=ARTIFACT_KIND_PRIMARY,
                staging_path=output_path,
                suggested_name=suggested_name,
                media_type=MEDIA_TYPE_DOCX,
                is_primary=True,
                metadata={
                    "source_format": "markdown",
                    "target_format": "docx",
                },
            )

            # ── 7. Return success ──────────────────────────────────────
            elapsed_ms = (time.monotonic() - t_start) * 1000.0
            diagnostics = [
                ConversionDiagnostic(
                    level="info",
                    message="MD→DOCX conversion successful",
                    code="MD2DOCX-OK",
                )
            ]
            diagnostics[0:0] = [_semantic_v3_diagnostic(item) for item in semantic_v3_plan.analysis.diagnostics]
            diagnostics[0:0] = [
                ConversionDiagnostic(
                    level=item.level,
                    message=item.message,
                    code=item.code,
                )
                for item in semantic_analysis.diagnostics
            ]
            diagnostics[0:0] = [
                ConversionDiagnostic(
                    level="warning",
                    message=item.message,
                    code=item.code,
                    location=(
                        f"style:{item.semantic_key};requested:{item.requested_style_id};"
                        f"resolved:{item.resolved_style_id}"
                    ),
                )
                for item in managed_styles.conflicts
            ]
            if approximate_warning:
                diagnostics.insert(
                    0,
                    ConversionDiagnostic(
                        level="warning",
                        message=approximate_warning,
                        code="MD2DOCX-NUMBERING-APPROXIMATE",
                    ),
                )

            result = ConversionResult(
                task_id=task_id,
                success=True,
                artifacts=[artifact],
                metrics=ConversionMetrics(
                    duration_ms=elapsed_ms,
                    input_bytes=input_bytes,
                    output_bytes=output_bytes,
                ),
                diagnostics=diagnostics,
            )
            # All fallible result/diagnostic construction is complete before
            # the workspace gains ownership. Registration is the final state
            # mutation, so a later generic error cannot leak a failed artifact.
            progress.report_progress(100.0, "Done")
            logger.info(f"MD→DOCX complete: {input_path} → {suggested_name}")
            workspace.add_artifact(artifact)
            return result

        except Exception as exc:
            pending_output = locals().get("output_path")
            if isinstance(pending_output, str):
                Path(pending_output).unlink(missing_ok=True)
            elapsed_ms = (time.monotonic() - t_start) * 1000.0
            logger.error(f"MD→DOCX failed: {exc}")
            return ConversionResult(
                task_id=task_id,
                success=False,
                error=ConversionErrorInfo(
                    error_type="conversion_failed",
                    message=str(exc),
                    diagnostic_code="MD2DOCX-ERROR",
                ),
                metrics=ConversionMetrics(duration_ms=elapsed_ms),
                diagnostics=[
                    ConversionDiagnostic(
                        level="error",
                        message=f"MD→DOCX conversion failed: {exc}",
                        code="MD2DOCX-ERROR",
                    )
                ],
            )

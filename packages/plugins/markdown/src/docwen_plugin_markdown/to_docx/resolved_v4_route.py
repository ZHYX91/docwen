"""Closed exact-two resolved-document MD→DOCX route.

The source-file converter delegates here before reading its primary input.
This module consumes only the two Workspace-owned JSON resources, re-admits
their typed port, stages only embedded resources, and gives Core's resolved
numbering session exclusive ownership of v4 semantic/numbering OOXML.
"""

from __future__ import annotations

import os
import shutil
import tempfile
import time
from contextlib import suppress
from copy import deepcopy
from dataclasses import dataclass
from pathlib import Path
from typing import TYPE_CHECKING, Any

from docx.oxml.ns import qn

from docwen_core._docx_recovery_map import ResolvedV4RecoveryInput
from docwen_core.docx_resolved_numbering import (
    ResolvedNumberingDocxError,
    ResolvedNumberingDocxSession,
)
from docwen_core.docx_semantics import DocxSemanticRenderer
from docwen_core.docx_semantics_v3 import CaptionStyleBindingV3
from docwen_core.docx_styles import SHIPPED_STYLE_LOCALES
from docwen_core.models.artifact import ARTIFACT_KIND_PRIMARY, ArtifactManifest
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
from docwen_core.resolved_resource_staging import (
    ResolvedResourceStagingError,
    bind_resolved_document_resources,
)
from docwen_plugin_markdown.ast_transforms import (
    annotate_ast_with_hr_attachments,
    annotate_ast_with_merges,
)
from docwen_plugin_markdown.field_registry import (
    collect_placeholder_rules,
    collect_special_placeholder_handlers,
    run_yaml_processors,
)
from docwen_plugin_markdown.manifest import RESOLVED_V4_MD_TO_DOCX_OPTIONS_SCHEMA
from docwen_plugin_markdown.mistune_extensions import parse_markdown_text
from docwen_plugin_markdown.preprocessor import detect_heading_merges, detect_hr_attachments
from docwen_plugin_markdown.renderer import MdToDocxRenderer
from docwen_plugin_markdown.resolved_conversion_v4 import (
    ResolvedConversionV4Unsupported,
    compose_resolved_v4_markdown,
    load_resolved_v4_inputs,
    prove_resolved_v4_image_inventory,
)
from docwen_plugin_markdown.resolved_runtime_v4 import (
    ResolvedRuntimeV4Unsupported,
    apply_resolved_runtime_v4,
)
from docwen_plugin_markdown.resolved_source_carriers_v4 import apply_resolved_source_carriers_v4
from docwen_plugin_markdown.runtime_semantics_v3 import RuntimeSemanticsV3Unsupported
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
    prepare_bibliography_anchor,
    validate_bibliography_placement,
    without_reserved_bibliography_placeholder,
)
from docwen_plugin_markdown.to_docx.converter import (
    MEDIA_TYPE_DOCX,
    _inject_docx_metadata,
    _markdown_image_alt_texts,
    _normalize_table_style_key,
    _option_or_config,
    _request_heading_merge_punctuation,
    _request_yaml_list_separator,
    _resolve_body_formatting_mode,
    _resolve_horizontal_rule_actions,
    _resolve_quote_style_levels,
    _resolve_table_style_name,
    _resolve_template_style_keys,
    _restore_markdown_image_alt_texts,
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
    ensure_title_fallback,
    extract_yaml_front_matter,
)

if TYPE_CHECKING:
    from docwen_core.protocols.execution_context import DocumentStyleConverterContext

_ACTIVE_OPTIONS = frozenset(RESOLVED_V4_MD_TO_DOCX_OPTIONS_SCHEMA["properties"])
_CAPTION_STYLE_KEYS = (
    "figure_caption",
    "table_caption",
    "equation_caption",
    "code_block_caption",
)


def convert_resolved_v4_to_docx(
    context: DocumentStyleConverterContext,
    *,
    started_at: float,
) -> ConversionResult:
    """Materialize one exact-two v4 port without entering legacy routes."""

    task_id = context.request.request_id
    workspace = context.workspace
    output_path: Path | None = None
    candidate_path: Path | None = None
    input_bytes = 0
    published = False
    state: dict[str, Any] = {"resource_root": None, "output_path": None}

    def failed_with_resource_cleanup(result: ConversionResult) -> ConversionResult:
        """Preserve the primary failure while reporting an unclean request scope."""

        resource_root = state["resource_root"]
        if resource_root is None:
            return result
        try:
            _remove_request_resource_root(Path(workspace.staging_dir), resource_root)
        except Exception as cleanup_exc:
            cleanup_fact = f"request-private resource cleanup failed ({type(cleanup_exc).__name__})"
            result.diagnostics.append(
                ConversionDiagnostic(
                    level="error",
                    message=cleanup_fact,
                    code="MD2DOCX-RESOLVED-V4-CLEANUP-FAILED",
                )
            )
            if result.error is not None:
                result.error.message = f"{result.error.message}; {cleanup_fact}"
        else:
            state["resource_root"] = None
        return result

    try:
        _rendered = _render_resolved_v4_docx(
            context,
            started_at=started_at,
            task_id=task_id,
            workspace=workspace,
            state=state,
        )
        output_path = _rendered.output_path
        input_bytes = _rendered.input_bytes
        for artifact in _rendered.artifacts:
            workspace.add_artifact(artifact)
        state["resource_root"] = None
        published = True
        return _rendered.result
    except ResolvedConversionV4Unsupported as exc:
        return failed_with_resource_cleanup(_failure(task_id, started_at, exc.code, str(exc), input_bytes=input_bytes))
    except (
        ResolvedResourceStagingError,
        ResolvedRuntimeV4Unsupported,
        RuntimeSemanticsV3Unsupported,
    ) as exc:
        return failed_with_resource_cleanup(
            _failure(
                task_id,
                started_at,
                "docwen.resolved_document.invalid",
                str(exc),
                input_bytes=input_bytes,
            )
        )
    except TemplatePackageError as exc:
        return failed_with_resource_cleanup(
            _failure(task_id, started_at, exc.diagnostic_code, str(exc), input_bytes=input_bytes)
        )
    except (ManagedStyleCompletionError, BibliographyConversionError, NoteWritebackError) as exc:
        code = getattr(exc, "diagnostic_code", "MD2DOCX-RESOLVED-V4-OUTPUT-INVALID")
        return failed_with_resource_cleanup(_failure(task_id, started_at, code, str(exc), input_bytes=input_bytes))
    except SemanticDocumentValidationError as exc:
        message = exc.diagnostics[0].message if exc.diagnostics else str(exc)
        return failed_with_resource_cleanup(
            _failure(
                task_id,
                started_at,
                "MD2DOCX-BIBLIOGRAPHY-OOXML-CONFLICT",
                message,
                input_bytes=input_bytes,
            )
        )
    except ResolvedNumberingDocxError as exc:
        return failed_with_resource_cleanup(
            _failure(
                task_id,
                started_at,
                "MD2DOCX-RESOLVED-V4-PACKAGE-PROOF",
                str(exc),
                input_bytes=input_bytes,
            )
        )
    except Exception as exc:
        return failed_with_resource_cleanup(
            _failure(
                task_id,
                started_at,
                "MD2DOCX-RESOLVED-V4-ERROR",
                str(exc),
                input_bytes=input_bytes,
            )
        )
    finally:
        if candidate_path is not None:
            candidate_path.unlink(missing_ok=True)
        if output_path is not None and not published:
            output_path.unlink(missing_ok=True)
        if state["output_path"] is not None and not published:
            Path(state["output_path"]).unlink(missing_ok=True)
        resource_root = state["resource_root"]
        if resource_root is not None:
            # The failed result already records this cleanup failure.  Retry
            # once for transient filesystem conditions without replacing the
            # primary structured error during Python's return-finalization.
            with suppress(Exception):
                _remove_request_resource_root(Path(workspace.staging_dir), resource_root)


@dataclass(frozen=True, slots=True)
class _ResolvedV4RenderedOutput:
    output_path: Path
    input_bytes: int
    artifacts: tuple[ArtifactManifest, ...]
    result: ConversionResult


def _render_resolved_v4_docx(
    context: DocumentStyleConverterContext,
    *,
    started_at: float,
    task_id: str,
    workspace: Any,
    state: dict[str, Any],
) -> _ResolvedV4RenderedOutput:
    """Render, write, and prove one exact-two v4 DOCX package."""

    context.cancellation.check()
    options = dict(context.request.options)
    style_catalog = context.document_style_catalog
    locale, template_name, heading_merge_mode = _validated_resolved_v4_options(options, style_catalog)

    context.progress.report_progress(5.0, "Re-admitting resolved v4 inputs")
    prepared = load_resolved_v4_inputs(workspace)
    input_bytes = prepared.neutral_document_path.stat().st_size + prepared.numbering_export_plan_path.stat().st_size

    proposed_resource_root = Path(workspace.staging_dir).resolve() / "resolved-v4-resources"
    if proposed_resource_root.exists() or proposed_resource_root.is_symlink():
        raise ResolvedResourceStagingError("request resource directory is not fresh")
    # Register ownership only after proving the fixed target was absent,
    # but before the binder can create it.  A binder failure after a
    # partial write is therefore still covered by the request cleanup.
    state["resource_root"] = proposed_resource_root
    resource_root = proposed_resource_root
    resources = bind_resolved_document_resources(prepared.port.document, proposed_resource_root)
    projection = compose_resolved_v4_markdown(prepared, resources)

    context.progress.report_progress(15.0, "Extracting YAML front matter")
    yaml_dict, md_body = extract_yaml_front_matter(projection.markdown)
    # The exact-two port owns a different input envelope, but the authored
    # Markdown inside that envelope still obeys the same frozen note syntax as
    # the source-file route.  Normalize only the request-local parser
    # projection so canonical endnotes retain their domain through Mistune.
    md_body = normalize_note_syntax(md_body)
    field_processors_config = context.config.get("field_processors", {})
    run_yaml_processors(yaml_dict, field_processors_config, current_locale=locale)
    placeholder_rules = collect_placeholder_rules(field_processors_config, current_locale=locale)
    special_placeholder_handlers = collect_special_placeholder_handlers(
        field_processors_config,
        current_locale=locale,
    )

    context.progress.report_progress(25.0, "Parsing authenticated Markdown projection")
    source_image_alt_texts = _markdown_image_alt_texts(md_body)
    merge_indices = detect_heading_merges(
        md_body,
        mode=heading_merge_mode,
        punctuation=_request_heading_merge_punctuation({}, context.config),
    )
    hr_attachments = detect_hr_attachments(md_body)
    raw_ast = parse_markdown_text(md_body, auto_link_bare_url=False)
    _restore_markdown_image_alt_texts(raw_ast, source_image_alt_texts)
    raw_ast = apply_resolved_runtime_v4(raw_ast, projection.runtime_plan)
    raw_ast = apply_resolved_source_carriers_v4(raw_ast, projection.source_carrier_plan)
    prove_resolved_v4_image_inventory(raw_ast, projection)
    annotate_ast_with_merges(raw_ast, merge_indices)
    annotate_ast_with_hr_attachments(raw_ast, hr_attachments, md_body)
    render_ast, note_ctx = extract_notes_from_ast(raw_ast)

    context.progress.report_progress(40.0, "Resolving template and managed styles")
    doc = resolve_template(template_name or None)
    bibliography_fragment = resources.bibliography
    validate_bibliography_placement(doc, bibliography_fragment)
    code_font = _option_or_config(
        {},
        "code_font",
        context.config,
        "conversion.code_detection.code_font",
        "Consolas",
    )
    code_background_color = _option_or_config(
        {},
        "code_background_color",
        context.config,
        "conversion.code_detection.code_background_color",
        "E7E6E6",
    )
    doc, managed_styles = complete_managed_styles(
        doc,
        style_catalog,
        code_font=code_font,
        code_background_color=code_background_color,
    )
    prepare_note_context_for_document(doc, note_ctx)
    placeholder_para = find_body_placeholder(doc)
    placeholder_map = scan_placeholders(doc)
    bibliography_anchor = prepare_bibliography_anchor(
        doc,
        bibliography_fragment,
        bibliography_style_id=managed_styles.style_id("bibliography"),
    )
    placeholder_map = without_reserved_bibliography_placeholder(placeholder_map)
    ensure_title_fallback(
        yaml_dict,
        placeholder_names=placeholder_map,
        source_stem=prepared.port.input_id,
    )

    session = ResolvedNumberingDocxSession(
        doc,
        prepared.port,
        heading_style_ids={level: managed_styles.style_id(f"heading_{level}") for level in range(1, 10)},
        heading_style_names={
            f"heading_{level}": managed_styles.get(f"heading_{level}").name or "" for level in range(1, 10)
        },
        caption_style_bindings=tuple(
            CaptionStyleBindingV3(
                semantic_key=semantic_key,
                resolved_style_id=managed_styles.style_id(semantic_key),
                visible_name=managed_styles.get(semantic_key).name or "",
            )
            for semantic_key in _CAPTION_STYLE_KEYS
        ),
        recovery_input=ResolvedV4RecoveryInput(
            neutral_raw=prepared.neutral_document_path.read_bytes(),
            plan_raw=prepared.numbering_export_plan_path.read_bytes(),
            authored_source=prepared.port.document.authored_markdown.encode("utf-8"),
            neutral_name="neutral-document.json",
            plan_name="numbering-export-plan.json",
            authored_name="authored-source.md",
            bibliography_owner="",
            bibliography_placeholder="",
            bibliography_media_type="",
        ),
    )

    renderer = _renderer(
        context,
        doc,
        managed_styles,
        session,
        projection.expected_image_urls,
        note_ctx,
        hr_attachments,
        code_font=str(code_font),
        code_background_color=str(code_background_color),
    )
    context.progress.report_progress(55.0, "Rendering resolved v4 AST")
    rendered = renderer.render(render_ast)
    fill_template(
        doc=doc,
        yaml_dict=yaml_dict,
        rendered_paragraphs=rendered,
        placeholder_para=placeholder_para,
        placeholder_map=placeholder_map,
        placeholder_rules=placeholder_rules,
        special_placeholder_handlers=special_placeholder_handlers,
        list_separator=_request_yaml_list_separator(context.config),
    )
    if bibliography_anchor is not None:
        DocxSemanticRenderer(doc).render_bibliography_fragment(
            bibliography_fragment or SemanticBibliographyFragment(entries=()),
            placeholder_anchor=bibliography_anchor,
            fallback_style_id=managed_styles.style_id("bibliography"),
            hyperlink_style_id=managed_styles.style_id("hyperlink"),
        )
    session.finalize_document()
    _install_list_numbering(doc, renderer.list_numbering)
    if yaml_dict:
        _inject_docx_metadata(doc, yaml_dict)

    context.cancellation.check()
    context.progress.report_progress(80.0, "Writing and proving resolved v4 DOCX")
    output_path = Path(workspace.create_artifact_path(ARTIFACT_KIND_PRIMARY, ".docx"))
    state["output_path"] = str(output_path)
    descriptor, candidate_name = tempfile.mkstemp(
        prefix=f".{output_path.name}.",
        suffix=".candidate",
        dir=output_path.parent,
    )
    os.close(descriptor)
    candidate_path = Path(candidate_name)
    candidate_path.unlink()

    def write_note_parts_before_recovery(path: Path) -> None:
        if note_ctx.has_notes:
            write_notes_to_docx(str(path), note_ctx)

    session.write_package(
        candidate_path,
        pre_recovery_package_transform=write_note_parts_before_recovery,
    )
    validate_managed_style_package(candidate_path.read_bytes(), style_catalog, managed_styles)
    session.prove_package(candidate_path)
    os.replace(candidate_path, output_path)
    _remove_request_resource_root(Path(workspace.staging_dir), resource_root)
    state["resource_root"] = None

    artifact = ArtifactManifest(
        artifact_id=f"{task_id}-docx",
        kind=ARTIFACT_KIND_PRIMARY,
        staging_path=str(output_path),
        suggested_name=f"{Path(prepared.port.input_id).name}.docx",
        media_type=MEDIA_TYPE_DOCX,
        is_primary=True,
        metadata={
            "source_format": "markdown",
            "target_format": "docx",
            "resolved_numbering": "v4",
        },
    )

    diagnostics = [
        ConversionDiagnostic(
            level="warning",
            message=item.message,
            code=item.code,
            location=(
                f"style:{item.semantic_key};requested:{item.requested_style_id};resolved:{item.resolved_style_id}"
            ),
        )
        for item in managed_styles.conflicts
    ]
    diagnostics.append(
        ConversionDiagnostic(
            level="info",
            message="Resolved v4 MD→DOCX conversion successful",
            code="MD2DOCX-RESOLVED-V4-OK",
        )
    )
    result = ConversionResult(
        task_id=task_id,
        success=True,
        artifacts=[artifact],
        metrics=ConversionMetrics(
            duration_ms=(time.monotonic() - started_at) * 1000.0,
            input_bytes=input_bytes,
            output_bytes=output_path.stat().st_size,
        ),
        diagnostics=diagnostics,
    )
    context.progress.report_progress(100.0, "Done")
    return _ResolvedV4RenderedOutput(output_path, input_bytes, (artifact,), result)


def _validated_resolved_v4_options(options: dict[str, Any], style_catalog: Any) -> tuple[str, str, str]:
    unknown_options = sorted(set(options) - _ACTIVE_OPTIONS)
    if unknown_options:
        raise ResolvedConversionV4Unsupported(
            "MD2DOCX-RESOLVED-V4-OPTIONS-INVALID",
            "resolved-v4 MD→DOCX accepts only locale, template_name, and heading_merge_mode",
        )
    if style_catalog is None:
        raise ResolvedConversionV4Unsupported(
            "MD2DOCX-RESOLVED-V4-OPTIONS-INVALID",
            "resolved-v4 conversion requires a request-owned document style catalog",
        )

    locale = options.get("locale", "zh_CN")
    if type(locale) is not str or locale not in SHIPPED_STYLE_LOCALES:
        raise ResolvedConversionV4Unsupported(
            "MD2DOCX-RESOLVED-V4-OPTIONS-INVALID",
            "locale must be one of the shipped resolved-v4 style locales",
        )
    if style_catalog.locale != locale:
        raise ResolvedConversionV4Unsupported(
            "MD2DOCX-RESOLVED-V4-OPTIONS-INVALID",
            "locale must equal the request-owned document style catalog locale",
        )

    template_name = options.get("template_name", "")
    if type(template_name) is not str or template_name != template_name.strip():
        raise ResolvedConversionV4Unsupported(
            "MD2DOCX-RESOLVED-V4-OPTIONS-INVALID",
            "template_name must be a normalized string resource selection",
        )

    heading_merge_mode = options.get("heading_merge_mode", "punct_required")
    if type(heading_merge_mode) is not str or heading_merge_mode not in {"punct_required", "never", "always"}:
        raise ResolvedConversionV4Unsupported(
            "MD2DOCX-RESOLVED-V4-OPTIONS-INVALID",
            "heading_merge_mode is outside the closed resolved-v4 set",
        )
    return locale, template_name, heading_merge_mode


def _renderer(
    context: DocumentStyleConverterContext,
    doc: Any,
    managed_styles: Any,
    session: ResolvedNumberingDocxSession,
    image_urls: tuple[str, ...],
    note_ctx: Any,
    hr_attachments: set[int],
    *,
    code_font: str,
    code_background_color: str,
) -> MdToDocxRenderer:
    table_style_mode = _option_or_config(
        {},
        "table_style_mode",
        context.config,
        "document.style.table.md_to_docx.table_style_mode",
        "builtin",
        allowed={"builtin", "custom"},
    )
    builtin_style_key = _option_or_config(
        {},
        "builtin_style_key",
        context.config,
        "document.style.table.md_to_docx.builtin_style_key",
        "three_line_table",
    )
    custom_style_name = _option_or_config(
        {},
        "custom_style_name",
        context.config,
        "document.style.table.md_to_docx.custom_style_name",
        "",
    )
    table_style_name = _resolve_table_style_name(table_style_mode, builtin_style_key, custom_style_name)
    table_style_key = _normalize_table_style_key(builtin_style_key) if table_style_mode == "builtin" else None
    return MdToDocxRenderer(
        doc=doc,
        body_font=extract_body_font(doc),
        body_style=extract_body_style(doc),
        body_paragraph_format=extract_body_paragraph_format(doc),
        code_font=code_font,
        code_bg_color=code_background_color,
        heading_formatting_mode=_option_or_config(
            {},
            "heading_formatting_mode",
            context.config,
            "conversion.md_to_docx.heading_formatting_mode",
            "remove",
            allowed={"apply", "keep", "remove"},
        ),
        table_header_formatting_mode=_option_or_config(
            {},
            "table_header_formatting_mode",
            context.config,
            "conversion.md_to_docx.table_header_formatting_mode",
            "remove",
            allowed={"apply", "keep", "remove"},
        ),
        formatting_mode=_resolve_body_formatting_mode({}, context.config),
        table_style_name=table_style_name,
        table_style_key=table_style_key,
        quote_style_levels=_resolve_quote_style_levels(context.config),
        template_style_keys=_resolve_template_style_keys(context.config),
        managed_styles=managed_styles,
        semantic_v3_session=None,
        resolved_numbering_session=session,
        resolved_image_urls=image_urls,
        hr_mapping=None,
        hr_actions=_resolve_horizontal_rule_actions(context.config),
        hr_attachments=hr_attachments,
        cancellation=context.cancellation,
        note_ctx=note_ctx,
        source_file_path=None,
        declared_resource_resolver=None,
    )


def _install_list_numbering(document: Any, numbering: Any) -> None:
    """Install ordinary Markdown list definitions before Core's atomic save."""

    if not numbering.has_definitions:
        return
    abstract_elements, num_elements = numbering._get_elements()
    root = document.part.numbering_part.element
    abstract_attr = qn("w:abstractNumId")
    num_attr = qn("w:numId")
    used_abstract = {item.get(abstract_attr) for item in root.findall(qn("w:abstractNum"))}
    used_num = {item.get(num_attr) for item in root.findall(qn("w:num"))}
    generated_abstract = {item.get(abstract_attr) for item in abstract_elements}
    generated_num = {item.get(num_attr) for item in num_elements}
    if used_abstract & generated_abstract or used_num & generated_num:
        raise ValueError("ordinary list numbering IDs collide with the resolved template")
    first_num = root.find(qn("w:num"))
    for element in abstract_elements:
        copy = deepcopy(element)
        if first_num is None:
            root.append(copy)
        else:
            first_num.addprevious(copy)
    for element in num_elements:
        root.append(deepcopy(element))


def _remove_request_resource_root(staging_dir: Path, resource_root: Path) -> None:
    """Remove only this route's fixed request-private staging directory."""

    staging = staging_dir.resolve()
    target = resource_root.resolve()
    if target.parent != staging or target.name != "resolved-v4-resources":
        raise RuntimeError("resolved-v4 resource cleanup target escaped its request staging directory")
    if not target.exists():
        return
    if target.is_symlink() or not target.is_dir():
        raise RuntimeError("resolved-v4 resource cleanup target changed type")
    shutil.rmtree(target)


def _failure(
    task_id: str,
    started_at: float,
    code: str,
    message: str,
    *,
    input_bytes: int,
) -> ConversionResult:
    return ConversionResult(
        task_id=task_id,
        success=False,
        error=ConversionErrorInfo(
            error_type="invalid_resolved_document",
            message=message,
            diagnostic_code=code,
        ),
        metrics=ConversionMetrics(
            duration_ms=(time.monotonic() - started_at) * 1000.0,
            input_bytes=input_bytes,
        ),
        diagnostics=[ConversionDiagnostic(level="error", message=message, code=code)],
    )


__all__ = ["convert_resolved_v4_to_docx"]

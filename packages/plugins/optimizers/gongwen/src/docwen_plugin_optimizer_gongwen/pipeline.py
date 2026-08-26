"""Public entry point: convert_docx_to_md_gongwen()."""

from __future__ import annotations

import re
from typing import TYPE_CHECKING, Any

from docwen_core.text.heading_numbering import (
    HeadingFormatter,
    resolve_heading_numbering_scheme,
)
from docwen_core.text.ocr import format_ocr_best_effort_warning

if TYPE_CHECKING:
    from docwen_core.export_semantics import MarkdownExportSemantics
    from docwen_core.protocols.execution_context import (
        CancellationTokenView,
        ProgressSink,
    )


_TABLE_STRUCTURAL_LABELS = frozenset(
    {
        "版记",
        "份号",
        "密级",
        "紧急程度",
        "发文字号",
        "签发人",
        "主送机关",
        "附件",
        "附件说明",
        "抄送",
        "抄送机关",
        "报送",
        "分送",
        "公开方式",
        "公开属性",
        "印发机关",
        "印发日期",
        "成文日期",
        "发文机关",
    }
)


def _exact_bool(options: dict[str, Any], key: str, default: bool) -> bool:
    value = options.get(key, default)
    if type(value) is not bool:
        raise ValueError(f"{key} must be a boolean")
    return value


def _exact_choice(options: dict[str, Any], key: str, default: str, choices: frozenset[str]) -> str:
    value = options.get(key, default)
    if not isinstance(value, str) or value not in choices:
        allowed = "/".join(sorted(choices))
        raise ValueError(f"{key} must be one of {allowed}")
    return value


def _is_table_structural_label(text: str) -> bool:
    compact = re.sub(r"\s", "", text)
    normalized = compact.rstrip("：:")
    if normalized in _TABLE_STRUCTURAL_LABELS:
        return True
    return any(
        compact.startswith(f"{label}{separator}") for label in _TABLE_STRUCTURAL_LABELS for separator in ("：", ":")
    )


def _prepare_table_rendering(features: list[Any], result: Any) -> None:
    """Keep mixed tables while suppressing recognition-only cell features.

    A table is structural only when it has no substantive body cell.  When a
    metadata cell originally owns the render anchor, transfer that anchor to a
    surviving body cell instead of dropping the whole table.
    """

    table_groups: dict[int, list[int]] = {}
    for feature in features:
        if feature.table_index is not None:
            table_groups.setdefault(feature.table_index, []).append(feature.index)

    skip_set = set(result.skip_indices)
    for indices in table_groups.values():
        anchor = next(
            (features[index] for index in indices if features[index].is_table_anchor),
            features[indices[0]],
        )
        structural_candidate_present = any(
            index in result.candidates and result.candidates[index].element_type != "body" for index in indices
        )
        substantive_body_indices = [
            index
            for index in indices
            if features[index].text.strip()
            and not _is_table_structural_label(features[index].text)
            and (index not in result.candidates or result.candidates[index].element_type == "body")
        ]

        if not substantive_body_indices:
            anchor.table_output_mode = (
                "structural_metadata"
                if structural_candidate_present
                or any(_is_table_structural_label(features[index].text) for index in indices)
                else "unrepresentable"
            )
            skip_set.update(indices)
            continue

        target_index = anchor.index if anchor.index in substantive_body_indices else substantive_body_indices[0]
        target = features[target_index]
        if target is not anchor:
            target.table_markdown = anchor.table_markdown
            target.table_fidelity_risks = anchor.table_fidelity_risks
            target.is_table_anchor = True
            anchor.table_markdown = ""
            anchor.is_table_anchor = False
            anchor.table_output_mode = ""
        target.table_output_mode = "rendered"
        skip_set.update(index for index in indices if index != target_index)
        skip_set.discard(target_index)

    result.skip_indices = sorted(skip_set)


def _report_ocr_best_effort(progress: ProgressSink | None, status: object, *, location: str) -> None:
    """Report one safe warning for a fallible OCR outcome."""
    if progress is None:
        return
    message = format_ocr_best_effort_warning(status)
    if message is None:
        return
    progress.report_diagnostic(
        "warning",
        message,
        code="OCR-BEST-EFFORT",
        location=location,
    )


def _attachment_line_text(feature: Any, *, remove_numbering: bool) -> str:
    """Return attachment text without silently losing a retained prefix."""
    if not remove_numbering and feature.heading_level > 0 and feature.heading_numbering_text:
        return feature.raw_text or f"{feature.heading_numbering_text}{feature.text}"
    return feature.text


def convert_docx_to_md_gongwen(
    doc: Any,  # python-docx Document
    input_path: str,
    options: dict | None = None,
    *,
    cancellation: CancellationTokenView | None = None,
    progress: ProgressSink | None = None,
    registry: Any = None,
    cleanup_rules: Any = (),
    export_semantics: MarkdownExportSemantics | None = None,
) -> dict:
    """Run full gongwen pipeline on a DOCX document.

    Args:
        doc: python-docx ``Document`` instance.
        input_path: Path of the source file.
        options: Conversion options dict.
        cancellation: Optional token for cooperative cancellation.
        progress: Optional sink for progress reporting.

    Returns dict with keys:
        success: bool
        yaml_info: dict — the 18 YAML fields
        markdown: str — main body Markdown with YAML frontmatter
        attachment_documents: list[AttachmentDocument] — typed attachment outputs
        stats: dict — paragraph counts
        metadata: dict — recognition metadata (confidence, findings)
    """
    if options is None:
        options = {}

    add_numbering = _exact_bool(options, "add_numbering", False)
    numbering_scheme = options.get("numbering_scheme", "")
    heading_formatter = None
    if add_numbering:
        scheme_config = resolve_heading_numbering_scheme(numbering_scheme, registry)
        heading_formatter = HeadingFormatter(scheme_config)

    from docwen_plugin_optimizer_gongwen.extraction.paragraph_reader import read_paragraphs
    from docwen_plugin_optimizer_gongwen.recognition.reevaluation import maybe_reevaluate
    from docwen_plugin_optimizer_gongwen.recognition.rounds import run_rounds
    from docwen_plugin_optimizer_gongwen.recognition.scorer import ElementScorer
    from docwen_plugin_optimizer_gongwen.recognition.yaml_builder import build_yaml
    from docwen_plugin_optimizer_gongwen.rendering.markdown_renderer import render
    from docwen_plugin_optimizer_gongwen.validation import get_confidence_summary, validate_result

    # ── 1.  Extraction ──────────────────────────────────────────────
    if progress:
        progress.report_progress(0.0, "Extracting paragraph features")
    output_dir = options.get("output_dir") if options else None
    features = read_paragraphs(
        doc,
        output_dir=output_dir,
        cleanup_rules=cleanup_rules,
        table_merge_strategy=_exact_choice(
            options,
            "table_merge_strategy",
            "fill",
            frozenset({"fill", "empty", "marker"}),
        ),
        diagnostic_sink=progress,
    )
    if cancellation:
        cancellation.check()

    # ── 1b. Optional OCR for embedded images ────────────────────────
    enable_ocr = _exact_bool(options, "to_md_enable_ocr", False)
    ocr_language = _exact_choice(
        options,
        "ocr_language",
        "auto",
        frozenset({"auto", "chinese", "chinese_cht", "english", "japanese", "korean", "latin", "cyrillic"}),
    )
    current_locale = options.get("locale", "zh_CN")
    if not isinstance(current_locale, str) or not current_locale:
        raise ValueError("locale must be a non-empty string")

    # ── 1c. Numbering options ──────────────────────────────────────
    remove_numbering = _exact_bool(options, "remove_numbering", True)
    keep_images = _exact_bool(options, "to_md_keep_images", True)
    image_mode = _exact_choice(
        options,
        "image_mode",
        "file",
        frozenset({"file", "base64", "embed", "omit"}),
    )
    if not keep_images:
        image_mode = "omit"
    image_link_style = _exact_choice(
        options,
        "image_link_style",
        "markdown_embed",
        frozenset({"wiki_embed", "wiki_link", "markdown_embed", "markdown_link"}),
    )

    if enable_ocr:
        if progress:
            progress.report_progress(10.0, "Running OCR on embedded images")
        from docwen_core.detection import detect_content_format
        from docwen_core.text.ocr import OcrOutcome, OcrStatus, run_ocr_outcome

        for pf in features:
            for img_path in pf.extracted_images:
                if img_path:
                    try:
                        outcome = run_ocr_outcome(
                            img_path,
                            source_format=detect_content_format(img_path).format,
                            ocr_language=ocr_language,
                            current_locale=current_locale,
                        )
                    except Exception as exc:
                        outcome = OcrOutcome(OcrStatus.RECOGNITION_FAILED, message=str(exc))
                    _report_ocr_best_effort(
                        progress,
                        outcome.status,
                        location=str(img_path),
                    )
                    if outcome.recognized_text and outcome.recognized_text.strip():
                        pf.image_ocr_texts[img_path] = outcome.recognized_text

    # ── 2.  Recognition (three rounds + re-evaluation) ──────────────
    if progress:
        progress.report_progress(20.0, "Running element recognition rounds")
    scorer = ElementScorer(diagnostic_sink=progress)
    result = run_rounds(scorer, features)

    if cancellation:
        cancellation.check()

    if progress:
        progress.report_progress(50.0, "Running validation re-evaluation")
    result = maybe_reevaluate(scorer, features, result)

    if cancellation:
        cancellation.check()

    # ── 3.  YAML building ───────────────────────────────────────────
    if progress:
        progress.report_progress(60.0, "Building YAML metadata")
    result = build_yaml(
        scorer,
        features,
        result,
        cleanup_rules=cleanup_rules,
    )
    result = validate_result(result)

    # Recognition consumes individual cells; Markdown consumes one table.
    # Metadata and body can coexist in the same table, so only wholly
    # structural tables are omitted from the body.
    _prepare_table_rendering(features, result)

    # ── 4.  Rendering ───────────────────────────────────────────────
    if progress:
        progress.report_progress(75.0, "Rendering Markdown output")
    from docwen_plugin_optimizer_gongwen.models import GongwenMetadata
    from docwen_plugin_optimizer_gongwen.rendering.attachment_renderer import render_attachment

    metadata = GongwenMetadata.from_dict(result.yaml_info)
    # Keep this list strictly one-to-one with ParagraphFeature.index.  Mixed
    # heading/body paragraphs are split only while rendering that same feature;
    # flattening them here shifts every later image, formula, skip and heading
    # lookup onto the wrong source paragraph.
    body_lines = [feature.text for feature in features]
    feature_map = {f.index: f for f in features}
    markdown = render(
        metadata,
        body_lines,
        skip_indices=result.skip_indices,
        feature_map=feature_map,
        remove_numbering=remove_numbering,
        heading_formatter=heading_formatter,
        image_mode=image_mode,
        image_link_style=image_link_style,
        export_semantics=export_semantics,
    )

    # ── 5.  Attachment rendering ─────────────────────────────────────
    if progress:
        progress.report_progress(85.0, "Rendering attachment content")
    attachment_documents = []
    # Attachment descriptions already live in YAML metadata.  A separate
    # attachment artifact is warranted only for actual attachment body
    # content; including header/following rows duplicated the same list.
    attachment_indices = [idx for idx, c in result.candidates.items() if c.element_type == "attachment_content"]
    if attachment_indices:
        attachment_lines = []
        for idx in sorted(attachment_indices):
            if idx >= len(features):
                continue
            line = _attachment_line_text(features[idx], remove_numbering=remove_numbering)
            if line.strip():
                attachment_lines.append(line)
        attachment_md = render_attachment(metadata, attachment_lines, input_path=input_path)
        if attachment_md.strip():
            from docwen_plugin_optimizer_gongwen.models import AttachmentDocument

            title = metadata.attachment[0] if len(metadata.attachment) == 1 else "附件汇总"
            attachment_documents.append(
                AttachmentDocument(
                    ordinal=1,
                    title=title or "附件",
                    markdown=attachment_md,
                    paragraph_indices=tuple(sorted(attachment_indices)),
                )
            )

    # ── 6.  Confidence summary ──────────────────────────────────────
    from docwen_plugin_optimizer_gongwen.recognition.signals import collect_structured_signals

    confidence = get_confidence_summary(result)
    structured_signals = collect_structured_signals(result, scorer, features)

    if progress:
        progress.report_progress(100.0, "Gongwen conversion complete")

    image_paths: list[str] = []
    if keep_images and image_mode in {"file", "embed"}:
        image_paths = list(dict.fromkeys(path for feature in features for path in feature.extracted_images if path))

    return {
        "success": True,
        "yaml_info": result.yaml_info,
        "markdown": markdown,
        "attachment_documents": attachment_documents,
        "stats": {
            "paragraphs": len(features),
            "scoring_rule_failure_count": len(scorer.rule_failures),
        },
        "image_paths": image_paths,
        "metadata": {
            "confidence": confidence,
            "findings": result.validation_finding_count,
            "missing_required": result.missing_required,
            "recognition_review_signals": structured_signals,
        },
    }

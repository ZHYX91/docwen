"""LayoutToMarkdownConverter — converts PDF (and other layout formats) to Markdown.

Uses pymupdf4llm for native text/image extraction.  When OCR is enabled each
physical page is independently rendered via PyMuPDF and emitted as one typed
OCR fragment; recognised page text is never duplicated in the primary
Markdown document.

When the source is HTML, ``preprocess_html_images`` (from the preprocess
layer) resolves ``<img>`` elements — including data URIs, remote URLs, and
local paths with ``<base href>`` resolution — before the HTML is converted
to Markdown.
"""

from __future__ import annotations

import contextlib
import logging
import os
import re
from pathlib import Path
from typing import TYPE_CHECKING

from docwen_core.detection import detect_content_format
from docwen_core.export_semantics import LinkRuntimeConfig
from docwen_core.formats import get_category, get_media_type
from docwen_core.paths import input_stem
from docwen_core.text.ocr import (
    OcrOutcome,
    OcrStatus,
    format_ocr_best_effort_warning,
    run_ocr_outcome,
)
from docwen_plugin_layout._common import file_size, new_artifact_id, request_source_format

logger = logging.getLogger(__name__)

if TYPE_CHECKING:
    from docwen_core.export_semantics import MarkdownExportSemantics
    from docwen_core.models.result import ConversionResult
    from docwen_core.protocols.execution_context import ConverterContext


def _detected_image_format(path: Path) -> str | None:
    try:
        detected_format = detect_content_format(str(path)).format
    except OSError:
        return None
    return detected_format if get_category(detected_format) == "image" else None


_WIKI_IMAGE_TARGET_RE = re.compile(r"!\[\[([^\]|#]+)(?:[|#][^\]]*)?\]\]")
_MARKDOWN_IMAGE_TARGET_RE = re.compile(r"!\[[^\]]*\]\(\s*(?:<([^>]+)>|([^\s)]+))")


def _image_target_spellings(image_file: Path) -> set[str]:
    """Return equivalent textual spellings for one owned extracted image."""
    spellings = {image_file.name, str(image_file).replace("\\", "/")}
    with contextlib.suppress(OSError):
        spellings.add(str(image_file.resolve(strict=True)).replace("\\", "/"))
    return spellings


def _referenced_extracted_images(
    markdown: str,
    *,
    images_prefix: str,
    image_files: set[Path],
) -> set[Path]:
    """Return request-created image files explicitly referenced by one page.

    This is a typed ownership observation over the page result, not a page
    inference from a generated filename.  Exact Markdown/wiki image targets
    are matched against the request-owned files that exist after that page's
    extraction call.  Callers deliberately exclude files present before the
    first page so a staging collision can never acquire page ownership.
    """
    targets = {match.group(1).strip().replace("\\", "/") for match in _WIKI_IMAGE_TARGET_RE.finditer(markdown)}
    targets.update(
        (match.group(1) or match.group(2) or "").strip().replace("\\", "/")
        for match in _MARKDOWN_IMAGE_TARGET_RE.finditer(markdown)
    )
    referenced: set[Path] = set()
    for image_file in image_files:
        candidates = _image_target_spellings(image_file) | {f"{images_prefix}{image_file.name}"}
        if targets & candidates:
            referenced.add(image_file)
    return referenced


def _ocr_page_outcomes(
    pdf_path: str,
    staging_dir: str,
    dpi: int = 200,
    *,
    ocr_language: str | None = None,
    current_locale: str = "zh_CN",
) -> list[OcrOutcome]:
    """Render every page of *pdf_path* to a PNG and run OCR on each.

    Returns one typed best-effort outcome per rendered page.
    """
    import fitz  # PyMuPDF

    results: list[OcrOutcome] = []
    doc = fitz.open(pdf_path, filetype="pdf")
    try:
        for page_num in range(len(doc)):
            page_img_path = os.path.join(staging_dir, f"_ocr_page_{page_num + 1}.png")
            try:
                page = doc[page_num]
                pix = page.get_pixmap(dpi=dpi)
                pix.save(page_img_path)
                outcome = run_ocr_outcome(
                    page_img_path,
                    source_format="png",
                    ocr_language=ocr_language,
                    current_locale=current_locale,
                )
            except Exception as exc:
                outcome = OcrOutcome(OcrStatus.RECOGNITION_FAILED, message=str(exc))
            finally:
                try:
                    Path(page_img_path).unlink(missing_ok=True)
                except OSError:
                    logger.warning("Unable to remove request-owned OCR page image %s", page_img_path)
            results.append(outcome)
    finally:
        doc.close()
    return results


def _rewrite_extracted_image_refs(
    md_text: str,
    *,
    images_prefix: str,
    image_files: list[Path],
    image_mode: str,
    image_link_style: str,
    export_semantics: MarkdownExportSemantics | None = None,
) -> tuple[str, dict[Path, str]]:
    """Rewrite pymupdf4llm image refs through the shared image Markdown helper."""
    from docwen_core.text.image_markdown import generate_image_markdown

    rewritten = md_text
    image_markdown_by_file: dict[Path, str] = {}
    for img_file in image_files:
        old_target = f"{images_prefix}{img_file.name}"
        image_path = img_file if image_mode == "base64" else img_file.name
        new_ref = generate_image_markdown(
            image_path=image_path,
            image_mode=image_mode,
            image_link_style=image_link_style,
            alt_text=img_file.stem,
            export_semantics=export_semantics,
        )
        targets = _image_target_spellings(img_file) | {old_target}
        candidates = {
            candidate
            for target in targets
            for candidate in (
                f"![[{target}]]",
                f"![]({target})",
                f"![](<{target}>)",
                f"![{img_file.stem}]({target})",
                f"![{img_file.name}]({target})",
            )
        }
        for candidate in candidates:
            rewritten = rewritten.replace(candidate, new_ref)
        image_markdown_by_file[img_file] = new_ref
        rewritten = rewritten.replace(old_target, img_file.name)
    return rewritten, image_markdown_by_file


def _rename_preprocessed_image_refs(
    md_text: str,
    *,
    image_files: list[Path],
    effective_input_path: str,
    original_input_path: str,
) -> tuple[str, list[Path]]:
    """Replace an internal preprocessed-PDF prefix with the source filename."""
    effective_name = Path(effective_input_path).name
    original_name = Path(original_input_path).name
    if effective_name == original_name:
        return md_text, image_files

    rewritten = md_text
    renamed_files: list[Path] = []
    for image_file in image_files:
        if not image_file.name.startswith(effective_name):
            renamed_files.append(image_file)
            continue

        target = image_file.with_name(original_name + image_file.name[len(effective_name) :])
        old_path = str(image_file)
        old_path_posix = image_file.as_posix()
        image_file.rename(target)
        rewritten = rewritten.replace(old_path, str(target))
        rewritten = rewritten.replace(old_path_posix, target.as_posix())
        rewritten = rewritten.replace(image_file.name, target.name)
        renamed_files.append(target)
    return rewritten, renamed_files


def _resolve_md_links(md_text: str, source_path: str) -> str:
    """Post-process extracted Markdown to resolve ``![[...]]`` wiki-embed
    links and handle non-embed ``[[link]]`` / ``[text](url)`` links via the
    shared ``process_markdown_links`` orchestrator.

    The call is a no-op when *md_text* contains no recognised link patterns.
    """
    from docwen_core.links import process_markdown_links

    return process_markdown_links(
        md_text,
        source_path,
        link_config=LinkRuntimeConfig(
            non_embed_wiki_mode="keep",
            non_embed_markdown_mode="keep",
            embed_wiki_image_mode="keep",
            embed_markdown_image_mode="keep",
            embed_md_file_mode="keep",
        ),
        target_format="md",
    )


def convert_html_to_markdown_text(
    *,
    html_text: str,
    html_path: str,
    output_folder: str,
    resource_dir: str | None = None,
    keep_images: bool = True,
    enable_ocr: bool = False,
    image_link_style: str = "wiki_embed",
    md_file_link_style: str = "wiki_embed",
    ocr_blockquote_title: str = "",
    unified_timestamp_desc: str = "export",
    ocr_language: str = "auto",
    current_locale: str = "zh_CN",
) -> str:
    """Convert HTML text to Markdown with image preprocessing.

    Preprocesses ``<img>`` elements via the shared image-materialisation
    chain (data URI decoding, base-href resolution, local-path copying,
    optional OCR), converts the HTML body to Markdown via ``markdownify``,
    then replaces image tokens with the generated Markdown links.

    This is the entry point for HTML→Markdown conversion within the
    layout plugin.  It reuses ``docwen_plugin_layout.preprocess`` for
    image handling and ``docwen_core.links`` for shared utilities.

    F-G3-004, F-G3-005, F-G3-007
    """
    from docwen_plugin_layout.preprocess import preprocess_html_images

    pre = preprocess_html_images(
        html_text=html_text,
        html_path=html_path,
        output_folder=output_folder,
        resource_dir=resource_dir,
        keep_images=keep_images,
        enable_ocr=enable_ocr,
        image_link_style=image_link_style,
        md_file_link_style=md_file_link_style,
        ocr_blockquote_title=ocr_blockquote_title,
        unified_timestamp_desc=unified_timestamp_desc,
        ocr_language=ocr_language,
        current_locale=current_locale,
    )

    import re

    from markdownify import markdownify

    # ── Extract title from HTML ──────────────────────────────────────
    title_m = re.search(
        r"<title[^>]*>(.*?)</title>",
        pre["html"],
        re.IGNORECASE | re.DOTALL,
    )
    title = title_m.group(1).strip() if title_m else ""

    # ── Convert to Markdown ───────────────────────────────────────────
    md_body = markdownify(pre["html"])

    # ── Replace tokens with Markdown image links ──────────────────────
    for token, link in pre["token_map"].items():  # pyright: ignore[reportAttributeAccessIssue]
        md_body = md_body.replace(token, link)

    # ── Prefix with title heading when available ──────────────────────
    if title:
        md_body = f"# {title}\n\n{md_body}"

    return md_body


class LayoutToMarkdownConverter:
    """Convert admitted fixed-layout files (PDF/OFD/XPS) to Markdown.

    For non-PDF inputs the file must first be preprocessed to PDF (handled
    by the preprocess layer).  This converter assumes the input is already
    a PDF.
    """

    def convert(
        self,
        context: ConverterContext,
        *,
        input_path_override: str | None = None,
        source_format_override: str | None = None,
    ) -> ConversionResult:
        from docwen_core.models.artifact import (
            ARTIFACT_KIND_IMAGE,
            ARTIFACT_KIND_PRIMARY,
            ArtifactManifest,
        )
        from docwen_core.models.result import (
            ConversionDiagnostic,
            ConversionErrorInfo,
            ConversionMetrics,
            ConversionResult,
        )

        task_id = context.request.request_id
        input_path = input_path_override or context.workspace.input_path
        options = context.request.options
        # Preprocessed OFD/XPS paths are internal implementation details.
        # Preserve the original request stem in Markdown/YAML/assets.
        input_stem_val = input_stem(context.workspace.input_path)
        effective_source = str(source_format_override or request_source_format(context)).strip().lower()

        context.cancellation.check()
        context.progress.report_progress(0.0, "Starting layout to Markdown conversion")

        keep_images = bool(options.get("to_md_keep_images", True))
        enable_ocr = bool(options.get("to_md_enable_ocr", False))
        render_dpi = int(options.get("render_dpi", 200))
        from docwen_core.export_semantics import get_markdown_export_modes, resolve_markdown_request_policy

        request_policy = resolve_markdown_request_policy(context)
        export_semantics = request_policy.export

        export_modes = get_markdown_export_modes(
            "layout",
            extraction_mode=options.get("image_mode"),
            semantics=export_semantics,
        )
        image_mode = str(export_modes["image_extraction_mode"]).strip().lower()
        if image_mode not in {"file", "base64", "embed", "omit"}:
            image_mode = "file"
        ocr_language = str(options.get("ocr_language") or "auto")
        current_locale = str(options.get("locale") or "zh_CN")

        try:
            import pymupdf4llm
        except ImportError:
            msg = "pymupdf4llm is not installed. Install it with: pip install pymupdf4llm"
            context.logger.error(msg)
            return ConversionResult(
                task_id=task_id,
                success=False,
                error=ConversionErrorInfo(
                    error_type="dependency_missing",
                    message=msg,
                    diagnostic_code="PDF2MD-DEPENDENCY-MISSING",
                ),
                diagnostics=[
                    ConversionDiagnostic(
                        level="error",
                        message=msg,
                        code="PDF2MD-DEPENDENCY-MISSING",
                    )
                ],
            )

        input_bytes = file_size(input_path)

        try:
            context.cancellation.check()
            context.progress.report_progress(10.0, "Extracting text from PDF")

            # Open the admitted PDF explicitly.  Passing a path directly to
            # pymupdf4llm would make its PyMuPDF boundary infer the parser from
            # a potentially misleading user suffix again.
            import fitz

            write_images = keep_images
            extracted_image_pages: dict[Path, set[int]] = {}
            images_prefix = f"{input_stem_val}_images/"
            with fitz.open(input_path, filetype="pdf") as document:
                physical_page_count = document.page_count
                images_dir_path = Path(context.workspace.staging_dir) / f"{input_stem_val}_images"
                if write_images:
                    images_dir_path.mkdir(parents=True, exist_ok=True)
                page_markdown: list[str] = []
                known_images = (
                    {
                        path
                        for path in images_dir_path.iterdir()
                        if write_images and path.is_file() and _detected_image_format(path) is not None
                    }
                    if write_images
                    else set()
                )
                baseline_images = set(known_images)
                for page_index in range(physical_page_count):
                    context.cancellation.check()
                    page_result = pymupdf4llm.to_markdown(
                        document,
                        pages=[page_index],
                        write_images=write_images,
                        image_path=str(images_dir_path) if write_images else "",
                        image_format="png",
                        use_ocr=False,
                        page_chunks=False,
                    )
                    if isinstance(page_result, list):
                        page_text = "\n".join(str(item) for item in page_result)
                    else:
                        page_text = str(page_result)
                    page_markdown.append(page_text)
                    if write_images:
                        current_images = {
                            path
                            for path in images_dir_path.iterdir()
                            if path.is_file() and _detected_image_format(path) is not None
                        }
                        for image_path in current_images - known_images:
                            extracted_image_pages.setdefault(image_path, set()).add(page_index + 1)
                        for image_path in _referenced_extracted_images(
                            page_text,
                            images_prefix=images_prefix,
                            image_files=current_images - baseline_images,
                        ):
                            extracted_image_pages.setdefault(image_path, set()).add(page_index + 1)
                        known_images = current_images

            md_text_raw = "\n".join(page_markdown)

            context.cancellation.check()

            # Normalise: pymupdf4llm may return a list of strings
            if isinstance(md_text_raw, list):
                md_text = "\n".join(str(item) for item in md_text_raw)
            else:
                md_text = str(md_text_raw)

            # ── Collect image artifacts (MF-1 fix) ──────────────────
            image_artifacts: list[ArtifactManifest] = []
            page_artifacts: list[ArtifactManifest] = []
            ocr_diagnostics: list[ConversionDiagnostic] = []
            resource_diagnostics: list[ConversionDiagnostic] = []
            image_link_style = "wiki_embed"

            if write_images:
                configured_style = str(options.get("image_link_style") or export_semantics.image_link_style).strip()
                if configured_style in {"wiki_embed", "wiki_link", "markdown_embed", "markdown_link"}:
                    image_link_style = configured_style
                images_dir_path = Path(context.workspace.staging_dir) / f"{input_stem_val}_images"
                extracted_image_files: list[Path] = []
                if images_dir_path.exists() and images_dir_path.is_dir():
                    for img_file in sorted(images_dir_path.iterdir()):
                        if not img_file.is_file() or _detected_image_format(img_file) is None:
                            continue
                        extracted_image_files.append(img_file)

                if input_path_override and extracted_image_files:
                    original_files = list(extracted_image_files)
                    md_text, extracted_image_files = _rename_preprocessed_image_refs(
                        md_text,
                        image_files=extracted_image_files,
                        effective_input_path=input_path,
                        original_input_path=context.workspace.input_path,
                    )
                    extracted_image_pages = {
                        renamed: extracted_image_pages.get(original, set())
                        for original, renamed in zip(original_files, extracted_image_files, strict=True)
                    }

                for img_file in extracted_image_files:
                    if keep_images and image_mode not in {"base64", "omit"}:
                        image_format = _detected_image_format(img_file)
                        if image_format is None:
                            continue
                        source_pages = extracted_image_pages.get(img_file, set())
                        metadata: dict[str, object] = {
                            "source": "pdf_extracted",
                            "parent_artifact": "markdown",
                        }
                        if len(source_pages) == 1:
                            metadata["source_page"] = next(iter(source_pages))
                        img_artifact = ArtifactManifest(
                            artifact_id=new_artifact_id(),
                            kind=ARTIFACT_KIND_IMAGE,
                            staging_path=str(img_file),
                            suggested_name=img_file.name,
                            media_type=get_media_type(image_format),
                            metadata=metadata,
                            is_primary=False,
                        )
                        image_artifacts.append(img_artifact)
                        if "source_page" not in metadata:
                            resource_diagnostics.append(
                                ConversionDiagnostic(
                                    level="warning",
                                    message="Extracted image page ownership could not be proven; assigned to the primary document.",
                                    code="resource_page_unresolved",
                                    artifact_id=img_artifact.artifact_id,
                                )
                            )

                if extracted_image_files:
                    md_text, _ = _rewrite_extracted_image_refs(
                        md_text,
                        images_prefix=images_prefix,
                        image_files=extracted_image_files,
                        image_mode=image_mode,
                        image_link_style=image_link_style,
                        export_semantics=export_semantics,
                    )

            page_outcomes: list[OcrOutcome] = []
            if enable_ocr:
                context.progress.report_progress(40.0, "Running physical-page OCR")
                try:
                    page_outcomes = _ocr_page_outcomes(
                        input_path,
                        context.workspace.staging_dir,
                        dpi=render_dpi,
                        ocr_language=ocr_language,
                        current_locale=current_locale,
                    )
                except Exception as exc:
                    page_outcomes = [
                        OcrOutcome(OcrStatus.RECOGNITION_FAILED, message=str(exc)) for _ in range(physical_page_count)
                    ]
                if len(page_outcomes) != physical_page_count:
                    raise ValueError("physical-page OCR did not return exactly one outcome per page")
                for page_number, outcome in enumerate(page_outcomes, start=1):
                    location = f"{Path(input_path).name}:page-{page_number}"
                    page_path = context.workspace.create_artifact_path("auxiliary", ".md")
                    page_text = outcome.recognized_text.rstrip()
                    Path(page_path).write_text(f"{page_text}\n" if page_text else "", encoding="utf-8")
                    page_artifact = ArtifactManifest(
                        artifact_id=new_artifact_id(),
                        kind="auxiliary",
                        staging_path=page_path,
                        suggested_name=f"{input_stem_val}__page_{page_number:04d}_ocr.md",
                        media_type="text/markdown",
                        metadata={
                            "fragment_kind": "page",
                            "page_index": page_number,
                            "page_count": physical_page_count,
                            "source_page": page_number,
                            "ocr_status": outcome.status.value,
                        },
                        is_primary=False,
                    )
                    page_artifacts.append(page_artifact)
                    if warning := format_ocr_best_effort_warning(outcome.status):
                        ocr_diagnostics.append(
                            ConversionDiagnostic(
                                level="warning",
                                message=warning,
                                code="OCR-BEST-EFFORT",
                                location=location,
                                artifact_id=page_artifact.artifact_id,
                            )
                        )
                context.progress.report_progress(60.0, "Physical-page OCR complete")

            # ── Resolve embedded wiki links only when no generated image refs
            # are present. Generated PDF image refs point at staging artifacts,
            # not paths next to the source PDF, so the generic link resolver
            # would incorrectly mark them as missing.
            if not write_images:
                md_text = _resolve_md_links(md_text, input_path)

            from docwen_core.yaml_tools import generate_basic_yaml_frontmatter

            yaml_frontmatter = generate_basic_yaml_frontmatter(
                input_stem_val,
                yaml_key_labels=options.get("yaml_key_labels"),
            )
            md_text = yaml_frontmatter + md_text.lstrip()

            # ── Write Markdown artifact to staging ──────────────────
            md_path = context.workspace.create_artifact_path(ARTIFACT_KIND_PRIMARY, ".md")
            Path(md_path).write_text(md_text, encoding="utf-8")

            artifact = ArtifactManifest(
                artifact_id=new_artifact_id(),
                kind=ARTIFACT_KIND_PRIMARY,
                staging_path=md_path,
                suggested_name=f"{input_stem_val}.md",
                media_type="text/markdown",
                metadata={
                    "source_format": effective_source,
                    "physical_page_count": physical_page_count,
                    "keep_images": keep_images,
                    "ocr_enabled": enable_ocr,
                },
                is_primary=True,
            )
            all_artifacts = [artifact, *page_artifacts, *image_artifacts]
            for output_artifact in all_artifacts:
                context.workspace.add_artifact(output_artifact)
                context.progress.report_artifact_ready(
                    output_artifact.artifact_id,
                    output_artifact.suggested_name,
                )
            output_bytes = file_size(md_path) + sum(file_size(a.staging_path) for a in page_artifacts + image_artifacts)
            context.progress.report_progress(100.0, "Layout to Markdown complete")

            diagnostics = [
                ConversionDiagnostic(
                    level="info",
                    message=f"Converted {input_stem_val} to Markdown ({output_bytes} bytes)",
                    code="PDF2MD-OK",
                )
            ]
            ocr_chars = sum(len(outcome.recognized_text) for outcome in page_outcomes)
            if page_outcomes:
                diagnostics.append(
                    ConversionDiagnostic(
                        level="info",
                        message=f"OCR processed {len(page_outcomes)} physical page(s), {ocr_chars} total characters",
                        code="PDF2MD-OCR-OK",
                    )
                )
            diagnostics.extend(ocr_diagnostics)
            diagnostics.extend(resource_diagnostics)

            return ConversionResult(
                task_id=task_id,
                success=True,
                artifacts=all_artifacts,
                diagnostics=diagnostics,
                metrics=ConversionMetrics(
                    input_bytes=input_bytes,
                    output_bytes=output_bytes,
                    extra={
                        "keep_images": keep_images,
                        "image_mode": image_mode,
                        "image_count": len(image_artifacts),
                        "ocr_enabled": enable_ocr,
                        "ocr_pages": len(page_outcomes),
                        "ocr_images": 0,
                    },
                ),
            )

        except Exception as exc:
            context.logger.error(f"Layout to Markdown failed: {exc}")
            return ConversionResult(
                task_id=task_id,
                success=False,
                error=ConversionErrorInfo(
                    error_type="conversion_failed",
                    message=str(exc),
                    diagnostic_code="PDF2MD-ERROR",
                ),
                diagnostics=[
                    ConversionDiagnostic(
                        level="error",
                        message=f"Layout to Markdown failed: {exc}",
                        code="PDF2MD-ERROR",
                    )
                ],
            )

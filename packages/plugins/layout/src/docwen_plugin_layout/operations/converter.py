"""PDF operation converters — merge and split.

- PdfMerger: merges multiple PDF files into one (action: merge_pdfs)
- PdfSplitter: splits a PDF by page ranges (action: split_pdf)
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import TYPE_CHECKING

from docwen_core.errors import CancellationRequested
from docwen_core.page_selection import PageSelectionError, validate_pdf_split_selection
from docwen_core.paths import input_stem
from docwen_plugin_layout._common import file_size, new_artifact_id

if TYPE_CHECKING:
    from typing import Any

    import pymupdf

    from docwen_core.models.result import ConversionResult
    from docwen_core.protocols.execution_context import ConverterContext


_EXPLICIT_GOTO = re.compile(r"\[\s*(\d+)\s+\d+\s+R\s*(/[^\]]+)\]")


def _ensure_fitz():
    """Import PyMuPDF; raise a user-friendly ImportError if absent."""
    try:
        import pymupdf as fitz
    except ImportError:
        raise ImportError("PyMuPDF is not installed. Install it with: pip install PyMuPDF") from None
    return fitz


def _is_valid_pdf(file_path: str) -> bool:
    """Check whether a file starts with the PDF magic bytes ``%PDF-``."""
    try:
        with open(file_path, "rb") as fh:
            return fh.read(5) == b"%PDF-"
    except OSError:
        return False


def _goto_links(page: pymupdf.Page, kind: int) -> list[dict[str, Any]]:
    return [link for link in page.get_links() if link.get("kind") == kind]


def _restore_explicit_gotos(
    source: pymupdf.Document,
    destination: pymupdf.Document,
    original_links: list[dict[str, Any]],
    copied_links: list[dict[str, Any]],
    destination_by_source: dict[int, int],
) -> None:
    """Retain the source PDF's hotspot, destination view and zoom.

    The copy routines emit links in order. Check the complete page mapping
    before changing a copied link, so a skipped link cannot shift the pairing.
    """
    if [destination_by_source[link["page"]] for link in original_links] != [link["page"] for link in copied_links]:
        return

    for original, copied in zip(original_links, copied_links, strict=True):
        xref, copied_xref = original.get("xref", 0), copied.get("xref", 0)
        if xref <= 0 or copied_xref <= 0:
            continue
        value_type, value = source.xref_get_key(xref, "A/D")
        if value_type == "null":
            value_type, value = source.xref_get_key(xref, "Dest")
        if value_type == "xref":
            value = source.xref_object(int(value.split()[0]), compressed=True)
        elif value_type != "array":
            continue
        match = _EXPLICIT_GOTO.fullmatch(value.strip())
        if match is None or int(match[1]) != source.page_xref(original["page"]):
            continue
        if destination.xref_get_key(copied_xref, "A/S") != ("name", "/GoTo"):
            continue
        target_xref = destination.page_xref(copied["page"])
        destination.xref_set_key(copied_xref, "A/D", f"[{target_xref} 0 R {match[2]}]")
        rect_type, rect = source.xref_get_key(xref, "Rect")
        if rect_type == "array":
            destination.xref_set_key(copied_xref, "Rect", rect)


def _inline_indirect_acroform(source: pymupdf.Document) -> None:
    """Prevent repeated page grafts from clearing the in-memory source field tree."""
    catalog = source.pdf_catalog()
    value_type, value = source.xref_get_key(catalog, "AcroForm")
    if value_type == "xref":
        source.xref_set_key(catalog, "AcroForm", source.xref_object(int(value.split()[0]), compressed=True))


def _copy_pages_preserving_internal_gotos(src, destination, pages: list[int], cancellation) -> None:
    """Copy 1-based pages and remap safe GOTO links kept in the same output."""
    fitz = _ensure_fitz()
    source_indices = [page - 1 for page in pages]
    destination_by_source = {
        source_index: destination_index for destination_index, source_index in enumerate(source_indices)
    }

    for source_index in source_indices:
        cancellation.check()
        destination.insert_pdf(src, from_page=source_index, to_page=source_index)

    for source_index, destination_index in destination_by_source.items():
        cancellation.check()
        link_candidates = [
            link for link in _goto_links(src[source_index], fitz.LINK_GOTO) if link.get("page") in destination_by_source
        ]
        local_links = [link for link in link_candidates if link["page"] == source_index]
        cross_page_links = [link for link in link_candidates if link["page"] != source_index]
        if not link_candidates:
            continue
        output_page = destination[destination_index]
        for link in cross_page_links:
            remapped_link = {
                "kind": fitz.LINK_GOTO,
                "from": link["from"],
                "page": destination_by_source[link["page"]],
                "to": link.get("to", fitz.Point(0, 0)),
                "zoom": link.get("zoom", 0.0),
            }
            output_page.insert_link(remapped_link)
        if cross_page_links:
            output_page = destination.reload_page(output_page)
        _restore_explicit_gotos(
            src,
            destination,
            local_links + cross_page_links,
            _goto_links(output_page, fitz.LINK_GOTO),
            destination_by_source,
        )


# ═══════════════════════════════════════════════════════════════════════
# PdfMerger
# ═══════════════════════════════════════════════════════════════════════


class PdfMerger:
    """Merge multiple PDF files into a single PDF.

    All input files must already be PDFs (preprocessing to PDF is the
    caller's responsibility).  Pages are concatenated in the order the
    input files appear in ``context.request.input_refs``.
    """

    def convert(self, context: ConverterContext) -> ConversionResult:
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

        fitz = _ensure_fitz()
        task_id = context.request.request_id
        input_refs = context.request.input_refs

        context.cancellation.check()

        if len(input_refs) < 2:
            msg = "At least two files are required for PDF merge."
            return ConversionResult(
                task_id=task_id,
                success=False,
                error=ConversionErrorInfo(
                    error_type="invalid_input",
                    message=msg,
                    diagnostic_code="PDF-MERGE-INVALID-INPUT",
                ),
                diagnostics=[ConversionDiagnostic(level="error", message=msg, code="PDF-MERGE-INVALID-INPUT")],
            )

        context.progress.report_progress(0.0, f"Merging {len(input_refs)} PDFs")

        # Validate all inputs are valid PDFs before attempting merge
        for ref in input_refs:
            if not _is_valid_pdf(ref.path):
                msg = f"File '{Path(ref.path).name}' is not a valid PDF. All input files for merge must be valid PDFs."
                return ConversionResult(
                    task_id=task_id,
                    success=False,
                    error=ConversionErrorInfo(
                        error_type="invalid_input",
                        message=msg,
                        diagnostic_code="PDF-MERGE-INVALID-INPUT",
                    ),
                    diagnostics=[ConversionDiagnostic(level="error", message=msg, code="PDF-MERGE-INVALID-INPUT")],
                )

        total_input_bytes = 0
        merged = fitz.open()

        try:
            for i, ref in enumerate(input_refs):
                context.cancellation.check()
                context.progress.report_progress(
                    (i / len(input_refs)) * 90.0,
                    f"Adding file {i + 1}/{len(input_refs)}",
                )
                total_input_bytes += file_size(ref.path)
                try:
                    with fitz.open(ref.path, filetype="pdf") as src:
                        first_page = merged.page_count
                        merged.insert_pdf(src)
                        destination_by_source = {index: first_page + index for index in range(src.page_count)}
                        for source_index, destination_index in destination_by_source.items():
                            context.cancellation.check()
                            originals = _goto_links(src[source_index], fitz.LINK_GOTO)
                            if originals:
                                _restore_explicit_gotos(
                                    src,
                                    merged,
                                    [link for link in originals if link.get("page") in destination_by_source],
                                    _goto_links(merged[destination_index], fitz.LINK_GOTO),
                                    destination_by_source,
                                )
                except CancellationRequested:
                    raise
                except Exception as exc:
                    raise RuntimeError(f"Failed to merge '{Path(ref.path).name}': {exc}") from exc

            context.cancellation.check()

            output_path = context.workspace.create_artifact_path(ARTIFACT_KIND_PRIMARY, ".pdf")
            merged.save(output_path)
        except CancellationRequested:
            raise
        except Exception as exc:
            context.logger.error(f"PDF merge failed: {exc}")
            return ConversionResult(
                task_id=task_id,
                success=False,
                error=ConversionErrorInfo(
                    error_type="conversion_failed",
                    message=str(exc),
                    diagnostic_code="PDF-MERGE-ERROR",
                ),
                diagnostics=[
                    ConversionDiagnostic(
                        level="error",
                        message=f"PDF merge failed: {exc}",
                        code="PDF-MERGE-ERROR",
                    )
                ],
            )
        finally:
            merged.close()

        artifact = ArtifactManifest(
            artifact_id=new_artifact_id(),
            kind=ARTIFACT_KIND_PRIMARY,
            staging_path=output_path,
            suggested_name="merged.pdf",
            media_type="application/pdf",
            metadata={
                "input_count": len(input_refs),
                "action": "merge_pdfs",
            },
            is_primary=True,
        )
        context.workspace.add_artifact(artifact)
        context.progress.report_artifact_ready(artifact.artifact_id, artifact.suggested_name)
        context.progress.report_progress(100.0, "PDF merge complete")

        out_bytes = file_size(output_path)
        return ConversionResult(
            task_id=task_id,
            success=True,
            artifacts=[artifact],
            diagnostics=[
                ConversionDiagnostic(
                    level="info",
                    message=f"Merged {len(input_refs)} PDFs ({out_bytes} bytes)",
                    code="PDF-MERGE-OK",
                )
            ],
            metrics=ConversionMetrics(
                input_bytes=total_input_bytes,
                output_bytes=out_bytes,
                extra={"input_count": len(input_refs)},
            ),
        )


# ═══════════════════════════════════════════════════════════════════════
# PdfSplitter
# ═══════════════════════════════════════════════════════════════════════


class PdfSplitter:
    """Split a PDF into multiple files.

    Supports three split modes:
    - ``custom``: split by user-provided page numbers (1-based)
    - ``every_page``: one file per page
    - ``odd_even``: odd pages in one file, even pages in another
    """

    def convert(self, context: ConverterContext) -> ConversionResult:
        from docwen_core.models.result import (
            ConversionDiagnostic,
            ConversionErrorInfo,
            ConversionResult,
        )

        fitz = _ensure_fitz()
        task_id = context.request.request_id
        input_path = context.workspace.input_path
        options = context.request.options
        stem = input_stem(input_path)

        split_mode = options.get("split_mode", "custom")

        context.cancellation.check()
        context.progress.report_progress(0.0, "Starting PDF split")

        try:
            with fitz.open(input_path, filetype="pdf") as src:
                total_pages = len(src)

                try:
                    selection = validate_pdf_split_selection(
                        split_mode, options.get("pages", []), total_pages=total_pages
                    )
                except PageSelectionError as exc:
                    msg = str(exc)
                    return ConversionResult(
                        task_id=task_id,
                        success=False,
                        error=ConversionErrorInfo(
                            error_type="invalid_input",
                            message=msg,
                            diagnostic_code="PDF-SPLIT-INVALID-INPUT",
                        ),
                        diagnostics=[ConversionDiagnostic(level="error", message=msg, code="PDF-SPLIT-INVALID-INPUT")],
                    )

                context.cancellation.check()
                _inline_indirect_acroform(src)

                if selection.mode == "every_page":
                    return self._split_every_page(src, total_pages, task_id, stem, context)
                elif selection.mode == "odd_even":
                    return self._split_odd_even(src, total_pages, task_id, stem, context)
                else:
                    return self._split_custom(src, list(selection.pages), total_pages, task_id, stem, context)
        except CancellationRequested:
            raise
        except Exception as exc:
            context.logger.error(f"PDF split failed: {exc}")
            return ConversionResult(
                task_id=task_id,
                success=False,
                error=ConversionErrorInfo(
                    error_type="conversion_failed",
                    message=str(exc),
                    diagnostic_code="PDF-SPLIT-ERROR",
                ),
                diagnostics=[
                    ConversionDiagnostic(
                        level="error",
                        message=f"PDF split failed: {exc}",
                        code="PDF-SPLIT-ERROR",
                    )
                ],
            )

    # ── split mode helpers ─────────────────────────────────────────

    def _split_every_page(self, src, total_pages: int, task_id: str, stem: str, context) -> ConversionResult:
        from docwen_core.models.artifact import (
            ARTIFACT_KIND_PRIMARY,
            ArtifactManifest,
        )
        from docwen_core.models.result import (
            ConversionDiagnostic,
            ConversionMetrics,
            ConversionResult,
        )

        fitz = _ensure_fitz()
        artifacts: list[ArtifactManifest] = []
        total_out = 0

        for page_num in range(1, total_pages + 1):
            context.cancellation.check()
            context.progress.report_progress(
                (page_num / total_pages) * 100.0,
                f"Extracting page {page_num}/{total_pages}",
            )

            output_path = context.workspace.create_artifact_path(ARTIFACT_KIND_PRIMARY, ".pdf")
            single = fitz.open()
            try:
                _copy_pages_preserving_internal_gotos(src, single, [page_num], context.cancellation)
                single.save(output_path)
            finally:
                single.close()

            suggested = f"{stem}_p{page_num}.pdf"
            artifact = ArtifactManifest(
                artifact_id=new_artifact_id(),
                kind=ARTIFACT_KIND_PRIMARY,
                staging_path=output_path,
                suggested_name=suggested,
                media_type="application/pdf",
                metadata={"page": page_num, "split_mode": "every_page"},
                is_primary=(page_num == 1),
            )
            context.workspace.add_artifact(artifact)
            context.progress.report_artifact_ready(artifact.artifact_id, suggested)
            artifacts.append(artifact)
            total_out += file_size(output_path)

        context.progress.report_progress(100.0, "PDF split complete")
        return ConversionResult(
            task_id=task_id,
            success=True,
            artifacts=artifacts,
            diagnostics=[
                ConversionDiagnostic(
                    level="info",
                    message=f"Split PDF into {total_pages} pages",
                    code="PDF-SPLIT-OK",
                )
            ],
            metrics=ConversionMetrics(
                input_bytes=file_size(context.workspace.input_path),
                output_bytes=total_out,
                extra={"split_mode": "every_page", "page_count": total_pages},
            ),
        )

    def _split_odd_even(self, src, total_pages: int, task_id: str, stem: str, context) -> ConversionResult:
        from docwen_core.models.artifact import (
            ARTIFACT_KIND_PRIMARY,
            ArtifactManifest,
        )
        from docwen_core.models.result import (
            ConversionDiagnostic,
            ConversionMetrics,
            ConversionResult,
        )

        fitz = _ensure_fitz()
        odd_pages = [p for p in range(1, total_pages + 1) if p % 2 == 1]
        even_pages = [p for p in range(1, total_pages + 1) if p % 2 == 0]

        artifacts: list[ArtifactManifest] = []
        total_out = 0

        # Odd pages
        context.progress.report_progress(20.0, "Extracting odd pages")
        odd_path = context.workspace.create_artifact_path(ARTIFACT_KIND_PRIMARY, ".pdf")
        odd_doc = fitz.open()
        try:
            _copy_pages_preserving_internal_gotos(src, odd_doc, odd_pages, context.cancellation)
            odd_doc.save(odd_path)
        finally:
            odd_doc.close()

        odd_artifact = ArtifactManifest(
            artifact_id=new_artifact_id(),
            kind=ARTIFACT_KIND_PRIMARY,
            staging_path=odd_path,
            suggested_name=f"{stem}_odd.pdf",
            media_type="application/pdf",
            metadata={"pages": odd_pages, "split_mode": "odd_even"},
            is_primary=True,
        )
        context.workspace.add_artifact(odd_artifact)
        artifacts.append(odd_artifact)
        total_out += file_size(odd_path)

        # Even pages (if any)
        if even_pages:
            context.cancellation.check()
            context.progress.report_progress(60.0, "Extracting even pages")
            even_path = context.workspace.create_artifact_path(ARTIFACT_KIND_PRIMARY, ".pdf")
            even_doc = fitz.open()
            try:
                _copy_pages_preserving_internal_gotos(src, even_doc, even_pages, context.cancellation)
                even_doc.save(even_path)
            finally:
                even_doc.close()

            even_artifact = ArtifactManifest(
                artifact_id=new_artifact_id(),
                kind=ARTIFACT_KIND_PRIMARY,
                staging_path=even_path,
                suggested_name=f"{stem}_even.pdf",
                media_type="application/pdf",
                metadata={"pages": even_pages, "split_mode": "odd_even"},
                is_primary=False,
            )
            context.workspace.add_artifact(even_artifact)
            artifacts.append(even_artifact)
            total_out += file_size(even_path)

        context.progress.report_progress(100.0, "PDF odd/even split complete")
        return ConversionResult(
            task_id=task_id,
            success=True,
            artifacts=artifacts,
            diagnostics=[
                ConversionDiagnostic(
                    level="info",
                    message=f"Split PDF into odd ({len(odd_pages)}p) and even ({len(even_pages)}p)",
                    code="PDF-SPLIT-OK",
                )
            ],
            metrics=ConversionMetrics(
                input_bytes=file_size(context.workspace.input_path),
                output_bytes=total_out,
                extra={"split_mode": "odd_even", "odd_pages": len(odd_pages), "even_pages": len(even_pages)},
            ),
        )

    def _split_custom(
        self,
        src,
        pages: list[int],
        total_pages: int,
        task_id: str,
        stem: str,
        context,
    ) -> ConversionResult:
        from docwen_core.models.artifact import (
            ARTIFACT_KIND_PRIMARY,
            ArtifactManifest,
        )
        from docwen_core.models.result import (
            ConversionDiagnostic,
            ConversionMetrics,
            ConversionResult,
        )

        fitz = _ensure_fitz()

        # convert() has validated the entire typed selection against this PDF.
        valid = pages
        selected = set(valid)
        remaining = [page for page in range(1, total_pages + 1) if page not in selected]

        artifacts: list[ArtifactManifest] = []
        total_out = 0

        # Part 1: user-selected pages
        context.progress.report_progress(30.0, "Creating part 1")
        part1_path = context.workspace.create_artifact_path(ARTIFACT_KIND_PRIMARY, ".pdf")
        doc1 = fitz.open()
        try:
            _copy_pages_preserving_internal_gotos(src, doc1, valid, context.cancellation)
            doc1.save(part1_path)
        finally:
            doc1.close()

        part1_artifact = ArtifactManifest(
            artifact_id=new_artifact_id(),
            kind=ARTIFACT_KIND_PRIMARY,
            staging_path=part1_path,
            suggested_name=f"{stem}_part1.pdf",
            media_type="application/pdf",
            metadata={"pages": valid, "split_mode": "custom"},
            is_primary=True,
        )
        context.workspace.add_artifact(part1_artifact)
        artifacts.append(part1_artifact)
        total_out += file_size(part1_path)

        # Part 2: remaining pages
        context.cancellation.check()
        context.progress.report_progress(70.0, "Creating part 2")
        part2_path = context.workspace.create_artifact_path(ARTIFACT_KIND_PRIMARY, ".pdf")
        doc2 = fitz.open()
        try:
            _copy_pages_preserving_internal_gotos(src, doc2, remaining, context.cancellation)
            doc2.save(part2_path)
        finally:
            doc2.close()

        part2_artifact = ArtifactManifest(
            artifact_id=new_artifact_id(),
            kind=ARTIFACT_KIND_PRIMARY,
            staging_path=part2_path,
            suggested_name=f"{stem}_part2.pdf",
            media_type="application/pdf",
            metadata={"pages": remaining, "split_mode": "custom"},
            is_primary=False,
        )
        context.workspace.add_artifact(part2_artifact)
        artifacts.append(part2_artifact)
        total_out += file_size(part2_path)

        context.progress.report_progress(100.0, "PDF custom split complete")
        return ConversionResult(
            task_id=task_id,
            success=True,
            artifacts=artifacts,
            diagnostics=[
                ConversionDiagnostic(
                    level="info",
                    message=f"Split PDF: part 1 ({len(valid)}p) + part 2 ({len(remaining)}p)",
                    code="PDF-SPLIT-OK",
                )
            ],
            metrics=ConversionMetrics(
                input_bytes=file_size(context.workspace.input_path),
                output_bytes=total_out,
                extra={"split_mode": "custom", "part1_pages": len(valid), "part2_pages": len(remaining)},
            ),
        )

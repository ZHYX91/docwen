"""MdNumberingProcessor — handles MD heading numbering (clean/add)."""

from __future__ import annotations

import time
from pathlib import Path
from typing import TYPE_CHECKING

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
from docwen_core.text.heading_numbering import NumberingSchemeResolutionError
from docwen_plugin_markdown.common_utils import (
    add_md_numbering,
    read_input_markdown,
    remove_md_numbering,
)

if TYPE_CHECKING:
    from docwen_core.protocols.execution_context import NumberingConverterContext


MEDIA_TYPE_MD = "text/markdown"


class MdNumberingProcessor:
    """Process heading numbering in Markdown files.

    Supports two operations (can be combined):
    - remove_numbering: strip existing heading numbering patterns
    - add_numbering: apply a numbering scheme to headings

    Writes the processed content to the staging directory.
    """

    def convert(self, context: NumberingConverterContext) -> ConversionResult:
        t_start = time.monotonic()
        task_id = context.request.request_id
        cancellable = context.cancellation
        logger = context.logger
        progress = context.progress
        workspace = context.workspace
        options = context.request.options

        try:
            cancellable.check()

            # ── Validate options ────────────────────────────────────
            remove_num = options.get("remove_numbering", True)
            add_num = options.get("add_numbering", False)
            scheme = options.get("numbering_scheme", "")

            if not remove_num and not add_num:
                return ConversionResult(
                    task_id=task_id,
                    success=False,
                    error=ConversionErrorInfo(
                        error_type="invalid_input",
                        message=("At least one of remove_numbering or add_numbering must be True"),
                        diagnostic_code="MDNUM-NO-OPERATION",
                    ),
                    metrics=ConversionMetrics(duration_ms=(time.monotonic() - t_start) * 1000.0),
                    diagnostics=[
                        ConversionDiagnostic(
                            level="error",
                            message=("At least one of remove_numbering or add_numbering must be True"),
                            code="MDNUM-NO-OPERATION",
                        )
                    ],
                )

            # ── Read input ──────────────────────────────────────────
            input_path = workspace.input_path
            progress.report_progress(10.0, "Reading Markdown input")
            content, input_bytes = read_input_markdown(input_path)
            # The legacy numbering processor writes normalized LF output; the
            # authenticated read is byte-faithful, so normalize here instead
            # of in the shared reader used by the fenced-source carriers.
            content = content.replace("\r\n", "\n").replace("\r", "\n")

            # ── Process ─────────────────────────────────────────────
            processed = content
            operations: list[str] = []

            if remove_num:
                cancellable.check()
                progress.report_progress(40.0, "Removing heading numbering")
                processed = remove_md_numbering(
                    processed,
                    rules=getattr(context, "heading_cleanup_rules", ()) or (),
                )
                operations.append("remove_numbering")

            if add_num:
                cancellable.check()
                progress.report_progress(70.0, "Adding heading numbering")
                processed = add_md_numbering(processed, scheme, registry=context.numbering_registry)
                operations.append(f"add_numbering({scheme})")

            # ── Write to staging ────────────────────────────────────
            cancellable.check()
            progress.report_progress(90.0, "Writing processed Markdown")

            output_path = workspace.create_artifact_path(ARTIFACT_KIND_PRIMARY, ".md")
            input_stem = Path(input_path).stem
            suggested_name = f"{input_stem}_numbered.md"

            with Path(output_path).open("w", encoding="utf-8", newline="") as handle:
                handle.write(processed)

            output_bytes = Path(output_path).stat().st_size

            # ── Register artifact ───────────────────────────────────
            artifact = ArtifactManifest(
                artifact_id=f"{task_id}-md-numbered",
                kind=ARTIFACT_KIND_PRIMARY,
                staging_path=output_path,
                suggested_name=suggested_name,
                media_type=MEDIA_TYPE_MD,
                is_primary=True,
                metadata={
                    "operations": operations,
                    "scheme": scheme if add_num else None,
                    "original_length": len(content),
                    "processed_length": len(processed),
                },
            )
            workspace.add_artifact(artifact)

            elapsed_ms = (time.monotonic() - t_start) * 1000.0
            progress.report_progress(100.0, "Done")

            logger.info(f"MD numbering complete: {input_path} → {suggested_name} ({', '.join(operations)})")

            return ConversionResult(
                task_id=task_id,
                success=True,
                artifacts=[artifact],
                metrics=ConversionMetrics(
                    duration_ms=elapsed_ms,
                    input_bytes=input_bytes,
                    output_bytes=output_bytes,
                ),
                diagnostics=[
                    ConversionDiagnostic(
                        level="info",
                        message=(f"MD numbering operations: {', '.join(operations)}"),
                        code="MDNUM-OK",
                    )
                ],
            )

        except NumberingSchemeResolutionError as exc:
            elapsed_ms = (time.monotonic() - t_start) * 1000.0
            logger.error(f"MD numbering rejected: {exc}")
            return ConversionResult(
                task_id=task_id,
                success=False,
                error=ConversionErrorInfo(
                    error_type=exc.error_type,
                    message=str(exc),
                    diagnostic_code=exc.diagnostic_code,
                ),
                metrics=ConversionMetrics(duration_ms=elapsed_ms),
                diagnostics=[
                    ConversionDiagnostic(
                        level="error",
                        message=str(exc),
                        code=exc.diagnostic_code,
                    )
                ],
            )
        except Exception as exc:
            elapsed_ms = (time.monotonic() - t_start) * 1000.0
            logger.error(f"MD numbering failed: {exc}")
            return ConversionResult(
                task_id=task_id,
                success=False,
                error=ConversionErrorInfo(
                    error_type="conversion_failed",
                    message=str(exc),
                    diagnostic_code="MDNUM-ERROR",
                ),
                metrics=ConversionMetrics(duration_ms=elapsed_ms),
                diagnostics=[
                    ConversionDiagnostic(
                        level="error",
                        message=f"MD numbering failed: {exc}",
                        code="MDNUM-ERROR",
                    )
                ],
            )

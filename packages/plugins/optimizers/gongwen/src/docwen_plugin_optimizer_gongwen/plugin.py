"""GongwenOptimizerPlugin — entry point for docwen_plugin_optimizer_gongwen."""

from __future__ import annotations

from typing import TYPE_CHECKING

from docwen_core.text.heading_numbering import NumberingSchemeResolutionError
from docwen_plugin_optimizer_gongwen.manifest import build_manifest
from docwen_plugin_optimizer_gongwen.pipeline import convert_docx_to_md_gongwen

if TYPE_CHECKING:
    from docwen_core.models.manifest import PluginManifest
    from docwen_core.models.result import ConversionResult
    from docwen_core.protocols.execution_context import PluginExecutionContext


_REQUIRED_FIELD_LABELS: dict[str, str] = {
    "title": "标题",
    "issuing_authority_signature": "发文机关署名",
    "issue_date": "成文日期",
}
_REVIEW_REASON_LABELS: dict[str, str] = {
    "close_unique_match": "存在相近候选",
    "used_fallback": "使用了回退识别",
    "low_confidence_pass": "存在低置信度识别",
}


class GongwenOptimizerPlugin:
    """Plugin for Chinese official document (公文) DOCX → Markdown conversion.

    Recognises gongwen structure, extracts 18 YAML metadata fields via
    three-round scoring and re-evaluation, and renders Markdown output.
    """

    plugin_id: str
    _manifest: PluginManifest | None

    def __init__(self) -> None:
        self.plugin_id = "docwen_plugin_optimizer_gongwen"
        self._manifest = None

    @property
    def manifest(self) -> PluginManifest:
        if self._manifest is None:
            self._manifest = build_manifest()
        return self._manifest

    def can_handle(self, source_format: str, target_format: str, action_name: str = "") -> bool:
        for route in self.manifest.routes:
            if (
                route.source_format == source_format
                and route.target_format == target_format
                and route.action_name == action_name
            ):
                return True
        return False

    def convert(self, context: PluginExecutionContext) -> ConversionResult:
        """Dispatch to the gongwen pipeline."""
        from docx import Document

        from docwen_core.export_semantics import resolve_markdown_request_policy
        from docwen_core.models.result import (
            ConversionDiagnostic,
            ConversionErrorInfo,
            ConversionResult,
        )

        source = context.request.input_refs[0].format if context.request.input_refs else ""
        target = context.request.target_format
        action = context.request.action_name

        # ── Gongwen route ──────────────────────────────────────────────
        if action == "gongwen" and source in ("docx", "document") and target == "md":
            input_path = context.workspace.input_path
            doc = Document(input_path)
            options = dict(context.request.options or {})
            # Image extraction belongs in the task staging directory so the
            # plugin can register resources and OutputFinalizer can place them.
            options["output_dir"] = str(context.workspace.staging_dir)
            request_policy = resolve_markdown_request_policy(context)

            try:
                gongwen_result = convert_docx_to_md_gongwen(
                    doc,
                    input_path,
                    options,
                    cancellation=context.cancellation,
                    progress=context.progress,
                    registry=context.numbering_registry,
                    cleanup_rules=getattr(context, "heading_cleanup_rules", ()) or (),
                    export_semantics=request_policy.export,
                )
            except NumberingSchemeResolutionError as exc:
                return ConversionResult(
                    task_id=context.request.request_id,
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
                )

            if not gongwen_result["success"]:
                raise RuntimeError("Gongwen pipeline returned success=False")

            markdown_content = gongwen_result["markdown"]
            attachment_documents = list(gongwen_result.get("attachment_documents", []))
            if attachment_documents:
                from pathlib import Path

                from docwen_core.markdown_utils import format_md_file_link

                attachment_links = []
                for attachment_document in attachment_documents:
                    attachment_name = f"{Path(input_path).stem}_附件{attachment_document.ordinal:02d}.md"
                    attachment_links.append(
                        f"- {format_md_file_link(attachment_name, style=request_policy.export.md_file_link_style)}"
                    )
                markdown_content = (
                    f"{markdown_content.rstrip()}\n\n## 附件文档\n\n" + "\n".join(attachment_links) + "\n"
                )
            review_signals = gongwen_result["metadata"]["recognition_review_signals"]
            review_summary = review_signals["recognition_summary"]
            needs_review = bool(review_summary["needs_review"])
            missing_required: list[str] = list(review_signals["missing_required"])
            review_reasons: list[str] = list(review_signals["needs_review_reasons"])

            diagnostics = [
                ConversionDiagnostic(
                    level="info",
                    message="Gongwen conversion complete",
                    code="GONGWEN-OK",
                )
            ]
            if needs_review:
                message_parts: list[str] = []
                if missing_required:
                    labels = [_REQUIRED_FIELD_LABELS.get(field, field) for field in missing_required]
                    message_parts.append(f"缺少必需字段：{'、'.join(labels)}")
                if review_reasons:
                    labels = [_REVIEW_REASON_LABELS.get(reason, reason) for reason in review_reasons]
                    message_parts.append(f"识别提示：{'、'.join(labels)}")
                diagnostics.append(
                    ConversionDiagnostic(
                        level="warning",
                        message="；".join(message_parts) or "公文识别结果需要人工复核",
                        code="GONGWEN-NEEDS-REVIEW",
                    )
                )

            # Write output to staging
            output_path = context.workspace.create_artifact_path("primary", ".md")
            try:
                with open(output_path, "w", encoding="utf-8") as f:
                    f.write(markdown_content)
            except OSError as exc:
                return ConversionResult(
                    task_id=context.request.request_id,
                    success=False,
                    error=ConversionErrorInfo(
                        error_type="conversion_failed",
                        message=f"Failed to write output file: {exc}",
                        diagnostic_code="GONGWEN-WRITE-ERROR",
                    ),
                    diagnostics=[
                        ConversionDiagnostic(
                            level="error",
                            message=f"File write error at {output_path}: {exc}",
                            code="GONGWEN-WRITE-ERROR",
                        )
                    ],
                )

            import os
            import uuid

            from docwen_core.models.artifact import ArtifactManifest
            from docwen_core.models.result import ConversionMetrics

            input_basename = os.path.basename(input_path)
            suggested_name = input_basename.rsplit(".", 1)[0] + ".md"

            artifact = ArtifactManifest(
                artifact_id=str(uuid.uuid4()),
                kind="primary",
                staging_path=output_path,
                suggested_name=suggested_name,
                media_type="text/markdown",
                metadata={
                    "paragraph_count": gongwen_result["stats"].get("paragraphs", 0),
                    "gongwen_fields": len([v for v in gongwen_result["yaml_info"].values() if v]),
                    "gongwen_confidence": gongwen_result["metadata"].get("confidence", {}).get("overall", "none"),
                    "gongwen_needs_review": needs_review,
                    "gongwen_missing_required": missing_required,
                    "gongwen_review_reasons": review_reasons,
                },
                is_primary=True,
            )
            context.workspace.add_artifact(artifact)
            artifacts = [artifact]

            for attachment_document in attachment_documents:
                attachment_path = context.workspace.create_artifact_path("auxiliary", ".md")
                try:
                    with open(attachment_path, "w", encoding="utf-8") as f:
                        f.write(attachment_document.markdown)
                except OSError as exc:
                    return ConversionResult(
                        task_id=context.request.request_id,
                        success=False,
                        error=ConversionErrorInfo(
                            error_type="conversion_failed",
                            message=f"Failed to write attachment output file: {exc}",
                            diagnostic_code="GONGWEN-WRITE-ERROR",
                        ),
                        diagnostics=[
                            ConversionDiagnostic(
                                level="error",
                                message=f"File write error at {attachment_path}: {exc}",
                                code="GONGWEN-WRITE-ERROR",
                            )
                        ],
                    )

                attachment_artifact = ArtifactManifest(
                    artifact_id=str(uuid.uuid4()),
                    kind="auxiliary",
                    staging_path=attachment_path,
                    suggested_name=(input_basename.rsplit(".", 1)[0] + f"_附件{attachment_document.ordinal:02d}.md"),
                    media_type="text/markdown",
                    metadata={
                        "source_kind": "gongwen_attachment",
                        "attachment_ordinal": attachment_document.ordinal,
                        "attachment_title": attachment_document.title,
                        "attachment_paragraph_indices": list(attachment_document.paragraph_indices),
                        "attachment_owner_artifact_id": artifact.artifact_id,
                        "paragraph_count": gongwen_result["stats"].get("paragraphs", 0),
                    },
                    is_primary=False,
                )
                context.workspace.add_artifact(attachment_artifact)
                artifacts.append(attachment_artifact)

            import mimetypes

            from docwen_core.models.artifact import ARTIFACT_KIND_IMAGE

            for image_path in dict.fromkeys(gongwen_result.get("image_paths", [])):
                image_name = os.path.basename(image_path)
                image_artifact = ArtifactManifest(
                    artifact_id=str(uuid.uuid4()),
                    kind=ARTIFACT_KIND_IMAGE,
                    staging_path=image_path,
                    suggested_name=image_name,
                    media_type=mimetypes.guess_type(image_name)[0] or "application/octet-stream",
                    metadata={
                        "source_format": "docx",
                        "source_kind": "gongwen_image",
                    },
                    is_primary=False,
                )
                context.workspace.add_artifact(image_artifact)
                artifacts.append(image_artifact)

            import time as _time

            return ConversionResult(
                task_id=context.request.request_id,
                success=True,
                artifacts=artifacts,
                diagnostics=diagnostics,
                error=None,
                metrics=ConversionMetrics(
                    duration_ms=(  # approximate: creation of this result
                        _time.monotonic() * 0
                    ),
                    input_bytes=os.path.getsize(input_path) if os.path.isfile(input_path) else 0,
                    output_bytes=(
                        len(markdown_content.encode("utf-8"))
                        + sum(len(document.markdown.encode("utf-8")) for document in attachment_documents)
                        + sum(
                            os.path.getsize(path)
                            for path in gongwen_result.get("image_paths", [])
                            if os.path.isfile(path)
                        )
                    ),
                    extra=gongwen_result.get("stats", {}),
                ),
            )

        # ── Fallback ──────────────────────────────────────────────────
        msg = f"No handler for action='{action}', source='{source}', target='{target}'"
        return ConversionResult(
            task_id=context.request.request_id,
            success=False,
            error=ConversionErrorInfo(
                error_type="unsupported_route",
                message=msg,
                diagnostic_code="GONGWEN-UNSUPPORTED-ROUTE",
            ),
            diagnostics=[
                ConversionDiagnostic(
                    level="error",
                    message=msg,
                    code="GONGWEN-UNSUPPORTED-ROUTE",
                )
            ],
        )

"""ProofreadPlugin — entry point for docwen_plugin_proofread."""

from __future__ import annotations

from typing import TYPE_CHECKING

from docwen_plugin_proofread.docx_validator import DocxValidator
from docwen_plugin_proofread.manifest import build_manifest
from docwen_plugin_proofread.md_validator import MarkdownValidator

if TYPE_CHECKING:
    from docwen_core.models.manifest import PluginManifest
    from docwen_core.models.result import ConversionResult
    from docwen_core.protocols.execution_context import PluginExecutionContext


class ProofreadPlugin:
    """Plugin for DOCX and Markdown text proofreading.

    Two action routes:
    - validate (docx→docx): Open DOCX, check text per paragraph,
      insert comments for issues, output proofread DOCX.
    - validate (markdown→markdown): Read Markdown, check text (with
      code block / YAML frontmatter sanitization), output JSON report.

    DOC/WPS/RTF/ODT inputs are converted to DOCX by the Application layer;
    this plugin deliberately receives and parses only the derived DOCX.
    """

    plugin_id: str
    _manifest: PluginManifest | None

    def __init__(self) -> None:
        self.plugin_id = "docwen_plugin_proofread"
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
        """Dispatch to the appropriate validator based on action_name."""
        from docwen_core.models.result import (
            ConversionDiagnostic,
            ConversionErrorInfo,
            ConversionResult,
        )

        source = context.request.input_refs[0].format if context.request.input_refs else ""
        target = context.request.target_format
        action = context.request.action_name

        if action == "validate" and source == "docx" and target in ("docx", ""):
            return DocxValidator().convert(context)

        if action == "validate" and source == "markdown" and target in ("markdown", ""):
            return MarkdownValidator().convert(context)

        msg = f"No handler for action='{action}', source='{source}', target='{target}'"
        return ConversionResult(
            task_id=context.request.request_id,
            success=False,
            error=ConversionErrorInfo(
                error_type="unsupported_route",
                message=msg,
                diagnostic_code="PROOFREAD-UNSUPPORTED-ROUTE",
            ),
            diagnostics=[
                ConversionDiagnostic(
                    level="error",
                    message=msg,
                    code="PROOFREAD-UNSUPPORTED-ROUTE",
                )
            ],
        )

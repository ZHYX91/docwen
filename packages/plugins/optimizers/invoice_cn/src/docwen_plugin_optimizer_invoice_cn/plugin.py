"""InvoicePlugin — entry point for docwen_plugin_optimizer_invoice_cn."""

from __future__ import annotations

from typing import TYPE_CHECKING

from docwen_plugin_optimizer_invoice_cn.invoice_cn.converter import InvoiceCnConverter
from docwen_plugin_optimizer_invoice_cn.manifest import build_manifest

if TYPE_CHECKING:
    from docwen_core.models.manifest import PluginManifest
    from docwen_core.models.result import ConversionResult
    from docwen_core.protocols.execution_context import PluginExecutionContext


class InvoicePlugin:
    """Plugin for Chinese invoice (PDF/OFD) to structured Markdown conversion.

    Converts Chinese invoice files with:
    - YAML frontmatter (20 metadata fields)
    - Markdown detail-line table
    - OCR parsing for image-based invoices
    """

    plugin_id: str
    _manifest: PluginManifest | None

    def __init__(self) -> None:
        self.plugin_id = "docwen_plugin_optimizer_invoice_cn"
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
        """Dispatch to the invoice converter."""
        from docwen_core.models.result import (
            ConversionDiagnostic,
            ConversionErrorInfo,
            ConversionResult,
        )

        source = context.request.input_refs[0].format if context.request.input_refs else ""
        target = context.request.target_format
        action = context.request.action_name

        # ── Invoice CN route ─────────────────────────────────────────
        if action == "invoice_cn" and target == "md":
            return InvoiceCnConverter().convert(context)

        # ── Fallback ─────────────────────────────────────────────────
        msg = f"No handler for action='{action}', source='{source}', target='{target}'"
        return ConversionResult(
            task_id=context.request.request_id,
            success=False,
            error=ConversionErrorInfo(
                error_type="unsupported_route",
                message=msg,
                diagnostic_code="INVOICE-UNSUPPORTED-ROUTE",
            ),
            diagnostics=[
                ConversionDiagnostic(
                    level="error",
                    message=msg,
                    code="INVOICE-UNSUPPORTED-ROUTE",
                )
            ],
        )

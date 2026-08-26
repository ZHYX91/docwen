"""SmartSheetConverter — spreadsheet format interconversion via xlsx hub."""

from __future__ import annotations

import os
import shutil
import time
import uuid
from pathlib import Path
from typing import TYPE_CHECKING, Any

from docwen_core.models.artifact import ArtifactManifest
from docwen_core.models.file_ref import FileRef
from docwen_core.models.request import ConversionRequest
from docwen_core.models.result import (
    ConversionDiagnostic,
    ConversionErrorInfo,
    ConversionMetrics,
    ConversionResult,
)
from docwen_core.office_bridge import BridgeCandidate, BridgeResult, convert_with_backend_priority
from docwen_core.protocols.hub_context import HubConversionContext, HubWorkspaceHandle
from docwen_plugin_spreadsheet.format_conversion.legacy_xls_limits import (
    LEGACY_XLS_MAX_COLUMNS,
    LEGACY_XLS_MAX_ROWS,
    LegacyXlsInspectionError,
    LegacyXlsLimitInspection,
    inspect_legacy_xls_limits,
)
from docwen_plugin_spreadsheet.format_conversion.ods_grid_compaction import (
    OdsGridCompaction,
    OdsGridCompactionError,
    compact_generated_ods_grid,
)
from docwen_plugin_spreadsheet.format_conversion.xlsx_ods_policy import (
    XlsxOdsPolicyError,
    XlsxOdsPreparation,
    prepare_xlsx_for_ods,
    validate_prepared_ods,
)
from docwen_plugin_spreadsheet.to_markdown.converter import SpreadsheetToMarkdownConverter

if TYPE_CHECKING:
    from docwen_core.protocols.execution_context import ConverterContext


_MEDIA_TYPES: dict[str, str] = {
    "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "xls": "application/vnd.ms-excel",
    "et": "application/vnd.ms-excel",
    "ods": "application/vnd.oasis.opendocument.spreadsheet",
    "csv": "text/csv",
    "tsv": "text/tab-separated-values",
    "md": "text/markdown",
}
_DEFAULT_SPREADSHEET_PRIORITY = ("wps_spreadsheets", "msoffice_excel", "libreoffice")
_DEFAULT_ODS_PRIORITY = ("msoffice_excel", "libreoffice")
_SPREADSHEET_OFFICE_CONVERSION_TIMEOUT_S = 300.0


def _excel_candidates(save_format: int, *, source_format: str, target_format: str) -> dict[str, BridgeCandidate]:
    candidates = {
        "wps_spreadsheets": BridgeCandidate("WPS Spreadsheets", "Ket.Application", save_format, "excel"),
        "msoffice_excel": BridgeCandidate("Microsoft Excel", "Excel.Application", save_format, "excel"),
    }
    if source_format == "ods" or target_format == "ods":
        return {"msoffice_excel": candidates["msoffice_excel"]}
    return candidates


def _configured_priority(context: ConverterContext, *, source_format: str, target_format: str) -> list[str]:
    uses_ods = source_format == "ods" or target_format == "ods"
    if uses_ods:
        key = "software.special_conversions.ods"
        default = _DEFAULT_ODS_PRIORITY
    else:
        key = "software.default_priority.spreadsheet_processors"
        default = _DEFAULT_SPREADSHEET_PRIORITY
    configured = context.config.get(key, list(default))
    if isinstance(configured, (list, tuple)):
        return [str(item) for item in configured if isinstance(item, str)]
    return list(default)


def _bridge_target_spec(target_format: str) -> tuple[str | None, int | None]:
    if target_format == "xlsx":
        return "xlsx", 51
    if target_format == "xls":
        return "xls", 56
    if target_format == "et":
        return "xls", 56
    if target_format == "ods":
        return "ods", 60
    return None, None


def _write_delimited_from_xlsx(input_path: str, output_path: str, *, sep: str) -> None:
    import pandas as pd

    df = pd.read_excel(input_path, sheet_name=0, header=None)
    df.to_csv(output_path, index=False, header=False, sep=sep, encoding="utf-8")


class SmartSheetConverter:
    """Spreadsheet format interconversion using the xlsx hub."""

    def convert(self, context: ConverterContext) -> Any:
        task_id = context.request.request_id
        if not context.request.input_refs:
            return self._error(
                task_id,
                "invalid_input",
                "SHEETFMT-NO-INPUT",
                "SmartSheetConverter requires one input file.",
            )

        started_at = time.monotonic()
        source = context.request.input_refs[0].format
        target = context.request.target_format
        input_path = context.workspace.input_path
        context.progress.report_progress(0.0, f"Starting {source.upper()}→{target.upper()} conversion")
        context.cancellation.check()

        if source == target:
            output_path = context.workspace.create_artifact_path("primary", f".{target}")
            shutil.copy2(input_path, output_path)
            return self._success(context, source, target, output_path, "copy", started_at)

        try:
            hub_xlsx, inbound_backend = self._prepare_hub_xlsx(context, input_path, source)
        except RuntimeError as exc:
            return self._error(task_id, "dependency_missing", "SHEETFMT-BACKEND", str(exc))

        context.progress.report_progress(55.0, "Intermediate XLSX ready")

        if target == "xlsx":
            return self._success(context, source, target, hub_xlsx, inbound_backend, started_at)

        if target == "md":
            proxy_request = ConversionRequest(
                request_id=task_id,
                input_refs=[
                    FileRef(
                        path=hub_xlsx,
                        format="xlsx",
                        category="spreadsheet",
                        metadata=dict(context.request.input_refs[0].metadata),
                    )
                ],
                target_format="md",
                action_name=context.request.action_name,
                options=dict(context.request.options),
                output_policy=context.request.output_policy,
                config_snapshot=dict(context.request.config_snapshot),
            )
            proxy_context = HubConversionContext(
                base=context,
                request=proxy_request,
                workspace=HubWorkspaceHandle(context.workspace, hub_xlsx),
            )
            proxy_context.logger.info(f"{source}->md preprocessed via {inbound_backend}")
            return SpreadsheetToMarkdownConverter().convert(proxy_context)

        if target == "csv":
            return self._convert_hub_xlsx_to_csv(
                context,
                hub_xlsx,
                source,
                inbound_backend or "pandas",
                started_at,
            )

        if target == "tsv":
            output_path = context.workspace.create_artifact_path("primary", ".tsv")
            _write_delimited_from_xlsx(hub_xlsx, output_path, sep="\t")
            return self._success(context, source, target, output_path, inbound_backend or "pandas", started_at)

        policy_preparation: XlsxOdsPreparation | None = None
        if source == "xlsx" and target == "ods":
            policy_path = context.workspace.create_artifact_path("auxiliary", ".xlsx")
            raw_password = context.request.options.get("spreadsheet_password")
            password = raw_password if isinstance(raw_password, str) else None
            try:
                policy_preparation = prepare_xlsx_for_ods(
                    hub_xlsx,
                    policy_path,
                    password=password,
                    allow_protection_loss=context.request.options.get("allow_spreadsheet_protection_loss") is True,
                )
            except XlsxOdsPolicyError as exc:
                return self._error(task_id, "invalid_input", exc.diagnostic_code, str(exc))
            hub_xlsx = policy_preparation.output_path

        legacy_xls_limits: LegacyXlsLimitInspection | None = None
        if target == "xls":
            try:
                legacy_xls_limits = inspect_legacy_xls_limits(hub_xlsx)
            except LegacyXlsInspectionError as exc:
                return self._error(task_id, "invalid_input", "INVALID_XLSX_PACKAGE", str(exc))

        output_path = context.workspace.create_artifact_path("primary", f".{target}")
        outbound = self._convert_xlsx_to_target(context, hub_xlsx, target, output_path, cancel=context.cancellation)
        ods_legacy_fallback = False
        ods_grid_compaction: OdsGridCompaction | None = None
        ods_grid_compaction_error: str | None = None
        if (
            target == "ods"
            and policy_preparation is not None
            and not outbound.success
            and outbound.available_backend_ids
            and outbound.message != "cancelled"
        ):
            try:
                legacy_xls_limits = inspect_legacy_xls_limits(hub_xlsx)
            except LegacyXlsInspectionError as exc:
                return self._error(task_id, "invalid_input", "INVALID_XLSX_PACKAGE", str(exc))
            fallback, ods_grid_compaction, ods_grid_compaction_error = self._convert_ods_via_legacy_xls(
                context,
                hub_xlsx,
                output_path,
                direct_failure=outbound,
            )
            if fallback.success and fallback.output_path:
                outbound = fallback
                ods_legacy_fallback = True
            else:
                outbound = fallback
        if not outbound.success or not outbound.output_path:
            if policy_preparation is not None and outbound.available_backend_ids:
                return self._error(
                    task_id,
                    "conversion_failed",
                    "SHEETFMT-BACKEND-FAILED",
                    outbound.message,
                )
            return self._error(task_id, "dependency_missing", "SHEETFMT-BACKEND", outbound.message)
        if policy_preparation is not None and (
            policy_preparation.external_links_flattened or policy_preparation.protection_removed or ods_legacy_fallback
        ):
            try:
                validate_prepared_ods(outbound.output_path, policy_preparation)
            except XlsxOdsPolicyError as exc:
                return self._error(task_id, "conversion_failed", exc.diagnostic_code, str(exc))
        backend = outbound.backend
        if inbound_backend and source != "xlsx":
            backend = f"{inbound_backend} -> {backend}"
        return self._success(
            context,
            source,
            target,
            outbound.output_path,
            backend,
            started_at,
            policy_preparation=policy_preparation,
            legacy_xls_limits=legacy_xls_limits,
            ods_legacy_fallback=ods_legacy_fallback,
            ods_grid_compaction=ods_grid_compaction,
            ods_grid_compaction_error=ods_grid_compaction_error,
        )

    def _convert_hub_xlsx_to_csv(
        self,
        context: ConverterContext,
        hub_xlsx: str,
        source: str,
        backend: str,
        started_at: float,
    ) -> ConversionResult:
        """Export every hub worksheet through the canonical multi-CSV converter."""
        from docwen_plugin_spreadsheet.csv_xlsx.converter import XlsxToCsvConverter

        task_id = context.request.request_id
        proxy_request = ConversionRequest(
            request_id=task_id,
            input_refs=[
                FileRef(
                    path=hub_xlsx,
                    format="xlsx",
                    category="spreadsheet",
                    metadata=dict(context.request.input_refs[0].metadata),
                )
            ],
            target_format="csv",
            action_name=context.request.action_name,
            options=dict(context.request.options),
            output_policy=context.request.output_policy,
            config_snapshot=dict(context.request.config_snapshot),
        )
        proxy_context = HubConversionContext(
            base=context,
            request=proxy_request,
            workspace=HubWorkspaceHandle(context.workspace, hub_xlsx),
        )
        downstream = XlsxToCsvConverter().convert(
            proxy_context,
            suggested_stem=Path(context.workspace.input_path).stem,
            metadata_base={
                "source_format": source,
                "target_format": "csv",
                "backend": backend,
            },
            progress_start=55.0,
        )
        if not downstream.success:
            return downstream

        for artifact in downstream.artifacts:
            context.progress.report_artifact_ready(artifact.artifact_id, artifact.suggested_name)

        sheet_count = len(downstream.artifacts)
        total_rows = int(downstream.metrics.extra.get("total_rows", 0))
        output_bytes = sum(
            os.path.getsize(artifact.staging_path)
            for artifact in downstream.artifacts
            if os.path.isfile(artifact.staging_path)
        )
        return ConversionResult(
            task_id=task_id,
            success=True,
            artifacts=downstream.artifacts,
            diagnostics=[
                ConversionDiagnostic(
                    level="info",
                    message=f"Converted {source.upper()} to CSV via {backend}: {sheet_count} sheets.",
                    code="SHEETFMT-OK",
                )
            ],
            metrics=ConversionMetrics(
                duration_ms=(time.monotonic() - started_at) * 1000.0,
                input_bytes=os.path.getsize(context.workspace.input_path)
                if os.path.exists(context.workspace.input_path)
                else 0,
                output_bytes=output_bytes,
                extra={
                    "backend": backend,
                    "sheet_count": sheet_count,
                    "total_rows": total_rows,
                },
            ),
        )

    def _prepare_hub_xlsx(
        self,
        context: ConverterContext,
        input_path: str,
        source: str,
    ) -> tuple[str, str]:
        if source == "xlsx":
            return input_path, "native"
        if source == "csv":
            hub_path = context.workspace.create_artifact_path("auxiliary", ".xlsx")
            self._write_delimited_hub(context, input_path, hub_path, sep=",")
            return hub_path, "openpyxl"
        if source == "tsv":
            hub_path = context.workspace.create_artifact_path("auxiliary", ".xlsx")
            self._write_delimited_hub(context, input_path, hub_path, sep="\t")
            return hub_path, "openpyxl"

        hub_path = context.workspace.create_artifact_path("auxiliary", ".xlsx")
        result = self._convert_binary_to_xlsx(context, input_path, source, hub_path, cancel=context.cancellation)
        if not result.success or not result.output_path:
            raise RuntimeError(result.message or f"Failed to convert {source} to xlsx.")
        return result.output_path, result.backend

    @staticmethod
    def _write_delimited_hub(
        context: ConverterContext,
        input_path: str,
        output_path: str,
        *,
        sep: str,
    ) -> None:
        from docwen_plugin_spreadsheet.csv_xlsx.converter import _build_delimited_workbook

        workbook, _row_count = _build_delimited_workbook(
            input_path,
            sep=sep,
            cancel_check=context.cancellation.check,
        )
        try:
            workbook.save(output_path)
        finally:
            workbook.close()

    def _convert_binary_to_xlsx(
        self,
        context: ConverterContext,
        input_path: str,
        source_format: str,
        output_path: str,
        *,
        cancel: object | None = None,
    ):
        libreoffice_format, save_format = _bridge_target_spec("xlsx")
        return convert_with_backend_priority(
            input_path,
            output_path,
            source_format=source_format,
            backend_priority=_configured_priority(context, source_format=source_format, target_format="xlsx"),
            com_candidates=_excel_candidates(save_format or 51, source_format=source_format, target_format="xlsx"),
            libreoffice_format=libreoffice_format,
            cancel=cancel,
            com_timeout_s=_SPREADSHEET_OFFICE_CONVERSION_TIMEOUT_S,
            libreoffice_timeout_s=_SPREADSHEET_OFFICE_CONVERSION_TIMEOUT_S,
            failure_subject="Configured spreadsheet bridge backends",
        )

    def _convert_xlsx_to_target(
        self,
        context: ConverterContext,
        input_path: str,
        target: str,
        output_path: str,
        *,
        cancel: object | None = None,
    ):
        libreoffice_format, save_format = _bridge_target_spec(target)
        actual_output = output_path
        if target == "et":
            actual_output = str(Path(output_path).with_suffix(".xls"))
        result = convert_with_backend_priority(
            input_path,
            actual_output,
            source_format="xlsx",
            backend_priority=_configured_priority(context, source_format="xlsx", target_format=target),
            com_candidates=_excel_candidates(save_format or 56, source_format="xlsx", target_format=target),
            libreoffice_format=libreoffice_format,
            cancel=cancel,
            com_timeout_s=_SPREADSHEET_OFFICE_CONVERSION_TIMEOUT_S,
            libreoffice_timeout_s=_SPREADSHEET_OFFICE_CONVERSION_TIMEOUT_S,
            failure_subject="Configured spreadsheet bridge backends",
        )
        if result.success and result.output_path and target == "et":
            Path(result.output_path).replace(output_path)
            result.output_path = output_path
        return result

    def _convert_ods_via_legacy_xls(
        self,
        context: ConverterContext,
        input_path: str,
        output_path: str,
        *,
        direct_failure: BridgeResult,
    ) -> tuple[BridgeResult, OdsGridCompaction | None, str | None]:
        """Deliver a best-effort ODS after direct backends reject a complex XLSX."""

        legacy_xls = context.workspace.create_artifact_path("auxiliary", ".xls")
        legacy = self._convert_xlsx_to_target(
            context,
            input_path,
            "xls",
            legacy_xls,
            cancel=context.cancellation,
        )
        if not legacy.success or not legacy.output_path:
            return self._fallback_failure(direct_failure, "XLSX-to-XLS", legacy), None, None

        roundtrip_xlsx = context.workspace.create_artifact_path("auxiliary", ".xlsx")
        roundtrip = self._convert_binary_to_xlsx(
            context,
            legacy.output_path,
            "xls",
            roundtrip_xlsx,
            cancel=context.cancellation,
        )
        if not roundtrip.success or not roundtrip.output_path:
            return self._fallback_failure(direct_failure, "XLS-to-XLSX", roundtrip), None, None

        Path(output_path).unlink(missing_ok=True)
        delivered = self._convert_xlsx_to_target(
            context,
            roundtrip.output_path,
            "ods",
            output_path,
            cancel=context.cancellation,
        )
        if not delivered.success or not delivered.output_path:
            return (
                self._fallback_failure(direct_failure, "round-tripped XLSX-to-ODS", delivered),
                None,
                None,
            )

        compaction: OdsGridCompaction | None = None
        compaction_error: str | None = None
        try:
            compaction = compact_generated_ods_grid(
                delivered.output_path,
                roundtrip.output_path,
            )
        except OdsGridCompactionError as exc:
            compaction_error = str(exc)

        return (
            BridgeResult(
                True,
                output_path=delivered.output_path,
                backend=f"{legacy.backend} -> {roundtrip.backend} -> {delivered.backend}",
                attempted_backend_ids=tuple(
                    dict.fromkeys(
                        (
                            *direct_failure.attempted_backend_ids,
                            *legacy.attempted_backend_ids,
                            *roundtrip.attempted_backend_ids,
                            *delivered.attempted_backend_ids,
                        )
                    )
                ),
                available_backend_ids=tuple(
                    dict.fromkeys(
                        (
                            *direct_failure.available_backend_ids,
                            *legacy.available_backend_ids,
                            *roundtrip.available_backend_ids,
                            *delivered.available_backend_ids,
                        )
                    )
                ),
            ),
            compaction,
            compaction_error,
        )

    @staticmethod
    def _fallback_failure(
        direct_failure: BridgeResult,
        fallback_step: str,
        fallback: BridgeResult,
    ) -> BridgeResult:
        fallback_message = fallback.message or "no artifact was produced"
        return BridgeResult(
            False,
            message=(
                f"{direct_failure.message} Best-effort legacy XLS fallback also failed "
                f"during {fallback_step}: {fallback_message}"
            ),
            attempted_backend_ids=tuple(
                dict.fromkeys((*direct_failure.attempted_backend_ids, *fallback.attempted_backend_ids))
            ),
            available_backend_ids=tuple(
                dict.fromkeys((*direct_failure.available_backend_ids, *fallback.available_backend_ids))
            ),
        )

    def _success(
        self,
        context: ConverterContext,
        source: str,
        target: str,
        output_path: str,
        backend: str,
        started_at: float,
        *,
        policy_preparation: XlsxOdsPreparation | None = None,
        legacy_xls_limits: LegacyXlsLimitInspection | None = None,
        ods_legacy_fallback: bool = False,
        ods_grid_compaction: OdsGridCompaction | None = None,
        ods_grid_compaction_error: str | None = None,
    ) -> ConversionResult:
        metadata: dict[str, Any] = {
            "source_format": source,
            "target_format": target,
            "backend": backend,
        }
        if legacy_xls_limits is not None and legacy_xls_limits.has_truncation_risk:
            metadata["legacy_xls_out_of_bounds_cells"] = legacy_xls_limits.out_of_bounds_cell_count
            metadata["legacy_xls_out_of_bounds_formulas"] = legacy_xls_limits.out_of_bounds_formula_count
        if ods_legacy_fallback:
            metadata["ods_legacy_xls_fallback"] = True
        if ods_grid_compaction is not None and ods_grid_compaction.changed:
            metadata["ods_grid_compacted_sheets"] = len(ods_grid_compaction.trimmed_sheets)
            metadata["ods_grid_removed_repeated_rows"] = ods_grid_compaction.removed_repeated_rows
            metadata["ods_grid_removed_repeated_columns"] = ods_grid_compaction.removed_repeated_columns
        if ods_grid_compaction_error is not None:
            metadata["ods_grid_compaction_failed"] = True
        if policy_preparation is not None and policy_preparation.removed_external_defined_names:
            metadata["removed_external_defined_names"] = len(policy_preparation.removed_external_defined_names)
        if policy_preparation is not None and policy_preparation.fidelity_risk_counts:
            metadata["ods_fidelity_risk_counts"] = dict(policy_preparation.fidelity_risk_counts)
        artifact = ArtifactManifest(
            artifact_id=str(uuid.uuid4()),
            kind="primary",
            staging_path=output_path,
            suggested_name=f"{Path(context.workspace.input_path).stem}.{target}",
            media_type=_MEDIA_TYPES.get(target, "application/octet-stream"),
            metadata=metadata,
            is_primary=True,
        )
        context.workspace.add_artifact(artifact)
        context.progress.report_artifact_ready(artifact.artifact_id, artifact.suggested_name)
        context.progress.report_progress(100.0, "Spreadsheet conversion complete")
        diagnostics = [
            ConversionDiagnostic(
                level="info",
                message=f"Converted {source.upper()} to {target.upper()} via {backend}.",
                code="SHEETFMT-OK",
            )
        ]
        if policy_preparation is not None and policy_preparation.external_links_flattened:
            removed_names = len(policy_preparation.removed_external_defined_names)
            removed_names_text = (
                ""
                if removed_names == 0
                else (
                    f" {removed_names} external defined name"
                    f"{'' if removed_names == 1 else 's'} without a safe cached value "
                    f"{'was' if removed_names == 1 else 'were'} removed."
                )
            )
            diagnostics.append(
                ConversionDiagnostic(
                    level="warning",
                    message=(
                        "External workbook formulas with package caches were replaced with static values, "
                        "and the live link graph was removed; the ODS result will not receive future "
                        f"link updates.{removed_names_text}"
                    ),
                    code="EXTERNAL_LINK_FLATTENED",
                )
            )
        if policy_preparation is not None and policy_preparation.protection_removed:
            diagnostics.append(
                ConversionDiagnostic(
                    level="warning",
                    message=(
                        "Workbook or sheet password protection was removed from the private "
                        "conversion copy and is not present in the ODS result."
                    ),
                    code="PROTECTION_REMOVED_FOR_TARGET",
                )
            )
        if policy_preparation is not None and policy_preparation.fidelity_risk_counts:
            risk_labels = {
                "data_validations": "data validations",
                "conditional_formatting_ranges": "conditional-formatting ranges",
                "charts": "chart parts",
                "drawings": "drawing parts",
                "tables": "table parts",
                "pivot_or_slicer_parts": "pivot/slicer parts",
                "defined_names": "defined names",
            }
            detected = ", ".join(
                f"{count:,} {risk_labels.get(name, name.replace('_', ' '))}"
                for name, count in policy_preparation.fidelity_risk_counts
            )
            diagnostics.append(
                ConversionDiagnostic(
                    level="warning",
                    message=(
                        "The ODS file was delivered as a best-effort result, but the source contains "
                        f"complex spreadsheet features ({detected}). ODS conversion may move or alter "
                        "charts and drawings, drop validation, formatting, names, or protection, and "
                        "materialize hidden cache worksheets; converting the result back to XLSX may "
                        "differ further. The original source was not changed. Review critical sheets "
                        "before relying on the result."
                    ),
                    code="ODS_FEATURE_FIDELITY_RISK",
                )
            )
        if legacy_xls_limits is not None and legacy_xls_limits.has_truncation_risk:
            affected_names = [sheet.name for sheet in legacy_xls_limits.affected_sheets]
            displayed_names = ", ".join(affected_names[:8])
            if len(affected_names) > 8:
                displayed_names = f"{displayed_names}, and {len(affected_names) - 8} more"
            diagnostics.append(
                ConversionDiagnostic(
                    level="warning",
                    message=(
                        f"Legacy XLS supports at most {LEGACY_XLS_MAX_ROWS:,} rows and "
                        f"{LEGACY_XLS_MAX_COLUMNS:,} columns per sheet. "
                        f"{legacy_xls_limits.out_of_bounds_cell_count:,} populated cell"
                        f"{'' if legacy_xls_limits.out_of_bounds_cell_count == 1 else 's'} "
                        f"including {legacy_xls_limits.out_of_bounds_formula_count:,} formula cell"
                        f"{'' if legacy_xls_limits.out_of_bounds_formula_count == 1 else 's'} "
                        f"lie outside that grid and may be truncated ({displayed_names})."
                    ),
                    code="LEGACY_XLS_LIMIT_TRUNCATION",
                )
            )
        if ods_legacy_fallback:
            diagnostics.append(
                ConversionDiagnostic(
                    level="warning",
                    message=(
                        "Direct XLSX-to-ODS conversion failed, so DocWen delivered a best-effort "
                        "ODS by round-tripping the policy-prepared private copy through legacy XLS. "
                        "This fallback can change formulas or cached results and can lose formatting, "
                        "validation, conditional formatting, names, links, protection, drawings, "
                        "charts, and print settings."
                    ),
                    code="ODS_LEGACY_XLS_FALLBACK",
                )
            )
        if ods_grid_compaction is not None and ods_grid_compaction.changed:
            diagnostics.append(
                ConversionDiagnostic(
                    level="info",
                    message=(
                        "Repeated empty ODS grid spans beyond the legacy round-trip worksheet "
                        f"dimensions were compacted on {len(ods_grid_compaction.trimmed_sheets)} "
                        "sheet(s); cells inside those dimensions were not changed."
                    ),
                    code="ODS_UNUSED_GRID_COMPACTED",
                )
            )
        if ods_grid_compaction_error is not None:
            diagnostics.append(
                ConversionDiagnostic(
                    level="warning",
                    message=(
                        "The best-effort ODS was delivered, but its repeated empty grid spans "
                        "could not be compacted. Some downstream spreadsheet applications may "
                        "materialize a much larger used range."
                    ),
                    code="ODS_GRID_COMPACTION_NOT_APPLIED",
                )
            )
        metrics_extra: dict[str, Any] = {"backend": backend}
        if legacy_xls_limits is not None and legacy_xls_limits.has_truncation_risk:
            metrics_extra["legacy_xls_out_of_bounds_cells"] = legacy_xls_limits.out_of_bounds_cell_count
            metrics_extra["legacy_xls_out_of_bounds_formulas"] = legacy_xls_limits.out_of_bounds_formula_count
        if ods_legacy_fallback:
            metrics_extra["ods_legacy_xls_fallback"] = True
        if ods_grid_compaction is not None and ods_grid_compaction.changed:
            metrics_extra["ods_grid_compacted_sheets"] = len(ods_grid_compaction.trimmed_sheets)
            metrics_extra["ods_grid_removed_repeated_rows"] = ods_grid_compaction.removed_repeated_rows
            metrics_extra["ods_grid_removed_repeated_columns"] = ods_grid_compaction.removed_repeated_columns
        if ods_grid_compaction_error is not None:
            metrics_extra["ods_grid_compaction_failed"] = True
        if policy_preparation is not None and policy_preparation.removed_external_defined_names:
            metrics_extra["removed_external_defined_names"] = len(policy_preparation.removed_external_defined_names)
        if policy_preparation is not None and policy_preparation.fidelity_risk_counts:
            metrics_extra["ods_fidelity_risk_counts"] = dict(policy_preparation.fidelity_risk_counts)
        return ConversionResult(
            task_id=context.request.request_id,
            success=True,
            artifacts=[artifact],
            diagnostics=diagnostics,
            metrics=ConversionMetrics(
                duration_ms=(time.monotonic() - started_at) * 1000.0,
                input_bytes=os.path.getsize(context.workspace.input_path)
                if os.path.exists(context.workspace.input_path)
                else 0,
                output_bytes=os.path.getsize(output_path) if os.path.exists(output_path) else 0,
                extra=metrics_extra,
            ),
        )

    def _error(self, task_id: str, error_type: str, code: str, message: str) -> ConversionResult:
        return ConversionResult(
            task_id=task_id,
            success=False,
            error=ConversionErrorInfo(
                error_type=error_type,
                message=message,
                diagnostic_code=code,
                recoverable=True,
            ),
            diagnostics=[ConversionDiagnostic(level="error", message=message, code=code)],
        )

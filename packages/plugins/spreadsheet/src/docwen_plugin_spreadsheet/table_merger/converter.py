"""Table merger converter (ACT-MERGE-TABLES).

Merges multiple spreadsheet files into one base file.
Supports three merge modes: by row, by column, and by cell.

Core algorithm adapted from old ``converter/table_merger/core.py``.
"""

from __future__ import annotations

import math
import os
import uuid
from collections.abc import Callable
from pathlib import Path
from typing import TYPE_CHECKING, Any

from docwen_core.errors import CancellationRequested
from docwen_core.spreadsheet_projection import SpreadsheetProjectionLimitError
from docwen_plugin_spreadsheet.table_merger.inputs import MissingFormulaCacheError, load_merge_workbook
from docwen_plugin_spreadsheet.table_merger.records import check_extent, insert_sparse, merge_records, write_scalar
from docwen_plugin_spreadsheet.table_merger.semantics import MergeSummary, value_identity
from docwen_plugin_spreadsheet.table_merger.text_limits import MergeCellTextTooLongError

if TYPE_CHECKING:
    from docwen_core.protocols.execution_context import ConverterContext


class TableMergerConverter:
    """Merge multiple table files into a single XLSX output.

    Supports:
    - MODE_BY_ROW (1): Insert rows from collection files into the base.
    - MODE_BY_COLUMN (2): Insert columns from collection files into the base.
    - MODE_BY_CELL (3): Add values from collection cells to base cells.

    Uses a sliding-window alignment algorithm to find the best
    row/column offset between files.
    """

    MODE_BY_ROW = 1
    MODE_BY_COLUMN = 2
    MODE_BY_CELL = 3

    OFFSET_RANGE = 10
    MAX_CELLS_FOR_ALIGNMENT = 500

    _MODE_MAP = {"row": 1, "col": 2, "cell": 3}  # noqa: RUF012

    def convert(self, context: ConverterContext) -> Any:
        """Run the table merge operation.

        The first input is the base; every remaining input is a collection.
        Only active worksheets are merged. Other base sheets are retained as
        cached values, while other collection sheets are not collected.

        Args:
            context: The plugin execution context.

        Returns:
            ``ConversionResult`` with the merged XLSX artifact.
        """
        from docwen_core.models.artifact import ArtifactManifest
        from docwen_core.models.result import (
            ConversionDiagnostic,
            ConversionErrorInfo,
            ConversionMetrics,
            ConversionResult,
        )

        task_id = context.request.request_id
        input_refs = context.request.input_refs
        options = context.request.options

        if len(input_refs) < 2:
            return ConversionResult(
                task_id=task_id,
                success=False,
                error=ConversionErrorInfo(
                    error_type="invalid_input",
                    message="Table merge requires at least 2 input files (base + 1+ collection files).",
                    diagnostic_code="MERGE-NEED-MORE-FILES",
                ),
            )

        source_paths = [ref.path for ref in input_refs]
        source_formats = [str(ref.format or "").strip().lower() for ref in input_refs]
        base_source_path = source_paths[0]
        base_path = base_source_path
        collect_paths = source_paths[1:]

        mode_str = options.get("merge_mode", "cell")
        mode = self._MODE_MAP.get(mode_str, self.MODE_BY_CELL)
        offset_range = options.get("offset_range", self.OFFSET_RANGE)

        context.cancellation.check()
        context.progress.report_progress(0.0, f"Starting table merge ({mode_str} mode)")
        context.logger.info(
            f"Table merge: base={os.path.basename(base_path)}, collect={len(collect_paths)} files, mode={mode_str}"
        )

        wb = None
        summaries: list[MergeSummary] = []
        merged_sources = [Path(base_source_path).name]
        try:
            from docwen_plugin_spreadsheet.format_conversion.converter import SmartSheetConverter

            prepared_paths: list[str] = []
            smart_converter = SmartSheetConverter()
            for source_path, source_format in zip(source_paths, source_formats, strict=True):
                context.cancellation.check()
                if not source_format or source_format == "spreadsheet":
                    raise RuntimeError(f"Missing concrete admitted spreadsheet format for '{Path(source_path).name}'.")
                if source_format in {"xlsx", "xlsm", "xltx", "xltm"}:
                    prepared_paths.append(source_path)
                    continue
                prepared_path, backend = smart_converter._prepare_hub_xlsx(
                    context,
                    source_path,
                    source_format,
                )
                context.logger.info(f"  Prepared {Path(source_path).name} as XLSX with {backend}")
                prepared_paths.append(prepared_path)

            base_path = prepared_paths[0]
            collect_paths = prepared_paths[1:]

            try:
                wb = load_merge_workbook(
                    base_path, base=True, cancel_check=context.cancellation.check, source_name=base_source_path
                )
            except (MissingFormulaCacheError, SpreadsheetProjectionLimitError, CancellationRequested):
                raise
            except Exception as exc:
                raise RuntimeError(f"Failed to merge '{Path(base_path).name}': {exc}") from exc
            base_ws = wb.active
            if base_ws is None:
                raise RuntimeError(f"Failed to merge '{Path(base_path).name}': base file has no active worksheet")

            # Unmerge base cells
            self._unmerge_all_cells(base_ws)

            for i, collect_file in enumerate(collect_paths):
                context.cancellation.check()
                merged_sources.append(Path(source_paths[i + 1]).name)
                progress = 50.0 * (i / max(len(collect_paths), 1))
                context.progress.report_progress(progress, f"Merging file {i + 1}/{len(collect_paths)}")
                context.logger.info(f"  Merging: {os.path.basename(collect_file)}")

                try:
                    collect_wb = load_merge_workbook(
                        collect_file,
                        base=False,
                        cancel_check=context.cancellation.check,
                        source_name=source_paths[i + 1],
                    )
                except (MissingFormulaCacheError, SpreadsheetProjectionLimitError, CancellationRequested):
                    raise
                except Exception as exc:
                    raise RuntimeError(f"Failed to merge '{Path(collect_file).name}': {exc}") from exc
                try:
                    collect_ws = collect_wb.active
                    if collect_ws is None:
                        raise RuntimeError(f"Failed to merge '{Path(collect_file).name}': no active worksheet")

                    self._unmerge_all_cells(collect_ws)
                    row_offset, col_offset = self._find_best_offset(
                        base_ws, collect_ws, offset_range, cancel_check=context.cancellation.check
                    )
                    # Make room for leading collection fields instead of
                    # clipping them at worksheet row/column zero.
                    if row_offset < 0 and mode != self.MODE_BY_ROW:
                        check_extent(base_ws.max_row - row_offset, base_ws.max_column)
                        insert_sparse(
                            base_ws, by_row=True, start=1, amount=-row_offset, cancel_check=context.cancellation.check
                        )
                        row_offset = 0
                    if col_offset < 0 and mode != self.MODE_BY_COLUMN:
                        check_extent(base_ws.max_row, base_ws.max_column - col_offset)
                        insert_sparse(
                            base_ws, by_row=False, start=1, amount=-col_offset, cancel_check=context.cancellation.check
                        )
                        col_offset = 0
                    summary = MergeSummary(
                        Path(source_paths[i + 1]).name, collect_ws.title, mode_str, row_offset, col_offset
                    )
                    if mode == self.MODE_BY_CELL:
                        self._merge_by_cell(
                            base_ws,
                            collect_ws,
                            row_offset,
                            col_offset,
                            summary=summary,
                            cancel_check=context.cancellation.check,
                        )
                    else:
                        by_row = mode == self.MODE_BY_ROW
                        merge_records(
                            base_ws,
                            collect_ws,
                            by_row=by_row,
                            offset=col_offset if by_row else row_offset,
                            summary=summary,
                            cancel_check=context.cancellation.check,
                        )
                    check_extent(base_ws.max_row, base_ws.max_column)
                    summaries.append(summary)
                finally:
                    collect_wb.close()

        except CancellationRequested:
            if wb is not None:
                wb.close()
            raise
        except Exception as exc:
            if wb is not None:
                wb.close()
            code = "MERGE-PARSE-ERROR"
            message = str(exc)
            if isinstance(exc, SpreadsheetProjectionLimitError):
                code = exc.diagnostic_code
                message = f"{exc} Merge inputs: {', '.join(merged_sources)}."
            elif isinstance(exc, MissingFormulaCacheError):
                code = "MERGE-FORMULA-CACHE-MISSING"
            elif isinstance(exc, MergeCellTextTooLongError):
                code = "MERGE-XLSX-CELL-TEXT-TOO-LONG"
                message = f"{exc} Merge inputs: {', '.join(merged_sources)}."
            context.logger.error(f"Table merge processing failed: {message}")
            return ConversionResult(
                task_id=task_id,
                success=False,
                error=ConversionErrorInfo(
                    error_type="conversion_failed",
                    message=message,
                    diagnostic_code=code,
                ),
                diagnostics=[
                    ConversionDiagnostic(
                        level="error",
                        message=f"Table merge processing failed: {message}",
                        code=code,
                        location=(
                            exc.location
                            if isinstance(exc, (MissingFormulaCacheError, MergeCellTextTooLongError))
                            else ""
                        ),
                    ),
                ],
                metrics=ConversionMetrics(
                    extra={"unknown_formula_values": exc.count} if isinstance(exc, MissingFormulaCacheError) else {}
                ),
            )

        # ── Write merged result to staging ────────────────────────────
        assert wb is not None
        context.progress.report_progress(90.0, "Writing merged result...")
        output_path = context.workspace.create_artifact_path("primary", ".xlsx")
        try:
            wb.save(output_path)
        except Exception as exc:
            context.logger.error(f"Table merge write failed: {exc}")
            return ConversionResult(
                task_id=task_id,
                success=False,
                error=ConversionErrorInfo(
                    error_type="conversion_failed",
                    message=f"Failed to write merged XLSX: {exc}",
                    diagnostic_code="MERGE-WRITE-ERROR",
                ),
                diagnostics=[
                    ConversionDiagnostic(
                        level="error",
                        message=f"File write error at {output_path}: {exc}",
                        code="MERGE-WRITE-ERROR",
                    ),
                ],
            )
        finally:
            wb.close()

        # Build artifact
        base_name = Path(base_source_path).stem
        suggested_name = f"{base_name}_merged.xlsx"

        artifact = ArtifactManifest(
            artifact_id=str(uuid.uuid4()),
            kind="primary",
            staging_path=output_path,
            suggested_name=suggested_name,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            metadata={
                "base_file": os.path.basename(base_source_path),
                "collect_count": len(collect_paths),
                "merge_mode": mode_str,
            },
            is_primary=True,
        )
        context.workspace.add_artifact(artifact)

        context.progress.report_artifact_ready(artifact.artifact_id, suggested_name)
        context.progress.report_progress(100.0, "Table merge complete")
        context.logger.info(f"Table merge complete: {len(collect_paths)} files merged")

        return ConversionResult(
            task_id=task_id,
            success=True,
            artifacts=[artifact],
            diagnostics=[
                ConversionDiagnostic(
                    level="info",
                    message=(f"Merged {len(collect_paths)} files into base ({mode_str} mode)"),
                    code="MERGE-OK",
                ),
                *[
                    ConversionDiagnostic(
                        level="info",
                        message=summary.message(),
                        code="MERGE-SUMMARY",
                        location=f"{summary.source_name}: {summary.sheet}",
                    )
                    for summary in summaries
                ],
            ],
            error=None,
            metrics=ConversionMetrics(
                duration_ms=0.0,
                input_bytes=sum(os.path.getsize(p) for p in source_paths if os.path.isfile(p)),
                output_bytes=os.path.getsize(output_path) if os.path.isfile(output_path) else 0,
                extra={
                    "collect_count": len(collect_paths),
                    "mode": mode_str,
                    "merge_summaries": [summary.to_dict() for summary in summaries],
                },
            ),
        )

    # ── Alignment ─────────────────────────────────────────────────────

    def _find_best_offset(
        self,
        base_ws: Any,
        collect_ws: Any,
        offset_range: int,
        *,
        cancel_check: Callable[[], None] | None = None,
    ) -> tuple[int, int]:
        """Find best row/column offset via sliding window alignment."""
        max_overlap = 0
        best_offset = (0, 0)

        collect_cells = list(self._iter_cells_for_alignment(collect_ws))
        base_values = {(cell.row, cell.column): self._cell_identity(cell) for cell in base_ws._cells.values()}

        for row_off in range(-offset_range, offset_range + 1):
            for col_off in range(-offset_range, offset_range + 1):
                if cancel_check is not None:
                    cancel_check()
                overlap = sum(
                    value is not None and base_values.get((row + row_off, col + col_off)) == value
                    for row, col, value in collect_cells
                )
                if overlap > max_overlap:
                    max_overlap = overlap
                    best_offset = (row_off, col_off)

        if max_overlap == 0:
            suggested = self._suggest_offset_by_value_matching(base_ws, collect_ws)
            if suggested is not None:
                return suggested

        return best_offset

    def _iter_cells_for_alignment(self, ws: Any) -> list[tuple[int, int, object]]:
        """Iterate non-empty cells for alignment, up to MAX_CELLS_FOR_ALIGNMENT."""
        result: list[tuple[int, int, object]] = []
        cells = getattr(ws, "_cells", None)
        it = cells.values() if isinstance(cells, dict) else (cell for row in ws.iter_rows() for cell in row)
        for cell in it:
            v = self._cell_identity(cell)
            if v is None:
                continue
            row = getattr(cell, "row", None)
            col = getattr(cell, "column", None)
            if not isinstance(row, int) or not isinstance(col, int):
                continue
            result.append((row, col, v))
            if len(result) >= self.MAX_CELLS_FOR_ALIGNMENT:
                break
        return result

    @staticmethod
    def _cell_identity(cell: Any) -> Any:
        return None if cell is None else value_identity(cell.value, source_type=cell.data_type)

    def _suggest_offset_by_value_matching(
        self,
        base_ws: Any,
        collect_ws: Any,
    ) -> tuple[int, int] | None:
        """Fallback alignment by matching cell values."""
        base_positions: dict[object, list[tuple[int, int]]] = {}
        for r, c, v in self._iter_cells_for_alignment(base_ws):
            base_positions.setdefault(v, []).append((r, c))

        counts: dict[tuple[int, int], int] = {}
        for r, c, v in self._iter_cells_for_alignment(collect_ws):
            positions = base_positions.get(v)
            if not positions:
                continue
            for br, bc in positions[:3]:
                diff = (br - r, bc - c)
                counts[diff] = counts.get(diff, 0) + 1

        if not counts:
            return None
        return max(counts.items(), key=lambda kv: kv[1])[0]

    # ── Unmerge ───────────────────────────────────────────────────────

    @staticmethod
    def _unmerge_all_cells(worksheet: Any) -> None:
        """Unmerge all merged cells and fail the conversion on any error.

        The caller owns an in-memory workbook that is saved only after the
        complete merge succeeds. Propagating an error therefore discards any
        partial in-memory mutation instead of publishing a silently damaged
        result.
        """

        merged_ranges = list(worksheet.merged_cells.ranges)
        for merged_range in merged_ranges:
            top_left = worksheet.cell(row=merged_range.min_row, column=merged_range.min_col)
            merged_value = top_left.value
            merged_type = top_left.data_type
            worksheet.unmerge_cells(str(merged_range))
            for row in range(merged_range.min_row, merged_range.max_row + 1):
                for col in range(merged_range.min_col, merged_range.max_col + 1):
                    cell = worksheet.cell(row=row, column=col)
                    if type(cell).__name__ != "MergedCell":
                        write_scalar(cell, merged_value, merged_type)

    # ── Merge by cell ─────────────────────────────────────────────────

    def _merge_by_cell(
        self,
        base_ws: Any,
        collect_ws: Any,
        row_offset: int,
        col_offset: int,
        *,
        summary: MergeSummary | None = None,
        cancel_check: Callable[[], None] | None = None,
    ) -> None:
        # Normal worksheets keep only instantiated cells in this mapping.
        # iter_rows() would create every blank cell in a sparse bounding box.
        cells = collect_ws._cells
        max_row, max_col = collect_ws.max_row, collect_ws.max_column
        if summary is not None:
            missing = max_row * max_col - len(cells) if cells else 0
            summary.counts["skipped"] += missing
            # Preserve bounded location evidence without visiting the rectangle.
            if missing:
                from openpyxl.utils.cell import get_column_letter

                samples = summary.locations.setdefault("skipped", [])
                next_position = 0
                for row_index, col_index in sorted(cells):
                    position = (row_index - 1) * max_col + col_index - 1
                    while next_position < position and len(samples) < 20:
                        row_number, column_number = divmod(next_position, max_col)
                        samples.append(f"{get_column_letter(column_number + 1)}{row_number + 1}")
                        next_position += 1
                    next_position = position + 1
                    if len(samples) == 20:
                        break
        for coordinate in sorted(cells):
            cell = cells[coordinate]
            if cancel_check is not None:
                cancel_check()
            if cell.value is None or str(cell.value).strip() == "":
                if summary is not None:
                    summary.record("skipped", cell.coordinate)
                continue
            base_row = cell.row + row_offset
            base_col = cell.column + col_offset
            if base_row < 1 or base_col < 1:
                continue
            check_extent(base_row, base_col)
            base_cell = base_ws.cell(base_row, base_col)
            if type(base_cell).__name__ != "MergedCell":
                original = base_cell.value
                value, source_type = self._merge_cell_values(original, cell.value, base_cell.data_type, cell.data_type)
                write_scalar(base_cell, value, source_type)
                if summary is not None:
                    decision = (
                        "added"
                        if value_identity(original) is None
                        else "skipped"
                        if value_identity(original) == value_identity(base_cell.value)
                        else "combined"
                    )
                    summary.record(decision, cell.coordinate)

    @staticmethod
    def _merge_cell_values(
        base_value: Any, collect_value: Any, base_type: str, collect_type: str
    ) -> tuple[Any, str | None]:
        """Merge saved scalars and retain the origin type when one value survives."""
        be = base_value is None or str(base_value).strip() == ""
        ce = collect_value is None or str(collect_value).strip() == ""
        if be and ce:
            return None, None
        if be:
            return collect_value, collect_type
        if ce:
            return base_value, base_type

        base_num = TableMergerConverter._try_number(base_value)
        col_num = TableMergerConverter._try_number(collect_value)
        if base_num is not None and col_num is not None:
            total = base_num + col_num
            try:
                finite = math.isfinite(total)
            except OverflowError:
                finite = False
            if not finite:
                raise ValueError("Cell merge sum exceeds the finite numeric range supported by XLSX.")
            return total, None

        if str(base_value).strip() == str(collect_value).strip():
            return base_value, base_type
        return f"{base_value},{collect_value}", "s"

    @staticmethod
    def _try_number(value: Any) -> float | int | None:
        if isinstance(value, int):
            return value
        if isinstance(value, float):
            return value if math.isfinite(value) else None
        try:
            num = float(str(value).strip())
            if not math.isfinite(num):
                return None
            if num.is_integer():
                return int(num)
            return num
        except (ValueError, AttributeError):
            return None

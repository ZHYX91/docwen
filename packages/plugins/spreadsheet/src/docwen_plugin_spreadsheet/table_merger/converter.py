"""Table merger converter (ACT-MERGE-TABLES).

Merges multiple spreadsheet files into one base file.
Supports three merge modes: by row, by column, and by cell.

Core algorithm adapted from old ``converter/table_merger/core.py``.
"""

from __future__ import annotations

import os
import uuid
from datetime import date, datetime
from pathlib import Path
from typing import TYPE_CHECKING, Any

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

        Expects exactly one input_ref (base file) with collect_files
        provided through options or multiple input_refs (first as base,
        rest as collection files).

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

        try:
            from docwen_plugin_spreadsheet.csv_xlsx.converter import _load_admitted_xlsx
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
                wb = _load_admitted_xlsx(base_path, data_only=True)
            except Exception as exc:
                raise RuntimeError(f"Failed to merge '{Path(base_path).name}': {exc}") from exc
            base_ws = wb.active
            if base_ws is None:
                raise RuntimeError(f"Failed to merge '{Path(base_path).name}': base file has no active worksheet")

            # Unmerge base cells
            self._unmerge_all_cells(base_ws)

            for i, collect_file in enumerate(collect_paths):
                context.cancellation.check()
                progress = 50.0 * (i / max(len(collect_paths), 1))
                context.progress.report_progress(progress, f"Merging file {i + 1}/{len(collect_paths)}")
                context.logger.info(f"  Merging: {os.path.basename(collect_file)}")

                try:
                    collect_wb = _load_admitted_xlsx(collect_file, data_only=True)
                except Exception as exc:
                    raise RuntimeError(f"Failed to merge '{Path(collect_file).name}': {exc}") from exc
                collect_ws = collect_wb.active
                if collect_ws is None:
                    collect_wb.close()
                    continue

                self._unmerge_all_cells(collect_ws)

                # Find best alignment
                row_offset, col_offset = self._find_best_offset(base_ws, collect_ws, offset_range)

                # Execute merge
                if mode == self.MODE_BY_ROW:
                    self._merge_by_row(base_ws, collect_ws, row_offset, col_offset)
                elif mode == self.MODE_BY_COLUMN:
                    self._merge_by_column(base_ws, collect_ws, row_offset, col_offset)
                elif mode == self.MODE_BY_CELL:
                    self._merge_by_cell(base_ws, collect_ws, row_offset, col_offset)

                collect_wb.close()

        except Exception as exc:
            context.logger.error(f"Table merge processing failed: {exc}")
            return ConversionResult(
                task_id=task_id,
                success=False,
                error=ConversionErrorInfo(
                    error_type="conversion_failed",
                    message=str(exc),
                    diagnostic_code="MERGE-PARSE-ERROR",
                ),
                diagnostics=[
                    ConversionDiagnostic(
                        level="error",
                        message=f"Table merge processing failed: {exc}",
                        code="MERGE-PARSE-ERROR",
                    ),
                ],
            )

        # ── Write merged result to staging ────────────────────────────
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
            ],
            error=None,
            metrics=ConversionMetrics(
                duration_ms=0.0,
                input_bytes=sum(os.path.getsize(p) for p in source_paths if os.path.isfile(p)),
                output_bytes=os.path.getsize(output_path) if os.path.isfile(output_path) else 0,
                extra={"collect_count": len(collect_paths), "mode": mode_str},
            ),
        )

    # ── Alignment ─────────────────────────────────────────────────────

    def _find_best_offset(
        self,
        base_ws: Any,
        collect_ws: Any,
        offset_range: int,
    ) -> tuple[int, int]:
        """Find best row/column offset via sliding window alignment."""
        max_overlap = 0
        best_offset = (0, 0)

        collect_cells = list(self._iter_cells_for_alignment(collect_ws))

        for row_off in range(-offset_range, offset_range + 1):
            for col_off in range(-offset_range, offset_range + 1):
                overlap = self._calc_overlap(base_ws, collect_cells, row_off, col_off)
                if overlap > max_overlap:
                    max_overlap = overlap
                    best_offset = (row_off, col_off)

        if max_overlap == 0:
            suggested = self._suggest_offset_by_value_matching(base_ws, collect_ws)
            if suggested is not None:
                return suggested

        return best_offset

    def _calc_overlap(
        self,
        base_ws: Any,
        collect_cells: list[tuple[int, int, object]],
        row_offset: int,
        col_offset: int,
    ) -> int:
        """Count matching cells at given offset."""
        overlap = 0
        for c_row, c_col, c_val in collect_cells:
            b_row = c_row + row_offset
            b_col = c_col + col_offset
            if b_row < 1 or b_col < 1:
                continue
            if b_row > base_ws.max_row or b_col > base_ws.max_column:
                continue
            b_val = self._normalize_for_alignment(base_ws.cell(b_row, b_col).value)
            if b_val is not None and c_val == b_val:
                overlap += 1
        return overlap

    def _iter_cells_for_alignment(self, ws: Any) -> list[tuple[int, int, object]]:
        """Iterate non-empty cells for alignment, up to MAX_CELLS_FOR_ALIGNMENT."""
        result: list[tuple[int, int, object]] = []
        cells = getattr(ws, "_cells", None)
        it = cells.values() if isinstance(cells, dict) else (cell for row in ws.iter_rows() for cell in row)
        for cell in it:
            v = self._normalize_for_alignment(cell.value)
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
    def _normalize_for_alignment(value: Any) -> Any:
        """Normalize a cell value for alignment comparison."""
        if value is None:
            return None
        if isinstance(value, bool):
            return ("bool", value)
        if isinstance(value, (int, float)):
            return ("num", float(value))
        if isinstance(value, (date, datetime)):
            return ("date", value.isoformat())
        if isinstance(value, str):
            s = value.strip()
            if not s:
                return None
            try:
                return ("num", float(s))
            except Exception:
                return ("str", s)
        s = str(value).strip()
        if not s:
            return None
        return ("str", s)

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
            worksheet.unmerge_cells(str(merged_range))
            for row in range(merged_range.min_row, merged_range.max_row + 1):
                for col in range(merged_range.min_col, merged_range.max_col + 1):
                    cell = worksheet.cell(row=row, column=col)
                    if type(cell).__name__ != "MergedCell":
                        cell.value = merged_value

    # ── Merge by row ──────────────────────────────────────────────────

    def _merge_by_row(self, base_ws: Any, collect_ws: Any, row_offset: int, col_offset: int) -> None:
        """Merge collection rows into base, inserting at most similar positions."""
        for collect_row_idx in range(1, collect_ws.max_row + 1):
            if self._is_row_empty(collect_ws, collect_row_idx):
                continue

            covered = False
            for base_row_idx in range(1, base_ws.max_row + 1):
                coverage = self._check_row_coverage(base_ws, base_row_idx, collect_ws, collect_row_idx, col_offset)
                if coverage == "skip_empty":
                    continue
                if coverage == "collect_covers_base":
                    self._replace_row(base_ws, base_row_idx, collect_ws, collect_row_idx, col_offset)
                    covered = True
                    break
                elif coverage == "base_covers_collect":
                    covered = True
                    break

            if not covered:
                self._insert_at_most_similar_row(base_ws, collect_ws, collect_row_idx, col_offset)

    @staticmethod
    def _is_row_empty(ws: Any, row_idx: int) -> bool:
        for col_idx in range(1, ws.max_column + 1):
            v = ws.cell(row_idx, col_idx).value
            if v is not None and str(v).strip() != "":
                return False
        return True

    def _check_row_coverage(
        self,
        base_ws: Any,
        base_row: int,
        collect_ws: Any,
        collect_row: int,
        col_offset: int,
    ) -> str:
        if base_row < 1 or base_row > base_ws.max_row:
            return "out_of_range"

        if self._is_row_empty(base_ws, base_row):
            return "skip_empty"

        base_has_unique = False
        collect_has_unique = False

        max_col = max(base_ws.max_column, collect_ws.max_column + col_offset)
        for col_idx in range(1, max_col + 1):
            base_val = self._normalize_for_alignment(
                base_ws.cell(base_row, col_idx).value if col_idx <= base_ws.max_column else None
            )
            base_empty = base_val is None

            cc = col_idx - col_offset
            if cc < 1 or cc > collect_ws.max_column:
                if not base_empty:
                    base_has_unique = True
                continue

            collect_val = self._normalize_for_alignment(collect_ws.cell(collect_row, cc).value)
            collect_empty = collect_val is None

            if base_empty and collect_empty:
                continue
            if base_empty and not collect_empty:
                collect_has_unique = True
            elif not base_empty and collect_empty:
                base_has_unique = True
            else:
                if base_val != collect_val:
                    return "conflict"

        if collect_has_unique and not base_has_unique:
            return "collect_covers_base"
        return "base_covers_collect"

    @staticmethod
    def _replace_row(base_ws: Any, base_row: int, collect_ws: Any, collect_row: int, col_offset: int) -> None:
        for col_idx in range(1, collect_ws.max_column + 1):
            base_col = col_idx + col_offset
            if base_col < 1:
                continue
            cell = base_ws.cell(base_row, base_col)
            if type(cell).__name__ != "MergedCell":
                cell.value = collect_ws.cell(collect_row, col_idx).value

    def _insert_at_most_similar_row(
        self,
        base_ws: Any,
        collect_ws: Any,
        collect_row: int,
        col_offset: int,
    ) -> None:
        max_sim = 0.0
        best_row = 1
        for base_row_idx in range(1, base_ws.max_row + 1):
            sim = self._row_similarity(base_ws, base_row_idx, collect_ws, collect_row, col_offset)
            if sim > max_sim:
                max_sim = sim
                best_row = base_row_idx

        insert_pos = best_row + 1
        base_ws.insert_rows(insert_pos)
        for col_idx in range(1, collect_ws.max_column + 1):
            base_col = col_idx + col_offset
            if base_col < 1:
                continue
            cell = base_ws.cell(insert_pos, base_col)
            if type(cell).__name__ != "MergedCell":
                cell.value = collect_ws.cell(collect_row, col_idx).value

    def _row_similarity(
        self,
        base_ws: Any,
        base_row: int,
        collect_ws: Any,
        collect_row: int,
        col_offset: int,
    ) -> float:
        total = 0
        match = 0
        max_col = max(base_ws.max_column, collect_ws.max_column + col_offset)
        for col_idx in range(1, max_col + 1):
            be = True
            bv = None
            if col_idx <= base_ws.max_column:
                bv = base_ws.cell(base_row, col_idx).value
                be = bv is None or str(bv).strip() == ""

            cc = col_idx - col_offset
            if cc < 1 or cc > collect_ws.max_column:
                if not be:
                    total += 1
                continue

            cv = collect_ws.cell(collect_row, cc).value
            ce = cv is None or str(cv).strip() == ""
            total += 1

            if (be and ce) or (not be and not ce and str(bv).strip() == str(cv).strip()):
                match += 1

        return match / total if total > 0 else 0.0

    # ── Merge by column ───────────────────────────────────────────────

    def _merge_by_column(
        self,
        base_ws: Any,
        collect_ws: Any,
        row_offset: int,
        col_offset: int,
    ) -> None:
        for collect_col_idx in range(1, collect_ws.max_column + 1):
            if self._is_col_empty(collect_ws, collect_col_idx):
                continue

            covered = False
            for base_col_idx in range(1, base_ws.max_column + 1):
                coverage = self._check_column_coverage(base_ws, base_col_idx, collect_ws, collect_col_idx, row_offset)
                if coverage == "skip_empty":
                    continue
                if coverage == "collect_covers_base":
                    self._replace_column(base_ws, base_col_idx, collect_ws, collect_col_idx, row_offset)
                    covered = True
                    break
                elif coverage == "base_covers_collect":
                    covered = True
                    break

            if not covered:
                self._insert_at_most_similar_column(base_ws, collect_ws, collect_col_idx, row_offset)

    @staticmethod
    def _is_col_empty(ws: Any, col_idx: int) -> bool:
        for row_idx in range(1, ws.max_row + 1):
            v = ws.cell(row_idx, col_idx).value
            if v is not None and str(v).strip() != "":
                return False
        return True

    def _check_column_coverage(
        self,
        base_ws: Any,
        base_col: int,
        collect_ws: Any,
        collect_col: int,
        row_offset: int,
    ) -> str:
        if base_col < 1 or base_col > base_ws.max_column:
            return "out_of_range"
        if self._is_col_empty(base_ws, base_col):
            return "skip_empty"

        base_has_unique = False
        collect_has_unique = False

        max_row = max(base_ws.max_row, collect_ws.max_row + row_offset)
        for row_idx in range(1, max_row + 1):
            base_val = self._normalize_for_alignment(
                base_ws.cell(row_idx, base_col).value if row_idx <= base_ws.max_row else None
            )
            base_empty = base_val is None

            cr = row_idx - row_offset
            if cr < 1 or cr > collect_ws.max_row:
                if not base_empty:
                    base_has_unique = True
                continue

            collect_val = self._normalize_for_alignment(collect_ws.cell(cr, collect_col).value)
            collect_empty = collect_val is None

            if base_empty and collect_empty:
                continue
            if base_empty and not collect_empty:
                collect_has_unique = True
            elif not base_empty and collect_empty:
                base_has_unique = True
            else:
                if base_val != collect_val:
                    return "conflict"

        if collect_has_unique and not base_has_unique:
            return "collect_covers_base"
        return "base_covers_collect"

    @staticmethod
    def _replace_column(base_ws: Any, base_col: int, collect_ws: Any, collect_col: int, row_offset: int) -> None:
        for row_idx in range(1, collect_ws.max_row + 1):
            base_row = row_idx + row_offset
            if base_row < 1:
                continue
            cell = base_ws.cell(base_row, base_col)
            if type(cell).__name__ != "MergedCell":
                cell.value = collect_ws.cell(row_idx, collect_col).value

    def _insert_at_most_similar_column(
        self,
        base_ws: Any,
        collect_ws: Any,
        collect_col: int,
        row_offset: int,
    ) -> None:
        max_sim = 0.0
        best_col = 1
        for base_col_idx in range(1, base_ws.max_column + 1):
            sim = self._col_similarity(base_ws, base_col_idx, collect_ws, collect_col, row_offset)
            if sim > max_sim:
                max_sim = sim
                best_col = base_col_idx

        insert_pos = best_col + 1
        base_ws.insert_cols(insert_pos)
        for row_idx in range(1, collect_ws.max_row + 1):
            base_row = row_idx + row_offset
            if base_row < 1:
                continue
            cell = base_ws.cell(base_row, insert_pos)
            if type(cell).__name__ != "MergedCell":
                cell.value = collect_ws.cell(row_idx, collect_col).value

    def _col_similarity(
        self,
        base_ws: Any,
        base_col: int,
        collect_ws: Any,
        collect_col: int,
        row_offset: int,
    ) -> float:
        total = 0
        match = 0
        max_row = max(base_ws.max_row, collect_ws.max_row + row_offset)
        for row_idx in range(1, max_row + 1):
            be = True
            bv = None
            if row_idx <= base_ws.max_row:
                bv = base_ws.cell(row_idx, base_col).value
                be = bv is None or str(bv).strip() == ""

            cr = row_idx - row_offset
            if cr < 1 or cr > collect_ws.max_row:
                if not be:
                    total += 1
                continue

            cv = collect_ws.cell(cr, collect_col).value
            ce = cv is None or str(cv).strip() == ""
            total += 1

            if (be and ce) or (not be and not ce and str(bv).strip() == str(cv).strip()):
                match += 1

        return match / total if total > 0 else 0.0

    # ── Merge by cell ─────────────────────────────────────────────────

    def _merge_by_cell(
        self,
        base_ws: Any,
        collect_ws: Any,
        row_offset: int,
        col_offset: int,
    ) -> None:
        for row in collect_ws.iter_rows():
            for cell in row:
                if cell.value is None or str(cell.value).strip() == "":
                    continue
                base_row = cell.row + row_offset
                base_col = cell.column + col_offset
                if base_row < 1 or base_col < 1:
                    continue
                base_cell = base_ws.cell(base_row, base_col)
                if type(base_cell).__name__ != "MergedCell":
                    base_cell.value = self._merge_cell_values(base_cell.value, cell.value)

    @staticmethod
    def _merge_cell_values(base_value: Any, collect_value: Any) -> Any:
        """Merge two cell values: numbers add, text concatenates."""
        be = base_value is None or str(base_value).strip() == ""
        ce = collect_value is None or str(collect_value).strip() == ""
        if be and ce:
            return None
        if be:
            return collect_value
        if ce:
            return base_value

        base_num = TableMergerConverter._try_number(base_value)
        col_num = TableMergerConverter._try_number(collect_value)
        if base_num is not None and col_num is not None:
            return base_num + col_num

        if str(base_value).strip() == str(collect_value).strip():
            return base_value
        return f"{base_value},{collect_value}"

    @staticmethod
    def _try_number(value: Any) -> float | int | None:
        if isinstance(value, (int, float)):
            return value
        try:
            num = float(str(value).strip())
            if num.is_integer():
                return int(num)
            return num
        except (ValueError, AttributeError):
            return None

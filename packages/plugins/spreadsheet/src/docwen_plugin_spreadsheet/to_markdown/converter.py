"""Real Spreadsheet → Markdown converter using openpyxl and pandas.

Implements ROUTE-SHEET-001 (spreadsheet → md) for XLSX and CSV inputs.

The converter:
- Only writes to staging via ``WorkspaceHandle``.
- Checks cancellation before expensive operations.
- Reports progress through ``ProgressSink``.
- Returns a ``ConversionResult`` with ``ArtifactManifest`` entries.
- Extracts embedded images from XLSX workbooks and registers them as
  staging artifacts (F-G6-012).
"""

from __future__ import annotations

import csv
import os
import re
import uuid
from collections import deque
from collections.abc import Callable
from typing import TYPE_CHECKING, Any

from docwen_core.detection import sniff_delimited_separator
from docwen_core.errors import CancellationRequested
from docwen_core.markdown_extensions import resolve_markdown_extensions
from docwen_core.markdown_utils import encode_markdown_table_cell
from docwen_core.text.ocr import report_ocr_outcome
from docwen_plugin_spreadsheet.delimited import decoded_samples
from docwen_plugin_spreadsheet.sparse_projection import (
    MAX_SPREADSHEET_PROJECTION_CELLS,
    ProjectionBudget,
    SpreadsheetProjectionLimitError,
    preflight_worksheet_allocations,
    worksheet_blocks,
)
from docwen_plugin_spreadsheet.worksheets import ordinary_worksheets

if TYPE_CHECKING:
    from docwen_core.protocols.execution_context import ConverterContext


def _report_ocr_best_effort(progress: Any, status: object, *, location: str) -> None:
    report_ocr_outcome(progress, status, location=location)


def _cell_has_value(value: Any) -> bool:
    """Check if a cell has meaningful content."""
    if value is None:
        return False
    if isinstance(value, str):
        return value.strip() != ""
    try:
        import pandas as pd

        return bool(pd.notna(value))
    except ImportError:
        return bool(value)


def _find_blocks(
    df: Any, *, cancel_check: Callable[[], None] | None = None, copy_budget: ProjectionBudget | None = None
) -> list[Any]:
    """Find all disconnected data blocks in a DataFrame using BFS.

    A "block" is a contiguous region of non-empty cells separated
    by empty rows/columns or DataFrame edges.
    """
    import pandas as pd

    if df.empty:
        return []

    rows, cols = df.shape
    if rows * cols > MAX_SPREADSHEET_PROJECTION_CELLS:
        raise SpreadsheetProjectionLimitError("DataFrame area exceeds the projection budget.")
    if copy_budget is None:
        copy_budget = ProjectionBudget(limit=MAX_SPREADSHEET_PROJECTION_CELLS)
    visited: set[tuple[int, int]] = set()
    blocks: list[pd.DataFrame] = []
    visits = 0

    for r in range(rows):
        for c in range(cols):
            if cancel_check is not None and (r * cols + c) % 1000 == 0:
                cancel_check()
            cell_value = df.iat[r, c]
            if (r, c) not in visited and _cell_has_value(cell_value):
                min_r, max_r, min_c, max_c = r, r, c, c

                q = deque([(r, c)])
                component = {(r, c)}
                visited.add((r, c))

                while q:
                    if cancel_check is not None and visits % 1000 == 0:
                        cancel_check()
                    visits += 1
                    curr_r, curr_c = q.popleft()
                    for dr, dc in [(-1, 0), (1, 0), (0, -1), (0, 1)]:
                        nr, nc = curr_r + dr, curr_c + dc
                        if not (0 <= nr < rows and 0 <= nc < cols):
                            continue
                        if (nr, nc) in visited:
                            continue
                        neighbor_value = df.iat[nr, nc]
                        if _cell_has_value(neighbor_value):
                            visited.add((nr, nc))
                            component.add((nr, nc))
                            q.append((nr, nc))
                            min_r = min(min_r, nr)
                            max_r = max(max_r, nr)
                            min_c = min(min_c, nc)
                            max_c = max(max_c, nc)

                copy_budget.reserve((max_r - min_r + 1) * (max_c - min_c + 1), location="Markdown block")
                block: pd.DataFrame = df.iloc[min_r : max_r + 1, min_c : max_c + 1].copy()  # type: ignore[assignment]
                # Another disconnected component can sit inside this bounding
                # rectangle (for example an L-shaped table around one cell).
                # Do not emit that cell in both blocks.
                for row_index in range(min_r, max_r + 1):
                    for column_index in range(min_c, max_c + 1):
                        if (
                            cancel_check is not None
                            and ((row_index - min_r) * (max_c - min_c + 1) + column_index - min_c) % 1000 == 0
                        ):
                            cancel_check()
                        if (row_index, column_index) not in component:
                            block.iat[row_index - min_r, column_index - min_c] = None
                block.dropna(axis="index", how="all", inplace=True)
                block.dropna(axis="columns", how="all", inplace=True)
                blocks.append(block.fillna(""))

    return blocks


def _process_cell_newlines(df: Any) -> Any:
    """Encode raw native cells once, before generated Markdown is added."""
    import pandas as pd

    result = df.copy()
    for column in result.columns:
        result[column] = result[column].apply(
            lambda cell: encode_markdown_table_cell(str(cell)) if pd.notna(cell) and cell != "" else cell
        )
    return result


_NUMERIC_TEXT_RE = re.compile(r"^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$")


def _requires_literal_numeric_text(value: Any) -> bool:
    """Return whether tabulate numeric parsing would risk changing authored text."""

    if not isinstance(value, str):
        return False
    text = value.strip()
    if not text or _NUMERIC_TEXT_RE.fullmatch(text) is None:
        return False

    unsigned = text.lstrip("+-")
    mantissa, exponent_sep, _exponent = unsigned.lower().partition("e")
    if exponent_sep or text.startswith("+"):
        return True
    integer, dot, fraction = mantissa.partition(".")
    if len(integer) > 1 and integer.startswith("0"):
        return True
    if not dot:
        return len(integer) > 15
    if fraction.endswith("0") and fraction:
        return True
    significant = (integer.lstrip("0") + fraction).lstrip("0")
    return len(significant) > 6


def _markdown_numparse_disabled_columns(data: Any) -> list[int]:
    disabled: list[int] = []
    for column_index in range(data.shape[1]):
        if any(_requires_literal_numeric_text(value) for value in data.iloc[:, column_index].tolist()):
            disabled.append(column_index)
    return disabled


def _dataframe_to_markdown(data: Any, headers: list[str]) -> str:
    if not hasattr(data, "to_markdown"):
        return ""
    disabled = _markdown_numparse_disabled_columns(data)
    if disabled:
        return data.to_markdown(index=False, headers=headers, disable_numparse=disabled)
    return data.to_markdown(index=False, headers=headers)


def _worksheet_merge_marker_positions(
    ws: Any, *, cancel_check: Callable[[], None] | None = None
) -> set[tuple[int, int]]:
    positions: set[tuple[int, int]] = set()
    visits = 0
    for merged_range in list(ws.merged_cells.ranges):
        anchor = (merged_range.min_row - 1, merged_range.min_col - 1)
        for row in range(merged_range.min_row - 1, merged_range.max_row):
            for column in range(merged_range.min_col - 1, merged_range.max_col):
                if cancel_check is not None and visits % 1000 == 0:
                    cancel_check()
                visits += 1
                if (row, column) != anchor:
                    positions.add((row, column))
    return positions


def _structural_header_rows_for_block(ws: Any, block: Any) -> int:
    if block.empty:
        return 0
    first_row = int(block.index[0])
    block_columns = [int(column) for column in block.columns]
    if not block_columns:
        return 0
    min_column = min(block_columns)
    max_column = max(block_columns)
    for merged_range in list(ws.merged_cells.ranges):
        start_row = merged_range.min_row - 1
        end_row = merged_range.max_row - 1
        start_column = merged_range.min_col - 1
        end_column = merged_range.max_col - 1
        crosses_boundary = start_row <= first_row < end_row
        overlaps_columns = start_column <= max_column and min_column <= end_column
        if crosses_boundary and overlaps_columns:
            return 0
    return 1


def _escape_literal_structural_markers(
    block: Any,
    structural_positions: set[tuple[int, int]],
) -> Any:
    escaped = block.copy()
    for row_offset, row_label in enumerate(escaped.index):
        for column_offset, column_label in enumerate(escaped.columns):
            value = escaped.iat[row_offset, column_offset]
            text = "" if value is None else str(value)
            stripped = text.strip()
            if stripped in {"<", "^"} and (int(row_label), int(column_label)) not in structural_positions:
                escaped.iat[row_offset, column_offset] = f"\\{stripped}"
    return escaped


def _render_structural_table_block(block: Any, *, header_rows: int) -> str:
    rows = [
        ["" if value is None else str(value) for value in block.iloc[row_index].tolist()]
        for row_index in range(block.shape[0])
    ]
    if not rows:
        return ""
    delimiter = ["---"] * block.shape[1]

    def render_row(values: list[str]) -> str:
        return "| " + " | ".join(values) + " |"

    output: list[str] = []
    if header_rows == 0:
        output.append(render_row(delimiter))
        output.extend(render_row(row) for row in rows)
    else:
        output.append(render_row(rows[0]))
        output.append(render_row(delimiter))
        output.extend(render_row(row) for row in rows[1:])
    return "\n".join(output)


def _worksheet_to_dataframe(
    ws: Any,
    table_merge_strategy: str = "fill",
) -> Any:
    """Bounded dense helper; production export uses sparse worksheet blocks.

    Only stored values are read. Original coordinate labels are retained, so a
    distant scalar does not allocate or mutate the origin rectangle.
    """
    import pandas as pd

    from docwen_core.export_semantics import TableMergeRegion, build_table_semantic_grid, render_table_semantic_grid

    values = {
        (cell.row - 1, cell.column - 1): str(cell.value)
        for cell in ws._cells.values()
        if cell.value is not None and (not isinstance(cell.value, str) or cell.value.strip())
    }
    merges = [
        TableMergeRegion(native.min_row - 1, native.min_col - 1, native.max_row - 1, native.max_col - 1)
        for native in ws.merged_cells.ranges
    ]
    positions = set(values)
    for region in merges:
        positions.update(((region.start_row, region.start_col), (region.end_row, region.end_col)))
    if not positions:
        return pd.DataFrame()
    top, left = min(row for row, _ in positions), min(column for _, column in positions)
    bottom, right = max(row for row, _ in positions), max(column for _, column in positions)
    ProjectionBudget(limit=MAX_SPREADSHEET_PROJECTION_CELLS).reserve(
        (bottom - top + 1) * (right - left + 1), location=ws.title
    )
    grid = build_table_semantic_grid(
        row_count=bottom - top + 1,
        col_count=right - left + 1,
        cell_text_by_position={(row - top, column - left): text for (row, column), text in values.items()},
        merge_regions=[
            TableMergeRegion(
                region.start_row - top, region.start_col - left, region.end_row - top, region.end_col - left
            )
            for region in merges
        ],
        max_cells=MAX_SPREADSHEET_PROJECTION_CELLS,
    )
    return pd.DataFrame(
        render_table_semantic_grid(grid, strategy=table_merge_strategy),
        index=range(top, bottom + 1),
        columns=range(left, right + 1),
    )


def _read_csv_flexible(file_path: str, source_format: str, *, cancel_check: Callable[[], None] | None = None) -> Any:
    """Read literal records using the same delimiter fact as admission/hub."""
    import pandas as pd

    candidates = decoded_samples(file_path) or [("utf-8-sig", "")]
    for index, (encoding, sample) in enumerate(candidates):
        delimiter = sniff_delimited_separator(sample) or ("\t" if source_format == "tsv" else ",")
        rows: list[list[str]] = []
        cell_count = 0
        widest = 0
        try:
            with open(file_path, encoding=encoding, errors="strict", newline="") as stream:
                for row_index, row in enumerate(csv.reader(stream, delimiter=delimiter, strict=True)):
                    if cancel_check is not None and row_index % 1000 == 0:
                        cancel_check()
                    # Preserve the existing TSV blank-row contract; CSV skips
                    # physically blank records, not empty fields in a record.
                    if not row and source_format != "tsv" and delimiter != "\t":
                        continue
                    cell_count += len(row)
                    widest = max(widest, len(row))
                    if (
                        cell_count > MAX_SPREADSHEET_PROJECTION_CELLS
                        or (len(rows) + 1) * widest > MAX_SPREADSHEET_PROJECTION_CELLS
                    ):
                        raise SpreadsheetProjectionLimitError("Delimited Markdown projection exceeds its cell budget.")
                    for _column in range(0, len(row), 1000):
                        if cancel_check is not None:
                            cancel_check()
                    rows.append(row)
        except UnicodeError:
            if index == len(candidates) - 1:
                raise
        else:
            if cancel_check is not None:
                cancel_check()
            return pd.DataFrame(rows, dtype=object).fillna("")
    raise AssertionError("Delimited decoding produced no records")


def _extract_sheet_images(
    ws: Any,
    sheet_name: str,
    context: ConverterContext,
    *,
    image_counter_start: int = 1,
    register_artifacts: bool = True,
) -> tuple[list[dict[str, Any]], int]:
    """Extract embedded images from an openpyxl worksheet to staging.

    Each extracted image is written to staging via
    ``context.workspace.create_artifact_path`` and registered as an
    ``ArtifactManifest`` with ``kind="image"``.

    Returns the extracted image records and a count of images that could not
    be extracted.  Each record contains:
    - ``markdown_link``: formatted Markdown image reference
    - ``artifact_id``: the registered artifact ID
    - ``suggested_name``: suggested output filename
    - ``row`` / ``col``: 1-based position within the sheet (or ``None``)
    """
    images_info: list[dict[str, Any]] = []
    extraction_loss_count = 0

    if not hasattr(ws, "_images") or not ws._images:
        return images_info, extraction_loss_count

    from docwen_core.markdown_utils import format_sanitized_image_link

    image_counter = image_counter_start

    for img in ws._images:
        context.cancellation.check()
        try:
            image_data = img._data()
            context.cancellation.check()

            # Detect image format
            ext = "png"
            if hasattr(img, "format"):
                ext = img.format.lower()

            # Save to staging
            staging_path = context.workspace.create_artifact_path(
                "image",
                f".{ext}",
            )
            with open(staging_path, "wb") as f:
                f.write(image_data)

            # Determine row/col position (1-based)
            row: int | None = None
            col: int | None = None
            if hasattr(img, "anchor") and img.anchor is not None and hasattr(img.anchor, "_from"):
                row = img.anchor._from.row + 1
                col = img.anchor._from.col + 1

            # Build a descriptive suggested filename
            suggested_name = f"{sheet_name}_image{image_counter}.{ext}"

            # Generate Markdown reference using shared core utility
            markdown_link = format_sanitized_image_link(suggested_name)

            if register_artifacts:
                from docwen_core.models.artifact import ArtifactManifest

                artifact = ArtifactManifest(
                    artifact_id=str(uuid.uuid4()),
                    kind="image",
                    staging_path=staging_path,
                    suggested_name=suggested_name,
                    media_type=f"image/{ext}",
                    metadata={
                        "sheet_name": sheet_name,
                        "row": row,
                        "col": col,
                    },
                )
                context.workspace.add_artifact(artifact)
                artifact_id = artifact.artifact_id
            else:
                artifact_id = None

            images_info.append(
                {
                    "markdown_link": markdown_link,
                    "artifact_id": artifact_id,
                    "suggested_name": suggested_name,
                    "row": row,
                    "col": col,
                    "staging_path": staging_path,
                }
            )

            context.logger.debug(
                f"Extracted image {suggested_name} from sheet {sheet_name!r}",
            )

            image_counter += 1

        except CancellationRequested:
            raise
        except Exception as exc:
            extraction_loss_count += 1
            context.logger.warning(
                f"Failed to extract image from sheet {sheet_name!r}: {exc}",
            )
            continue

    return images_info, extraction_loss_count


class SpreadsheetToMarkdownConverter:
    """Convert XLSX or CSV files to Markdown.

    Uses ``openpyxl`` for XLSX parsing and ``pandas`` for data handling.
    Produces a structured Markdown document with YAML frontmatter.
    """

    def convert(self, context: ConverterContext) -> Any:
        """Run the spreadsheet → Markdown conversion.

        Args:
            context: The plugin execution context.

        Returns:
            ``ConversionResult`` with staging artifacts.
        """
        from docwen_core.models.result import (
            ConversionDiagnostic,
            ConversionErrorInfo,
            ConversionMetrics,
            ConversionResult,
        )

        task_id = context.request.request_id
        input_path = context.workspace.input_path
        if not context.request.input_refs:
            raise ValueError("Spreadsheet conversion requires an admitted input reference.")
        source_format = str(context.request.input_refs[0].format or "").strip().lower().lstrip(".")

        # 1. Check cancellation
        context.cancellation.check()

        # 2. Read options
        options = context.request.options
        keep_images = options.get("to_md_keep_images", True)
        merge_strategy = options.get("table_merge_strategy", "fill")
        extensions = resolve_markdown_extensions(options, context.config, direction="output")
        if extensions.structural_tables:
            merge_strategy = "marker"
        elif merge_strategy == "marker":
            merge_strategy = "fill"

        # All previously reserved options have been implemented.
        # No reserved options remain — the optimize_for concept has been
        # deleted; action_name routing subsumes it entirely.

        # 3. Convert
        context.progress.report_progress(0.0, "Starting spreadsheet → Markdown conversion")
        context.logger.info(f"Sheet→MD: reading {input_path}")

        try:
            markdown_content, stats = self._convert_to_markdown(
                input_path=input_path,
                source_format=source_format,
                merge_strategy=merge_strategy,
                keep_images=keep_images,
                context=context,
                options=options,
            )
        except CancellationRequested:
            raise
        except Exception as exc:
            context.logger.error(f"Sheet→MD conversion failed: {exc}")
            return ConversionResult(
                task_id=task_id,
                success=False,
                error=ConversionErrorInfo(
                    error_type="conversion_failed",
                    message=str(exc),
                    diagnostic_code=getattr(exc, "diagnostic_code", "SHEET2MD-PARSE-ERROR"),
                ),
                diagnostics=[
                    ConversionDiagnostic(
                        level="error",
                        message=f"Failed to parse spreadsheet: {exc}",
                        code=getattr(exc, "diagnostic_code", "SHEET2MD-PARSE-ERROR"),
                    ),
                ],
            )

        # 4. Write to staging
        context.cancellation.check()
        context.progress.report_progress(80.0, "Writing Markdown to staging...")

        output_path = context.workspace.create_artifact_path("primary", ".md")
        try:
            with open(output_path, "w", encoding="utf-8") as f:
                f.write(markdown_content)
        except OSError as exc:
            context.logger.error(f"Sheet→MD write failed: {exc}")
            return ConversionResult(
                task_id=task_id,
                success=False,
                error=ConversionErrorInfo(
                    error_type="conversion_failed",
                    message=f"Failed to write output file: {exc}",
                    diagnostic_code="SHEET2MD-WRITE-ERROR",
                ),
                diagnostics=[
                    ConversionDiagnostic(
                        level="error",
                        message=f"File write error at {output_path}: {exc}",
                        code="SHEET2MD-WRITE-ERROR",
                    ),
                ],
            )

        # 5. Build artifact manifest
        input_basename = os.path.basename(input_path)
        suggested_name = input_basename.rsplit(".", 1)[0] + ".md"

        from docwen_core.models.artifact import ArtifactManifest

        artifact = ArtifactManifest(
            artifact_id=str(uuid.uuid4()),
            kind="primary",
            staging_path=output_path,
            suggested_name=suggested_name,
            media_type="text/markdown",
            metadata={
                "sheet_count": stats.get("sheets", 0),
                "row_count": stats.get("rows", 0),
                "col_count": stats.get("cols", 0),
                "block_count": stats.get("blocks", 0),
                "image_count": stats.get("images", 0),
                "merge_strategy": merge_strategy,
                **(
                    {"formula_cache_unavailable_count": stats["formula_cache_unavailable_count"]}
                    if stats.get("formula_cache_unavailable_count")
                    else {}
                ),
            },
            is_primary=True,
        )
        image_extraction_loss_count = stats.get("image_extraction_loss_count", 0)
        if image_extraction_loss_count:
            artifact.metadata["image_extraction_loss_count"] = image_extraction_loss_count
        context.workspace.add_artifact(artifact)
        registered_artifacts = [
            registered
            for registered in getattr(context.workspace, "registered_artifacts", [])
            if registered.artifact_id != artifact.artifact_id
        ]

        # 6. Report completion
        context.progress.report_artifact_ready(artifact.artifact_id, suggested_name)
        context.progress.report_progress(100.0, "Conversion complete")
        context.logger.info(
            f"Sheet→MD complete: {stats['sheets']} sheets, "
            f"{stats['rows']} rows, {stats['blocks']} blocks, "
            f"{stats.get('images', 0)} images"
        )

        diagnostics = [
            ConversionDiagnostic(
                level="info",
                message=(
                    f"Converted spreadsheet to Markdown: "
                    f"{stats['sheets']} sheets, {stats['rows']} rows, "
                    f"{stats['blocks']} blocks, "
                    f"{stats.get('images', 0)} images"
                ),
                code="SHEET2MD-OK",
            ),
        ]
        if image_extraction_loss_count:
            diagnostics.append(
                ConversionDiagnostic(
                    level="warning",
                    message=(
                        f"{image_extraction_loss_count} embedded image(s) could not be extracted. "
                        "The spreadsheet tables were converted, but those images are missing."
                    ),
                    code="SHEET2MD-IMAGE-EXTRACTION-LOSS",
                    location="workbook embedded images",
                )
            )
        formula_cache_warning = stats.get("formula_cache_warning_message")
        if isinstance(formula_cache_warning, str) and formula_cache_warning:
            diagnostics.append(
                ConversionDiagnostic(
                    level="warning",
                    message=formula_cache_warning,
                    code="SHEET2MD-FORMULA-CACHE-UNAVAILABLE",
                    location="workbook formulas",
                )
            )

        return ConversionResult(
            task_id=task_id,
            success=True,
            artifacts=[artifact, *registered_artifacts],
            diagnostics=diagnostics,
            error=None,
            metrics=ConversionMetrics(
                duration_ms=0.0,
                input_bytes=os.path.getsize(input_path) if os.path.isfile(input_path) else 0,
                output_bytes=len(markdown_content.encode("utf-8")),
                extra=stats,
            ),
        )

    # ── Internal conversion ────────────────────────────────────────────

    def _convert_to_markdown(
        self,
        input_path: str,
        source_format: str,
        merge_strategy: str,
        keep_images: bool,
        context: ConverterContext,
        options: dict | None = None,
    ) -> tuple[str, dict[str, int]]:
        """Convert spreadsheet file to Markdown text.

        Returns (markdown_text, stats_dict).
        """
        file_stem = context.request.source_stem
        stats: dict[str, Any] = {"sheets": 0, "rows": 0, "cols": 0, "blocks": 0}

        # YAML frontmatter — routed through shared core utility (F-I2b-001)
        from docwen_core.yaml_tools import generate_basic_yaml_frontmatter

        md_content = generate_basic_yaml_frontmatter(
            file_stem,
            yaml_key_labels=(options or {}).get("yaml_key_labels"),
        )

        if source_format in ("csv", "tsv"):
            return self._convert_csv(input_path, source_format, file_stem, md_content, stats, context)
        if source_format in ("xlsx", "xlsm", "xltx", "xltm"):
            return self._convert_xlsx(
                input_path,
                md_content,
                merge_strategy,
                keep_images,
                stats,
                context,
                options=options,
            )
        raise ValueError(f"Unsupported admitted spreadsheet format: {source_format or 'unknown'}")

    def _convert_csv(
        self,
        input_path: str,
        source_format: str,
        file_stem: str,
        md_content: str,
        stats: dict[str, Any],
        context: ConverterContext,
    ) -> tuple[str, dict[str, Any]]:
        """Convert a CSV file to Markdown."""
        context.progress.report_progress(20.0, "Reading CSV...")

        md_content += f"# {file_stem}\n\n"

        df = _read_csv_flexible(input_path, source_format, cancel_check=context.cancellation.check)
        stats["sheets"] = 1
        stats["rows"] = df.shape[0]
        stats["cols"] = df.shape[1]

        blocks = _find_blocks(df, cancel_check=context.cancellation.check)
        stats["blocks"] = len(blocks)
        stats["images"] = 0

        for _idx, block in enumerate(blocks):
            context.cancellation.check()

            block = _process_cell_newlines(block)

            if block.shape[0] > 0:
                headers = [
                    str(h) if h is not None and str(h).strip() != "" else f"Col{j}"
                    for j, h in enumerate(block.iloc[0].tolist())
                ]
                data = block.iloc[1:]
                block_md = _dataframe_to_markdown(data, headers)
            else:
                block_md = ""

            md_content += (block_md or "") + "\n\n"

        return md_content.strip() + "\n", stats

    def _convert_xlsx(
        self,
        input_path: str,
        md_content: str,
        merge_strategy: str,
        keep_images: bool,
        stats: dict[str, Any],
        context: ConverterContext,
        options: dict | None = None,
    ) -> tuple[str, dict[str, Any]]:
        """Convert an XLSX file to Markdown.

        When *keep_images* is ``True``, embedded images are extracted from
        each worksheet, written to staging, and referenced in the Markdown
        output (F-G6-012).
        """
        if options is None:
            options = context.request.options
        from docwen_plugin_spreadsheet.csv_xlsx.converter import (
            _find_unavailable_formula_caches,
            _formula_cache_warning_message,
            _load_xlsx_views,
        )

        context.progress.report_progress(20.0, "Opening workbook...")

        preflight_worksheet_allocations(
            input_path, cancel_check=context.cancellation.check, limit=MAX_SPREADSHEET_PROJECTION_CELLS
        )
        wb, formula_wb = _load_xlsx_views(input_path)
        try:
            worksheets = ordinary_worksheets(wb)
            budget = ProjectionBudget(limit=MAX_SPREADSHEET_PROJECTION_CELLS)
            copy_budget = ProjectionBudget(limit=MAX_SPREADSHEET_PROJECTION_CELLS)
            unavailable_formula_count, unavailable_formula_locations = _find_unavailable_formula_caches(
                wb,
                formula_wb,
                input_path=input_path,
                sheet_names=[worksheet.title for worksheet in worksheets],
                cancel_check=context.cancellation.check,
            )
            if unavailable_formula_count:
                stats["formula_cache_unavailable_count"] = unavailable_formula_count
                stats["formula_cache_warning_message"] = _formula_cache_warning_message(
                    unavailable_formula_count,
                    unavailable_formula_locations,
                )
            stats["sheets"] = len(worksheets)

            total_images_extracted = 0
            total_image_extraction_losses = 0

            total_sheets = len(worksheets)
            for sheet_idx, ws in enumerate(worksheets):
                sheet_name = ws.title
                context.cancellation.check()
                progress = 20.0 + 60.0 * (sheet_idx / max(total_sheets, 1))
                context.progress.report_progress(progress, f"Processing sheet: {sheet_name}")

                md_content += f"# {sheet_name}\n\n"

                # ── Image extraction with position injection (H6) and OCR support (M8) ──
                enable_ocr = options.get("to_md_enable_ocr", False)
                extract_images = keep_images or enable_ocr
                image_positions: dict[tuple[int, int], list[str]] = {}
                sheet_images: list[dict[str, Any]] = []

                if extract_images:
                    from docwen_core.export_semantics import (
                        get_markdown_export_modes,
                        resolve_markdown_request_policy,
                    )

                    request_policy = resolve_markdown_request_policy(context)
                    export_semantics = request_policy.export

                    export_modes = get_markdown_export_modes(
                        "xlsx",
                        extraction_mode=options.get("image_mode"),
                        ocr_placement_mode=options.get("ocr_placement"),
                        semantics=export_semantics,
                    )
                    image_mode = export_modes["image_extraction_mode"]
                    if image_mode not in {"file", "base64", "embed", "omit"}:
                        image_mode = "file"
                    ocr_placement = export_modes["ocr_placement_mode"]
                    if ocr_placement not in {"image_md", "main_md"}:
                        ocr_placement = "main_md"
                    ocr_language = str(options.get("ocr_language") or "auto")
                    current_locale = str(options.get("locale") or "zh_CN")

                    if total_images_extracted + len(ws._images) > MAX_SPREADSHEET_PROJECTION_CELLS:
                        raise SpreadsheetProjectionLimitError("Image occurrences exceed the projection budget.")
                    sheet_images, sheet_image_extraction_losses = _extract_sheet_images(
                        ws,
                        sheet_name,
                        context,
                        image_counter_start=total_images_extracted + 1,
                        register_artifacts=keep_images and image_mode in {"file", "embed"},
                    )
                    total_image_extraction_losses += sheet_image_extraction_losses

                    if sheet_images:
                        from docwen_core.text.image_markdown import generate_image_markdown

                        image_link_style = options.get("image_link_style", export_semantics.image_link_style)
                        md_link_style: str = export_semantics.md_file_link_style

                        # Sidecar names must be unique across the entire workbook,
                        # not just within one sheet.  Relocation maps old references
                        # by name; duplicate names would make later links point at
                        # the first sheet's sidecar.
                        img_seq = total_images_extracted
                        main_stem = os.path.splitext(os.path.basename(str(input_path)))[0]

                        for img_info in sheet_images:
                            # Build image markdown link
                            img_path = (
                                img_info["staging_path"] if image_mode == "base64" else img_info["suggested_name"]
                            )
                            img_md = ""
                            if keep_images:
                                img_md = generate_image_markdown(
                                    image_path=img_path,
                                    image_mode=image_mode,
                                    image_link_style=image_link_style,
                                    alt_text=img_info.get("suggested_name", ""),
                                    export_semantics=export_semantics,
                                )

                            # M8: OCR always runs when enabled, regardless of keep_images
                            if enable_ocr:
                                from docwen_core.detection import detect_content_format
                                from docwen_core.text.ocr import run_ocr_outcome
                                from docwen_core.text.table_recognition import (
                                    enrich_ocr_table_structure,
                                )

                                outcome = run_ocr_outcome(
                                    img_info["staging_path"],
                                    source_format=detect_content_format(img_info["staging_path"]).format,
                                    ocr_language=ocr_language,
                                    current_locale=current_locale,
                                )
                                outcome, table_outcome = enrich_ocr_table_structure(
                                    img_info["staging_path"], outcome, context=context
                                )
                                if table_outcome.fallback_required:
                                    context.progress.report_diagnostic(
                                        "warning",
                                        "Table structure recognition failed; plain OCR text was retained.",
                                        code=table_outcome.diagnostic_code,
                                        location=(
                                            f"{sheet_name}:{img_info.get('suggested_name', img_info['staging_path'])}"
                                        ),
                                    )
                                _report_ocr_best_effort(
                                    context.progress,
                                    outcome.status,
                                    location=(
                                        f"{sheet_name}:{img_info.get('suggested_name', img_info['staging_path'])}"
                                    ),
                                )
                                ocr_text = outcome.recognized_text
                                ocr_markdown = outcome.structured_markdown
                                if ocr_text or ocr_markdown:
                                    if ocr_placement == "image_md":
                                        img_seq += 1
                                        sidecar_stem = f"{main_stem}__img_{img_seq:03d}_ocr"
                                        from docwen_core.text.image_markdown import build_image_ocr_sidecar

                                        sidecar_text, repl_link = build_image_ocr_sidecar(
                                            sidecar_stem=sidecar_stem,
                                            source_format="xlsx",
                                            image_markdown=img_md,
                                            ocr_text=ocr_text,
                                            md_link_style=md_link_style,
                                            ocr_markdown=ocr_markdown or None,
                                            ocr_blockquote_title=request_policy.ocr_blockquote_title,
                                        )
                                        from pathlib import Path as _Path

                                        sidecar_path = context.workspace.create_artifact_path("auxiliary", ".md")
                                        _Path(sidecar_path).write_text(sidecar_text, encoding="utf-8")
                                        import uuid

                                        from docwen_core.markdown_utils import sanitize_filename
                                        from docwen_core.models.artifact import ArtifactManifest

                                        ocr_md_name = sanitize_filename(f"{sidecar_stem}.md")
                                        sidecar_artifact = ArtifactManifest(
                                            artifact_id=str(uuid.uuid4()),
                                            kind="auxiliary",
                                            staging_path=sidecar_path,
                                            suggested_name=ocr_md_name,
                                            media_type="text/markdown",
                                            metadata={"source_format": "xlsx", "ocr": True},
                                            is_primary=False,
                                        )
                                        context.workspace.add_artifact(sidecar_artifact)
                                        img_md = repl_link
                                    else:
                                        if ocr_markdown:
                                            structured = f"\n\n{ocr_markdown}\n"
                                            img_md += structured if img_md else structured.lstrip()
                                        else:
                                            ocr_block = f"\n> {encode_markdown_table_cell(request_policy.ocr_blockquote_title)}\n> "
                                            ocr_block += (
                                                "\n> ".join(
                                                    encode_markdown_table_cell(line) for line in ocr_text.splitlines()
                                                )
                                                + "\n"
                                            )
                                            img_md += "\n" + ocr_block if img_md else ocr_block

                            # Record position for H6 cell injection (0-indexed)
                            r = img_info.get("row")
                            c = img_info.get("col")
                            if r is not None and c is not None:
                                key = (r - 1, c - 1)
                                image_positions.setdefault(key, []).append(img_md)

                            # Clean up temp file if not keeping images
                            if not keep_images:
                                import contextlib

                                with contextlib.suppress(OSError):
                                    os.unlink(img_info["staging_path"])

                total_images_extracted += len(sheet_images)

                structural_tables = resolve_markdown_extensions(
                    options,
                    context.config,
                    direction="output",
                ).structural_tables
                if ws.merged_cells.ranges and not structural_tables:
                    context.progress.report_diagnostic(
                        "warning",
                        "Merged cells were flattened to a standard Markdown table.",
                        code="docwen.conversion.markdown_extension.structural_tables.flattened",
                        location=ws.title,
                    )
                blocks = worksheet_blocks(
                    ws,
                    strategy=merge_strategy,
                    generated=image_positions,
                    budget=budget,
                    copy_budget=copy_budget,
                    cancel_check=context.cancellation.check,
                    split_rendered=_find_blocks,
                )
                # Retain metadata's source-coordinate extent without using it
                # as a dense allocation shape. Image anchors also contributed
                # to this extent before sparse projection was introduced.
                row_extent, column_extent = ws.max_row, ws.max_column
                for row, column in image_positions:
                    row_extent = max(row_extent, row + 1)
                    column_extent = max(column_extent, column + 1)
                stats["rows"] += row_extent
                stats["cols"] = max(stats["cols"], column_extent)

                stats["blocks"] += len(blocks)

                if not blocks:
                    md_content += " (this sheet is empty)\n\n"
                else:
                    structural_positions = _worksheet_merge_marker_positions(
                        ws, cancel_check=context.cancellation.check
                    )
                    for block in blocks:
                        if block.shape[0] > 0 and structural_tables:
                            block = _escape_literal_structural_markers(block, structural_positions)
                            block_md = _render_structural_table_block(
                                block,
                                header_rows=_structural_header_rows_for_block(ws, block),
                            )
                        elif block.shape[0] > 0:
                            headers = [
                                str(h) if h is not None and str(h).strip() != "" else f"Col{j}"
                                for j, h in enumerate(block.iloc[0].tolist())
                            ]
                            data = block.iloc[1:]
                            block_md = _dataframe_to_markdown(data, headers)
                        else:
                            block_md = ""

                        md_content += (block_md or "") + "\n\n"

            stats["images"] = total_images_extracted
            if total_image_extraction_losses:
                stats["image_extraction_loss_count"] = total_image_extraction_losses
            return md_content.strip() + "\n", stats
        finally:
            wb.close()
            formula_wb.close()

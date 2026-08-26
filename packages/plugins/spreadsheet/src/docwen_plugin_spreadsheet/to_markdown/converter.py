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
import uuid
from collections import deque
from pathlib import Path
from typing import TYPE_CHECKING, Any

from docwen_core.text.ocr import format_ocr_best_effort_warning

if TYPE_CHECKING:
    from docwen_core.protocols.execution_context import ConverterContext


def _report_ocr_best_effort(progress: Any, status: object, *, location: str) -> None:
    """Report one safe, request-visible warning for a fallible OCR outcome."""
    message = format_ocr_best_effort_warning(status)
    if message is None:
        return
    progress.report_diagnostic(
        "warning",
        message,
        code="OCR-BEST-EFFORT",
        location=location,
    )


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


def _find_blocks(df: Any) -> list[Any]:
    """Find all disconnected data blocks in a DataFrame using BFS.

    A "block" is a contiguous region of non-empty cells separated
    by empty rows/columns or DataFrame edges.
    """
    import pandas as pd

    if df.empty:
        return []

    rows, cols = df.shape
    visited: set[tuple[int, int]] = set()
    blocks: list[pd.DataFrame] = []

    for r in range(rows):
        for c in range(cols):
            cell_value = df.iat[r, c]
            if (r, c) not in visited and _cell_has_value(cell_value):
                min_r, max_r, min_c, max_c = r, r, c, c

                q = deque([(r, c)])
                visited.add((r, c))

                while q:
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
                            q.append((nr, nc))
                            min_r = min(min_r, nr)
                            max_r = max(max_r, nr)
                            min_c = min(min_c, nc)
                            max_c = max(max_c, nc)

                block: pd.DataFrame = df.iloc[min_r : max_r + 1, min_c : max_c + 1].copy()  # type: ignore[assignment]
                block.dropna(axis="index", how="all", inplace=True)
                block.dropna(axis="columns", how="all", inplace=True)
                blocks.append(block)

    return blocks


def _process_cell_newlines(df: Any) -> Any:
    """Replace newlines in cells with <br> and escape pipe characters."""
    import pandas as pd

    df_copy = df.copy()

    for col in df_copy.columns:
        df_copy[col] = df_copy[col].apply(
            lambda x: (
                str(x).replace("\r\n", "<br>").replace("\n", "<br>").replace("\r", "<br>").replace("|", "\\|")
                if pd.notna(x) and x != ""
                else x
            )
        )

    return df_copy


def _worksheet_to_dataframe(
    ws: Any,
    table_merge_strategy: str = "fill",
) -> Any:
    """Convert an openpyxl worksheet to a pandas DataFrame.

    Handles merged cells according to the specified strategy, delegating
    grid construction and rendering to the shared
    ``docwen_core.export_semantics`` module.

    F-H3b-027, F-H3b-029
    """
    import pandas as pd

    from docwen_core.export_semantics import (
        TableMergeRegion,
        build_table_semantic_grid,
        normalize_table_merge_export_strategy,
        render_table_semantic_grid,
    )

    strategy = normalize_table_merge_export_strategy(table_merge_strategy)

    cell_text_by_position: dict[tuple[int, int], str] = {}
    max_row = 0
    max_col = 0

    # Determine dimensions from populated cells.
    if getattr(ws, "_cells", None):
        for cell in ws._cells.values():
            if cell.value is None:
                continue
            if isinstance(cell.value, str) and cell.value.strip() == "":
                continue
            max_row = max(max_row, cell.row)
            max_col = max(max_col, cell.column)

    # Collect merged cell regions and ensure anchor values are recorded.
    merge_regions: list[TableMergeRegion] = []
    for merged_range in list(ws.merged_cells.ranges):
        range_min_row = merged_range.min_row
        range_min_col = merged_range.min_col
        range_max_row = merged_range.max_row
        range_max_col = merged_range.max_col
        max_row = max(max_row, range_max_row)
        max_col = max(max_col, range_max_col)
        merge_regions.append(
            TableMergeRegion(
                start_row=range_min_row - 1,
                start_col=range_min_col - 1,
                end_row=range_max_row - 1,
                end_col=range_max_col - 1,
            )
        )
        anchor_value = ws.cell(row=range_min_row, column=range_min_col).value
        cell_text_by_position.setdefault((range_min_row - 1, range_min_col - 1), str(anchor_value or ""))

    if max_row <= 0 or max_col <= 0:
        return pd.DataFrame()

    # Populate cell_text_by_position for every (row, col) so that
    # build_table_semantic_grid has a complete sparse-text map.
    for row_idx in range(max_row):
        for col_idx in range(max_col):
            pos = (row_idx, col_idx)
            if pos in cell_text_by_position:
                continue
            value = ws.cell(row=row_idx + 1, column=col_idx + 1).value
            if value is None:
                continue
            if isinstance(value, str) and value.strip() == "":
                continue
            cell_text_by_position[pos] = str(value)

    # Build and render the semantic grid.
    grid = build_table_semantic_grid(
        row_count=max_row,
        col_count=max_col,
        cell_text_by_position=cell_text_by_position,
        merge_regions=merge_regions,
    )
    rendered = render_table_semantic_grid(grid, strategy=strategy)

    return pd.DataFrame(rendered)


def _read_csv_flexible(file_path: str, source_format: str) -> Any:
    """Read a CSV file with flexible encoding and delimiter detection."""
    import pandas as pd

    candidates = ("utf-8-sig", "utf-8", "gbk", "utf-16")
    delimiters = ",;\t|"
    is_tsv = source_format == "tsv"

    def should_skip_blank_lines(sep: str | None = None) -> bool:
        return not (is_tsv or sep == "\t")

    try:
        with Path(file_path).open("rb") as f:
            sample_bytes = f.read(65536)
    except Exception:
        return pd.read_csv(
            file_path,
            header=None,
            keep_default_na=False,
            skip_blank_lines=should_skip_blank_lines(),
        )

    for encoding in candidates:
        try:
            sample_text = sample_bytes.decode(encoding)
        except Exception:
            continue

        sep = None
        try:
            sep = csv.Sniffer().sniff(sample_text, delimiters=delimiters).delimiter
        except Exception:
            sep = None

        try:
            if sep:
                return pd.read_csv(
                    file_path,
                    header=None,
                    keep_default_na=False,
                    encoding=encoding,
                    sep=sep,
                    skip_blank_lines=should_skip_blank_lines(sep),
                )
            return pd.read_csv(
                file_path,
                header=None,
                keep_default_na=False,
                encoding=encoding,
                sep=None,
                engine="python",
                skip_blank_lines=should_skip_blank_lines(),
            )
        except Exception:
            continue

    return pd.read_csv(
        file_path,
        header=None,
        keep_default_na=False,
        skip_blank_lines=should_skip_blank_lines(),
    )


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
        try:
            image_data = img._data()

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
        except Exception as exc:
            context.logger.error(f"Sheet→MD conversion failed: {exc}")
            return ConversionResult(
                task_id=task_id,
                success=False,
                error=ConversionErrorInfo(
                    error_type="conversion_failed",
                    message=str(exc),
                    diagnostic_code="SHEET2MD-PARSE-ERROR",
                ),
                diagnostics=[
                    ConversionDiagnostic(
                        level="error",
                        message=f"Failed to parse spreadsheet: {exc}",
                        code="SHEET2MD-PARSE-ERROR",
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
        file_stem = Path(input_path).stem
        stats: dict[str, int] = {"sheets": 0, "rows": 0, "cols": 0, "blocks": 0}

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
        stats: dict[str, int],
        context: ConverterContext,
    ) -> tuple[str, dict[str, int]]:
        """Convert a CSV file to Markdown."""
        context.progress.report_progress(20.0, "Reading CSV...")

        md_content += f"# {file_stem}\n\n"

        df = _read_csv_flexible(input_path, source_format)
        stats["sheets"] = 1
        stats["rows"] = df.shape[0]
        stats["cols"] = df.shape[1]

        blocks = _find_blocks(df)
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
                block_md = data.to_markdown(index=False, headers=headers) if hasattr(data, "to_markdown") else ""
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
        stats: dict[str, int],
        context: ConverterContext,
        options: dict | None = None,
    ) -> tuple[str, dict[str, int]]:
        """Convert an XLSX file to Markdown.

        When *keep_images* is ``True``, embedded images are extracted from
        each worksheet, written to staging, and referenced in the Markdown
        output (F-G6-012).
        """
        if options is None:
            options = context.request.options
        from docwen_plugin_spreadsheet.csv_xlsx.converter import _load_admitted_xlsx

        context.progress.report_progress(20.0, "Opening workbook...")

        wb = _load_admitted_xlsx(input_path, data_only=True)
        stats["sheets"] = len(wb.sheetnames)

        total_images_extracted = 0
        total_image_extraction_losses = 0

        total_sheets = len(wb.sheetnames)
        for sheet_idx, sheet_name in enumerate(wb.sheetnames):
            context.cancellation.check()
            progress = 20.0 + 60.0 * (sheet_idx / max(total_sheets, 1))
            context.progress.report_progress(progress, f"Processing sheet: {sheet_name}")

            md_content += f"# {sheet_name}\n\n"
            ws = wb[sheet_name]

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

                    img_seq = 0
                    main_stem = os.path.splitext(os.path.basename(str(input_path)))[0]

                    for img_info in sheet_images:
                        # Build image markdown link
                        img_path = img_info["staging_path"] if image_mode == "base64" else img_info["suggested_name"]
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

                            outcome = run_ocr_outcome(
                                img_info["staging_path"],
                                source_format=detect_content_format(img_info["staging_path"]).format,
                                ocr_language=ocr_language,
                                current_locale=current_locale,
                            )
                            _report_ocr_best_effort(
                                context.progress,
                                outcome.status,
                                location=(f"{sheet_name}:{img_info.get('suggested_name', img_info['staging_path'])}"),
                            )
                            ocr_text = outcome.recognized_text
                            if ocr_text:
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
                                    ocr_block = f"\n> {request_policy.ocr_blockquote_title}\n> "
                                    ocr_block += "\n> ".join(ocr_text.splitlines()) + "\n"
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

            df = _worksheet_to_dataframe(ws, table_merge_strategy=merge_strategy)

            # H6: Inject image/OCR content into the correct cell positions
            if image_positions:
                # Extend dataframe to accommodate images outside data range
                max_img_row = max(r for (r, _c) in image_positions)
                max_img_col = max(c for (_r, c) in image_positions)
                need_rows = max(df.shape[0], max_img_row + 1)
                need_cols = max(df.shape[1], max_img_col + 1)
                if need_rows > df.shape[0] or need_cols > df.shape[1]:
                    df = df.reindex(index=range(need_rows), columns=range(need_cols), fill_value="")
                    df = df.fillna("")

                for (r, c), cell_contents in image_positions.items():
                    if r < df.shape[0] and c < df.shape[1]:
                        cell_val = df.iat[r, c]
                        inject_text = " ".join(cell_contents)
                        if cell_val and str(cell_val).strip():
                            df.iat[r, c] = f"{cell_val}<br>{inject_text}"
                        else:
                            df.iat[r, c] = inject_text

            stats["rows"] += df.shape[0]
            stats["cols"] = max(stats["cols"], df.shape[1])

            blocks = _find_blocks(df)
            stats["blocks"] += len(blocks)

            if not blocks:
                md_content += " (this sheet is empty)\n\n"
            else:
                for block in blocks:
                    block = _process_cell_newlines(block)

                    if block.shape[0] > 0:
                        headers = [
                            str(h) if h is not None and str(h).strip() != "" else f"Col{j}"
                            for j, h in enumerate(block.iloc[0].tolist())
                        ]
                        data = block.iloc[1:]
                        block_md = (
                            data.to_markdown(index=False, headers=headers) if hasattr(data, "to_markdown") else ""
                        )
                    else:
                        block_md = ""

                    md_content += (block_md or "") + "\n\n"

        wb.close()
        stats["images"] = total_images_extracted
        if total_image_extraction_losses:
            stats["image_extraction_loss_count"] = total_image_extraction_losses
        return md_content.strip() + "\n", stats

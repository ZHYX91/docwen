"""Output-specific image rendering for structured clipboard documents."""

from __future__ import annotations

import csv
from io import BytesIO
from pathlib import Path
from typing import Any

from docwen_core.models.clipboard_document import (
    ClipboardDocument,
    ClipboardHardBreak,
    ClipboardImageRef,
    ClipboardParagraph,
    ClipboardText,
    iter_clipboard_inlines,
)
from docwen_core.models.result import ConversionDiagnostic
from docwen_plugin_markdown.structured_clipboard_images import (
    BoundClipboardImage,
    ClipboardImageOccurrenceInfo,
    MarkdownImagePlan,
)


def image_resource_diagnostics(document: ClipboardDocument) -> list[ConversionDiagnostic]:
    missing_count = sum(
        isinstance(inline, ClipboardImageRef) and inline.resource_id is None
        for inline in iter_clipboard_inlines(document)
    )
    renderable_ids = {resource.resource_id for resource in document.resources if resource.media_type == "image/png"}
    not_rendered_count = sum(
        isinstance(inline, ClipboardImageRef)
        and inline.resource_id is not None
        and inline.resource_id not in renderable_ids
        for inline in iter_clipboard_inlines(document)
    )
    diagnostics = []
    if missing_count:
        diagnostics.append(
            ConversionDiagnostic(
                level="warning",
                code="CLIPBOARD-IMAGE-RESOURCE-UNAVAILABLE",
                message=(
                    f"{missing_count} clipboard image occurrence(s) had no verified resource bytes "
                    "and were kept as text placeholders."
                ),
            )
        )
    if not_rendered_count:
        diagnostics.append(
            ConversionDiagnostic(
                level="warning",
                code="CLIPBOARD-IMAGE-RESOURCE-NOT-RENDERED",
                message=(
                    f"{not_rendered_count} clipboard image occurrence(s) had verified opaque resources "
                    "without a supported PNG representation and were kept as text placeholders."
                ),
            )
        )
    return diagnostics


def image_placeholder(image: ClipboardImageRef) -> str:
    state = "not rendered" if image.resource_id is not None else "unavailable"
    return f"[Image {state}: {image.alt}]" if image.alt else f"[Image {state}]"


def write_docx_paragraph(
    container,
    paragraph: ClipboardParagraph,
    resources: dict[str, BoundClipboardImage],
):
    """Render text/break/image occurrences in original paragraph order."""

    from docx.shared import Emu

    output = container.add_paragraph()
    run = output.add_run()
    for inline in paragraph.inlines:
        if isinstance(inline, ClipboardText):
            # add_run converts literal tabs/newlines into native w:tab/w:br.
            # add_text would leave them inside w:t and lose the visible breaks.
            run = output.add_run(inline.value)
        elif isinstance(inline, ClipboardHardBreak):
            run.add_break()
        elif inline.resource_id is None:
            run.add_text(image_placeholder(inline))
        else:
            bound = resources.get(inline.resource_id)
            if bound is None:
                run.add_text(image_placeholder(inline))
                continue
            kwargs = {}
            if inline.extent_cx_emu is not None and inline.extent_cy_emu is not None:
                kwargs = {"width": Emu(inline.extent_cx_emu), "height": Emu(inline.extent_cy_emu)}
            run.add_picture(BytesIO(bound.path.read_bytes()), **kwargs)
    return output


def markdown_paragraph_blocks(
    paragraph: ClipboardParagraph,
    plan: MarkdownImagePlan,
    resources: dict[str, BoundClipboardImage],
    fence,
) -> list[str]:
    """Keep authored text literal while placing image references at occurrences."""

    output: list[str] = []
    text: list[str] = []

    def flush() -> None:
        value = "".join(text)
        text.clear()
        if value:
            output.extend([fence(value), ""])

    for inline in paragraph.inlines:
        if isinstance(inline, ClipboardText):
            text.append(inline.value)
        elif isinstance(inline, ClipboardHardBreak):
            text.append("\n")
        else:
            flush()
            output.extend([plan.reference(inline, resources), ""])
    flush()
    if not output:
        output.extend([fence(""), ""])
    return output


def add_xlsx_images(
    projection,
    occurrences: tuple[ClipboardImageOccurrenceInfo, ...],
    resources: dict[str, BoundClipboardImage],
    table_sheets: dict[str, Any],
) -> list[tuple[str, ...]]:
    """Add one drawing per bound occurrence and return visible semantics rows."""

    from openpyxl.drawing.image import Image as XlsxImage
    from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
    from openpyxl.drawing.xdr import XDRPositiveSize2D
    from openpyxl.utils import coordinate_to_tuple, get_column_letter

    semantics = projection.workbook.create_sheet("Image Semantics")
    headers = (
        "Ordinal",
        "State",
        "Resource ID",
        "Logical Path",
        "Table",
        "Cell Anchor",
        "Drawing Sheet",
        "Drawing Anchor",
        "Extent Cx Emu",
        "Extent Cy Emu",
    )
    rows: list[tuple[str, ...]] = []
    top_level_row = 2
    for occurrence in occurrences:
        image = occurrence.image
        if image.resource_id is None:
            rows.append(
                (
                    str(occurrence.ordinal),
                    "missing",
                    "",
                    "",
                    occurrence.table_id,
                    occurrence.cell_anchor,
                    "",
                    "",
                    str(image.extent_cx_emu or ""),
                    str(image.extent_cy_emu or ""),
                )
            )
            continue
        bound = resources.get(image.resource_id)
        if bound is None:
            rows.append(
                (
                    str(occurrence.ordinal),
                    "not_rendered",
                    image.resource_id,
                    "",
                    occurrence.table_id,
                    occurrence.cell_anchor,
                    "",
                    "",
                    str(image.extent_cx_emu or ""),
                    str(image.extent_cy_emu or ""),
                )
            )
            continue

        if occurrence.table_id:
            sheet = table_sheets[occurrence.table_id]
            raw = occurrence.cell_anchor.removeprefix("R")
            row_part, column_part = raw.split("C", 1)
            drawing_anchor = f"{get_column_letter(int(column_part))}{int(row_part)}"
        else:
            sheet = projection.order_sheet
            drawing_anchor = f"I{top_level_row}"
            top_level_row += 2

        drawing = XlsxImage(str(bound.path))
        if image.extent_cx_emu is not None and image.extent_cy_emu is not None:
            row, column = coordinate_to_tuple(drawing_anchor)
            drawing.anchor = OneCellAnchor(
                _from=AnchorMarker(col=column - 1, row=row - 1),
                ext=XDRPositiveSize2D(cx=image.extent_cx_emu, cy=image.extent_cy_emu),
            )
        elif bound.resource.pixel_width is not None and bound.resource.pixel_height is not None:
            drawing.width = bound.resource.pixel_width
            drawing.height = bound.resource.pixel_height
        if image.extent_cx_emu is not None and image.extent_cy_emu is not None:
            sheet.add_image(drawing)
        else:
            sheet.add_image(drawing, drawing_anchor)
        rows.append(
            (
                str(occurrence.ordinal),
                "bound",
                image.resource_id,
                bound.resource.logical_path,
                occurrence.table_id,
                occurrence.cell_anchor,
                sheet.title,
                drawing_anchor,
                str(image.extent_cx_emu or ""),
                str(image.extent_cy_emu or ""),
            )
        )

    from docwen_plugin_markdown.structured_clipboard_xlsx import write_string_matrix

    write_string_matrix(semantics, headers, rows)
    return rows


def write_csv_image_semantics(
    path: Path,
    occurrences: tuple[ClipboardImageOccurrenceInfo, ...],
    resources: dict[str, BoundClipboardImage],
) -> None:
    """Write honest image occurrence/resource projection for CSV output."""

    with path.open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.writer(stream)
        writer.writerow(
            (
                "Ordinal",
                "State",
                "Resource ID",
                "Logical Path",
                "Suggested Resource",
                "Table",
                "Cell Anchor",
                "Extent Cx Emu",
                "Extent Cy Emu",
            )
        )
        for occurrence in occurrences:
            image = occurrence.image
            bound = resources.get(image.resource_id or "")
            writer.writerow(
                (
                    occurrence.ordinal,
                    "bound" if bound is not None else ("not_rendered" if image.resource_id is not None else "missing"),
                    image.resource_id or "",
                    bound.resource.logical_path if bound is not None else "",
                    bound.suggested_name if bound is not None else "",
                    occurrence.table_id,
                    occurrence.cell_anchor,
                    image.extent_cx_emu or "",
                    image.extent_cy_emu or "",
                )
            )


__all__ = [
    "add_xlsx_images",
    "image_placeholder",
    "image_resource_diagnostics",
    "markdown_paragraph_blocks",
    "write_csv_image_semantics",
    "write_docx_paragraph",
]

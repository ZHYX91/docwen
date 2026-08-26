"""Generate review-draft GUI screenshot contact sheets.

The public screenshot inventory under ``docs/assets/screenshots`` is the source.
This helper builds temporary side-by-side review sheets under ``build`` without
adding generated review artifacts to the documentation tree.
"""

from __future__ import annotations

import argparse
import sys
from dataclasses import dataclass
from pathlib import Path

try:
    from PIL import Image, ImageDraw, ImageFont
except ImportError as exc:  # pragma: no cover - exercised by dependency gate
    raise SystemExit("Pillow is required to generate GUI contact sheets") from exc


PROJECT_ROOT = Path(__file__).resolve().parents[1]
BASELINE_ROOT = PROJECT_ROOT / "docs" / "assets" / "screenshots"
OUTPUT_ROOT = PROJECT_ROOT / "build" / "gui-contact-sheets"

THUMB_WIDTH = 360
CELL_PADDING = 12
HEADER_HEIGHT = 34
GAP = 14
LABEL_COLOR = (31, 41, 55)
MUTED_COLOR = (107, 114, 128)
BORDER_COLOR = (209, 213, 219)
BG_COLOR = (255, 255, 255)
MISSING_BG = (243, 244, 246)


@dataclass(frozen=True)
class Cell:
    label: str
    image: str | None
    note: str = ""


@dataclass(frozen=True)
class SheetSpec:
    name: str
    title: str
    rows: tuple[tuple[Cell, ...], ...]


def _cell(label: str, image: str | None, note: str = "") -> Cell:
    return Cell(label=label, image=image, note=note)


SHEETS: dict[str, SheetSpec] = {
    "core-scenes-light": SheetSpec(
        name="core-scenes-light",
        title="Core Scenes - Tk / old PySide6 / current light",
        rows=(
            (
                _cell("Tk main", "old-ref/tk-existing/main-blank.png"),
                _cell("PySide6 main", "old-ref/pyside6-fresh/main-light.png"),
                _cell("Current main", "current-new-expanded/main-light.png"),
            ),
            (
                _cell("Tk Settings Text", "old-ref/tk-existing/settings-tab-text.png"),
                _cell("PySide6 Settings Text", "old-ref/pyside6-fresh/settings-tab-text-light.png"),
                _cell("Current Settings Text", "current-new-expanded/settings-tab-text-light.png"),
            ),
            (
                _cell("Tk batch", "old-ref/tk-existing/batch.png"),
                _cell("PySide6 batch", "old-ref/pyside6-fresh/batch-light.png"),
                _cell("Current batch", "current-new-expanded/batch-light.png"),
            ),
            (
                _cell("Tk About", "old-ref/tk-existing/about-normalized.png"),
                _cell("PySide6 About", "old-ref/pyside6-fresh/about-light.png"),
                _cell("Current About", "current-new-expanded/about-light.png"),
            ),
            (
                _cell("Tk document", "old-ref/tk-existing/document.png"),
                _cell("PySide6 document", "old-ref/pyside6-fresh/conversion-document-light.png"),
                _cell("Current document", "current-new-expanded/conversion-document-light.png"),
            ),
            (
                _cell("Tk spreadsheet", "old-ref/tk-existing/spreadsheet.png"),
                _cell("PySide6 conversion", "old-ref/pyside6-fresh/conversion-light.png"),
                _cell("Current spreadsheet", "current-new-expanded/conversion-spreadsheet-light.png"),
            ),
            (
                _cell("Tk image", "old-ref/tk-existing/image.png"),
                _cell("PySide6 image", "old-ref/pyside6-fresh/conversion-image-light.png"),
                _cell("Current image", "current-new-expanded/conversion-image-light.png"),
            ),
            (
                _cell("Tk layout", "old-ref/tk-existing/layout.png"),
                _cell("PySide6 conversion", "old-ref/pyside6-fresh/conversion-light.png"),
                _cell("Current layout", "current-new-expanded/conversion-layout-light.png"),
            ),
        ),
    ),
    "main-window-empty-template": SheetSpec(
        name="main-window-empty-template",
        title="Main Window and Markdown Template Workflow",
        rows=(
            (
                _cell("Tk main", "old-ref/tk-existing/main-blank.png"),
                _cell("PySide6 main", "old-ref/pyside6-fresh/main-light.png"),
                _cell("Current main light", "current-new-expanded/main-light.png"),
                _cell("Current main dark", "current-new-expanded/main-dark.png"),
            ),
            (
                _cell("Tk Markdown", "old-ref/tk-existing/markdown.png"),
                _cell("PySide6 template", "old-ref/pyside6-fresh/template-text-light.png"),
                _cell("Current template light", "current-new-expanded/template-light.png"),
                _cell("Current template dark", "current-new-expanded/template-dark.png"),
            ),
        ),
    ),
    "settings-tabs-light": SheetSpec(
        name="settings-tabs-light",
        title="Settings Tabs - Tk / old PySide6 / current light",
        rows=tuple(
            (
                _cell(f"Tk {tk_label}", tk_image),
                _cell(f"PySide6 {tab}", f"old-ref/pyside6-fresh/settings-tab-{tab}-light.png"),
                _cell(f"Current {tab}", f"current-new-expanded/settings-tab-{tab}-light.png"),
            )
            for tab, tk_label, tk_image in (
                ("general", "general", "old-ref/tk-existing/settings-tab-general.png"),
                ("text", "text", "old-ref/tk-existing/settings-tab-text.png"),
                ("proofread", "Proofread", None),
                ("export", "export", "old-ref/tk-existing/settings-tab-export.png"),
                ("document", "document", "old-ref/tk-existing/settings-tab-document.png"),
                ("spreadsheet", "spreadsheet", "old-ref/tk-existing/settings-tab-spreadsheet.png"),
                ("image", "image", "old-ref/tk-existing/settings-tab-image.png"),
                ("layout", "layout", "old-ref/tk-existing/settings-tab-layout.png"),
                ("other", "other", "old-ref/tk-existing/settings-tab-other.png"),
                ("link", "link", "old-ref/tk-existing/settings-tab-link.png"),
                ("formatting", "formatting", "old-ref/tk-existing/settings-tab-formatting.png"),
                ("output", "output", "old-ref/tk-existing/settings-tab-output.png"),
                ("logging", "logging", "old-ref/tk-existing/settings-tab-logging.png"),
            )
        ),
    ),
    "batch-list-and-runtime": SheetSpec(
        name="batch-list-and-runtime",
        title="Batch List and Runtime States",
        rows=(
            (
                _cell("Tk batch", "old-ref/tk-existing/batch.png"),
                _cell("PySide6 batch", "old-ref/pyside6-fresh/batch-light.png"),
                _cell("Current batch light", "current-new-expanded/batch-light.png"),
                _cell("Current batch dark", "current-new-expanded/batch-dark.png"),
            ),
            (
                _cell("Old progress", None, "No matching old runtime screenshot"),
                _cell("Current progress light", "current-new-expanded/progress-light.png"),
                _cell("Current progress dark", "current-new-expanded/progress-dark.png"),
            ),
            (
                _cell("Old completed", None, "No matching old runtime screenshot"),
                _cell("Current completed light", "current-new-expanded/completed-light.png"),
                _cell("Current completed dark", "current-new-expanded/completed-dark.png"),
            ),
            (
                _cell("Old failed", None, "No matching old runtime screenshot"),
                _cell("Current failed light", "current-new-expanded/failed-light.png"),
                _cell("Current failed dark", "current-new-expanded/failed-dark.png"),
            ),
            (
                _cell("Old cancelled", None, "No matching old runtime screenshot"),
                _cell("Current cancelled light", "current-new-expanded/cancelled-light.png"),
                _cell("Current cancelled dark", "current-new-expanded/cancelled-dark.png"),
            ),
        ),
    ),
    "conversion-flows-light-dark": SheetSpec(
        name="conversion-flows-light-dark",
        title="Conversion Flows - old light / current light-dark",
        rows=(
            (
                _cell("Tk document", "old-ref/tk-existing/document.png"),
                _cell("PySide6 document", "old-ref/pyside6-fresh/conversion-document-light.png"),
                _cell("Current document light", "current-new-expanded/conversion-document-light.png"),
                _cell("Current document dark", "current-new-expanded/conversion-document-dark.png"),
            ),
            (
                _cell("Tk spreadsheet", "old-ref/tk-existing/spreadsheet.png"),
                _cell("PySide6 conversion", "old-ref/pyside6-fresh/conversion-light.png"),
                _cell("Current spreadsheet light", "current-new-expanded/conversion-spreadsheet-light.png"),
                _cell("Current spreadsheet dark", "current-new-expanded/conversion-spreadsheet-dark.png"),
            ),
            (
                _cell("Tk image", "old-ref/tk-existing/image.png"),
                _cell("PySide6 image", "old-ref/pyside6-fresh/conversion-image-light.png"),
                _cell("Current image light", "current-new-expanded/conversion-image-light.png"),
                _cell("Current image dark", "current-new-expanded/conversion-image-dark.png"),
            ),
            (
                _cell("Tk layout", "old-ref/tk-existing/layout.png"),
                _cell("PySide6 conversion", "old-ref/pyside6-fresh/conversion-light.png"),
                _cell("Current layout light", "current-new-expanded/conversion-layout-light.png"),
                _cell("Current layout dark", "current-new-expanded/conversion-layout-dark.png"),
            ),
        ),
    ),
    "about-help": SheetSpec(
        name="about-help",
        title="About and Help Surface",
        rows=(
            (
                _cell("Tk About", "old-ref/tk-existing/about-normalized.png"),
                _cell("PySide6 About light", "old-ref/pyside6-fresh/about-light.png"),
                _cell("PySide6 About dark", "old-ref/pyside6-fresh/about-dark.png"),
                _cell("Current About light", "current-new-expanded/about-light.png"),
            ),
            (
                _cell("Tk dark", None, "No old Tk dark About screenshot"),
                _cell("Current About bottom light", "current-new-expanded/about-bottom-light.png"),
                _cell("Current About dark", "current-new-expanded/about-dark.png"),
                _cell("Current About bottom dark", "current-new-expanded/about-bottom-dark.png"),
            ),
        ),
    ),
    "runtime-states-current-light-dark": SheetSpec(
        name="runtime-states-current-light-dark",
        title="Current Runtime States - light and dark",
        rows=tuple(
            (
                _cell(f"{state} light", f"current-new-expanded/{state}-light.png"),
                _cell(f"{state} dark", f"current-new-expanded/{state}-dark.png"),
            )
            for state in ("status", "progress", "completed", "failed", "cancelled")
        ),
    ),
}


def _font(size: int, *, bold: bool = False) -> ImageFont.ImageFont:
    names = (
        "arialbd.ttf" if bold else "arial.ttf",
        "segoeuib.ttf" if bold else "segoeui.ttf",
    )
    for name in names:
        try:
            return ImageFont.truetype(name, size)
        except OSError:
            continue
    return ImageFont.load_default()


def _wrap_text(text: str, max_chars: int = 44) -> list[str]:
    if not text:
        return []
    words = text.split()
    lines: list[str] = []
    current = ""
    for word in words:
        candidate = word if not current else f"{current} {word}"
        if len(candidate) <= max_chars:
            current = candidate
            continue
        if current:
            lines.append(current)
        current = word
    if current:
        lines.append(current)
    return lines


def _thumbnail(path: Path) -> Image.Image:
    with Image.open(path) as source:
        source.load()
        image = source.convert("RGB")
    ratio = THUMB_WIDTH / image.width
    thumb_height = max(1, int(image.height * ratio))
    return image.resize((THUMB_WIDTH, thumb_height), Image.Resampling.LANCZOS)


def _prepare_rows(spec: SheetSpec, baseline_root: Path) -> list[list[tuple[Cell, Image.Image | None]]]:
    rows: list[list[tuple[Cell, Image.Image | None]]] = []
    for row in spec.rows:
        prepared_row: list[tuple[Cell, Image.Image | None]] = []
        for cell in row:
            image = None
            if cell.image is not None:
                image_path = baseline_root / cell.image
                if not image_path.exists():
                    raise FileNotFoundError(f"missing contact-sheet source image: {cell.image}")
                image = _thumbnail(image_path)
            prepared_row.append((cell, image))
        rows.append(prepared_row)
    return rows


def render_sheet(spec: SheetSpec, output_path: Path, baseline_root: Path = BASELINE_ROOT) -> Path:
    prepared_rows = _prepare_rows(spec, baseline_root)
    label_font = _font(16, bold=True)
    note_font = _font(14)
    title_font = _font(22, bold=True)
    max_columns = max(len(row) for row in prepared_rows)
    cell_width = THUMB_WIDTH + (CELL_PADDING * 2)

    row_heights: list[int] = []
    for row in prepared_rows:
        max_thumb_height = max((image.height for _, image in row if image is not None), default=180)
        note_lines = max((len(_wrap_text(cell.note)) for cell, _ in row), default=0)
        row_heights.append(HEADER_HEIGHT + max_thumb_height + CELL_PADDING + (note_lines * 18) + CELL_PADDING)

    width = (cell_width * max_columns) + (GAP * (max_columns + 1))
    height = 70 + sum(row_heights) + (GAP * (len(row_heights) + 1))
    sheet = Image.new("RGB", (width, height), BG_COLOR)
    draw = ImageDraw.Draw(sheet)
    draw.text((GAP, 18), spec.title, fill=LABEL_COLOR, font=title_font)

    y = 70
    for row, row_height in zip(prepared_rows, row_heights, strict=True):
        x = GAP
        for cell, image in row:
            draw.rounded_rectangle(
                (x, y, x + cell_width, y + row_height),
                radius=8,
                outline=BORDER_COLOR,
                width=1,
                fill=BG_COLOR,
            )
            draw.text((x + CELL_PADDING, y + 8), cell.label, fill=LABEL_COLOR, font=label_font)
            image_top = y + HEADER_HEIGHT
            if image is not None:
                draw.rectangle(
                    (
                        x + CELL_PADDING - 1,
                        image_top - 1,
                        x + CELL_PADDING + image.width + 1,
                        image_top + image.height + 1,
                    ),
                    outline=BORDER_COLOR,
                )
                sheet.paste(image, (x + CELL_PADDING, image_top))
                note_y = image_top + image.height + CELL_PADDING
            else:
                placeholder = (
                    x + CELL_PADDING,
                    image_top,
                    x + CELL_PADDING + THUMB_WIDTH,
                    image_top + 180,
                )
                draw.rectangle(placeholder, outline=BORDER_COLOR, fill=MISSING_BG)
                draw.text(
                    (placeholder[0] + 14, placeholder[1] + 76), "Not available", fill=MUTED_COLOR, font=label_font
                )
                note_y = placeholder[3] + CELL_PADDING
            for line in _wrap_text(cell.note):
                draw.text((x + CELL_PADDING, note_y), line, fill=MUTED_COLOR, font=note_font)
                note_y += 18
            x += cell_width + GAP
        y += row_height + GAP

    output_path.parent.mkdir(parents=True, exist_ok=True)
    sheet.save(output_path)
    return output_path


def _parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate DocWen GUI visual parity contact sheets.")
    parser.add_argument(
        "--sheet",
        action="append",
        choices=sorted(SHEETS),
        help="Sheet name to generate. Repeat for multiple sheets. Defaults to all sheets.",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=OUTPUT_ROOT,
        help="Directory for generated PNG files.",
    )
    parser.add_argument(
        "--baseline-root",
        type=Path,
        default=BASELINE_ROOT,
        help="Current GUI screenshot source root.",
    )
    parser.add_argument("--list", action="store_true", help="List available sheet names and exit.")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = _parse_args(sys.argv[1:] if argv is None else argv)
    if args.list:
        for name in sorted(SHEETS):
            print(name)
        return 0

    sheet_names = args.sheet or sorted(SHEETS)
    for name in sheet_names:
        spec = SHEETS[name]
        output_path = args.output_dir / f"{spec.name}-contact-sheet.png"
        print(render_sheet(spec, output_path, args.baseline_root))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

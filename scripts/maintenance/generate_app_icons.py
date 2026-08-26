"""Generate committed application icon derivatives from ``assets/icon.svg``."""

from __future__ import annotations

import argparse
import io
import struct
from pathlib import Path

from PIL import Image
from PySide6.QtCore import QRectF, Qt
from PySide6.QtGui import QImage, QPainter
from PySide6.QtSvg import QSvgRenderer

ROOT = Path(__file__).resolve().parents[2]
SVG_PATH = ROOT / "assets" / "icon.svg"
PNG_PATH = ROOT / "assets" / "icon.png"
ICO_PATH = ROOT / "assets" / "icon.ico"
PNG_SIZE = 256
ICO_SIZES = (16, 24, 32, 48, 64, 72, 80, 96, 128, 256)


def _render_png(svg_bytes: bytes, size: int) -> bytes:
    renderer = QSvgRenderer(svg_bytes)
    if not renderer.isValid():
        raise ValueError(f"Invalid SVG source: {SVG_PATH}")

    canvas = QImage(size, size, QImage.Format.Format_ARGB32_Premultiplied)
    canvas.fill(Qt.GlobalColor.transparent)
    painter = QPainter(canvas)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing)
    painter.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform)
    renderer.render(painter, QRectF(0, 0, size, size))
    painter.end()

    rgba = canvas.convertToFormat(QImage.Format.Format_RGBA8888)
    pixels = bytes(rgba.constBits())
    image = Image.frombytes(
        "RGBA",
        (size, size),
        pixels,
        "raw",
        "RGBA",
        rgba.bytesPerLine(),
        1,
    )
    output = io.BytesIO()
    image.save(output, format="PNG", compress_level=9, optimize=False)
    return output.getvalue()


def _build_ico(png_frames: list[tuple[int, bytes]]) -> bytes:
    header = struct.pack("<HHH", 0, 1, len(png_frames))
    offset = len(header) + (16 * len(png_frames))
    entries: list[bytes] = []
    payloads: list[bytes] = []

    for size, payload in png_frames:
        dimension = 0 if size == 256 else size
        entries.append(
            struct.pack(
                "<BBBBHHII",
                dimension,
                dimension,
                0,
                0,
                1,
                32,
                len(payload),
                offset,
            )
        )
        payloads.append(payload)
        offset += len(payload)

    return b"".join((header, *entries, *payloads))


def _generated_assets() -> tuple[bytes, bytes]:
    svg_bytes = SVG_PATH.read_bytes()
    frames = [(size, _render_png(svg_bytes, size)) for size in ICO_SIZES]
    png = next(payload for size, payload in frames if size == PNG_SIZE)
    return png, _build_ico(frames)


def _check_asset(path: Path, expected: bytes) -> bool:
    if not path.is_file() or path.read_bytes() != expected:
        print(f"stale generated icon: {path.relative_to(ROOT)}")
        return False
    return True


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--check",
        action="store_true",
        help="fail when committed PNG/ICO assets differ from the SVG source",
    )
    args = parser.parse_args()

    png, ico = _generated_assets()
    if args.check:
        return 0 if _check_asset(PNG_PATH, png) and _check_asset(ICO_PATH, ico) else 1

    PNG_PATH.write_bytes(png)
    ICO_PATH.write_bytes(ico)
    print(f"generated {PNG_PATH.relative_to(ROOT)} and {ICO_PATH.relative_to(ROOT)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

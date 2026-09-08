"""Generate committed application icon derivatives from ``assets/icon.svg``."""

from __future__ import annotations

import argparse
import io
import struct
import sys
from pathlib import Path

from PIL import Image
from PySide6.QtCore import QRectF, Qt
from PySide6.QtGui import QImage, QPainter
from PySide6.QtSvg import QSvgRenderer

ROOT = Path(__file__).resolve().parents[2]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from scripts.icon_pixels import unpremultiply_rgba  # noqa: E402 - standalone script bootstrap

SVG_PATH = ROOT / "assets" / "icon.svg"
PNG_PATH = ROOT / "assets" / "icon.png"
ICO_PATH = ROOT / "assets" / "icon.ico"
PNG_SIZE = 256
ICO_SIZES = (16, 24, 32, 48, 64, 72, 80, 96, 128, 256)


def _render_png(svg_bytes: bytes, size: int) -> bytes:
    renderer = QSvgRenderer(svg_bytes)
    if not renderer.isValid():
        raise ValueError(f"Invalid SVG source: {SVG_PATH}")

    canvas = QImage(size, size, QImage.Format.Format_RGBA8888_Premultiplied)
    canvas.fill(Qt.GlobalColor.transparent)
    painter = QPainter(canvas)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing)
    painter.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform)
    renderer.render(painter, QRectF(0, 0, size, size))
    painter.end()

    pixels = unpremultiply_rgba(bytes(canvas.constBits()))
    image = Image.frombytes(
        "RGBA",
        (size, size),
        pixels,
        "raw",
        "RGBA",
        canvas.bytesPerLine(),
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


def _decoded_frames(data: bytes) -> tuple[str | None, dict[tuple[int, int], bytes]]:
    with Image.open(io.BytesIO(data)) as image:
        if image.format == "ICO":
            return image.format, {
                size: image.ico.getimage(size).convert("RGBA").tobytes() for size in image.ico.sizes()
            }
        return image.format, {image.size: image.convert("RGBA").tobytes()}


def _check_asset(path: Path, expected: bytes) -> bool:
    # Pillow wheels use different PNG compression libraries across platforms.
    # Require exact decoded pixels for every frame, without a visual tolerance.
    if not path.is_file() or _decoded_frames(path.read_bytes()) != _decoded_frames(expected):
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

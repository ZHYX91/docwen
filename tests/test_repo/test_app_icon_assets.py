"""Contract tests for the canonical DocWen application icon and derivatives."""

from __future__ import annotations

import io
import math
import re
import struct
import subprocess
import sys
from pathlib import Path
from xml.etree import ElementTree

import pytest
from PIL import Image

pytestmark = pytest.mark.contract

ROOT = Path(__file__).resolve().parents[2]
SVG_PATH = ROOT / "assets" / "icon.svg"
PNG_PATH = ROOT / "assets" / "icon.png"
ICO_PATH = ROOT / "assets" / "icon.ico"
GENERATOR = ROOT / "scripts" / "maintenance" / "generate_app_icons.py"
EXPECTED_ICO_SIZES = (16, 24, 32, 48, 64, 72, 80, 96, 128, 256)


def _element(root: ElementTree.Element, element_id: str) -> ElementTree.Element:
    for element in root.iter():
        if element.get("id") == element_id:
            return element
    raise AssertionError(f"missing SVG element: {element_id}")


def _point_reflected_across_line(
    point: tuple[float, float],
    start: tuple[float, float],
    end: tuple[float, float],
) -> tuple[float, float]:
    dx = end[0] - start[0]
    dy = end[1] - start[1]
    scale = ((point[0] - start[0]) * dx + (point[1] - start[1]) * dy) / (dx * dx + dy * dy)
    projected = (start[0] + scale * dx, start[1] + scale * dy)
    return (2 * projected[0] - point[0], 2 * projected[1] - point[1])


def _ico_frames(data: bytes) -> list[tuple[int, bytes]]:
    reserved, image_type, count = struct.unpack_from("<HHH", data)
    assert (reserved, image_type) == (0, 1)
    frames: list[tuple[int, bytes]] = []
    for index in range(count):
        width, height, colors, reserved_byte, planes, bits, length, offset = struct.unpack_from(
            "<BBBBHHII", data, 6 + index * 16
        )
        assert width == height
        assert colors == reserved_byte == 0
        assert (planes, bits) == (1, 32)
        frames.append((256 if width == 0 else width, data[offset : offset + length]))
    return frames


def test_svg_uses_approved_geometry_and_palette() -> None:
    root = ElementTree.parse(SVG_PATH).getroot()
    assert root.get("viewBox") == "0 0 128 128"

    background = _element(root, "background")
    shafts = _element(root, "arrow-shafts")
    heads = _element(root, "arrow-heads")
    assert background.get("fill") == shafts.get("stroke") == heads.get("stroke") == "#2F6FEB"
    assert shafts.get("stroke-width") == "5.2"
    assert heads.get("stroke-width") == "4.6"


def test_fold_is_an_exact_sixty_degree_mirror() -> None:
    root = ElementTree.parse(SVG_PATH).getroot()
    fold = _element(root, "folded-corner")
    coordinates = [float(value) for value in re.findall(r"-?\d+(?:\.\d+)?", fold.attrib["d"])]
    crease_start = (coordinates[0], coordinates[1])
    folded_tip = (coordinates[2], coordinates[3])
    crease_end = (coordinates[4], coordinates[5])
    source_values = [float(value) for value in fold.attrib["data-source-corner"].split()]
    assert len(source_values) == 2
    source_corner = (source_values[0], source_values[1])

    angle = math.degrees(math.atan2(crease_end[1] - crease_start[1], crease_end[0] - crease_start[0]))
    assert angle == pytest.approx(60.0, abs=1e-6)
    reflected = _point_reflected_across_line(source_corner, crease_start, crease_end)
    assert folded_tip == pytest.approx(reflected, abs=1e-6)


def test_generated_png_and_ico_are_current() -> None:
    subprocess.run([sys.executable, str(GENERATOR), "--check"], cwd=ROOT, check=True)

    with Image.open(PNG_PATH) as png:
        assert png.size == (256, 256)
        assert png.mode == "RGBA"

    frames = _ico_frames(ICO_PATH.read_bytes())
    assert tuple(size for size, _ in frames) == EXPECTED_ICO_SIZES
    for size, payload in frames:
        with Image.open(io.BytesIO(payload)) as frame:
            assert frame.size == (size, size)
            assert frame.mode == "RGBA"
    assert frames[-1][1] == PNG_PATH.read_bytes()

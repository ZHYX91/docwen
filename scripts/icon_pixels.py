"""Exact alpha conversion shared by application and Store icon generation."""

from __future__ import annotations


def unpremultiply_rgba(pixels: bytes) -> bytes:
    """Use one integer rounding rule for straight-alpha icon pixels.

    Qt's CPU-specific image conversion paths can round halfway values
    differently. Keep rasterization premultiplied and convert its byte-ordered
    RGBA channels with exact arithmetic so build hosts produce identical PNGs.
    """
    result = bytearray(pixels)
    for offset in range(0, len(result), 4):
        alpha = result[offset + 3]
        if 0 < alpha < 255:
            for channel in range(offset, offset + 3):
                result[channel] = min(255, (result[channel] * 255 + alpha // 2) // alpha)
    return bytes(result)

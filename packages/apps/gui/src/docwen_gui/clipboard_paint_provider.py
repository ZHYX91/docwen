"""Recognize a bounded Paint bitmap source without reading its source path."""

from __future__ import annotations

import struct

from docwen_core.models.clipboard_document import MAX_CLIPBOARD_IMAGE_PIXELS_PER_RESOURCE, MAX_CLIPBOARD_IMAGE_SIDE
from docwen_gui.clipboard_office_provider import ClipboardOfficeProviderError, _read_cfb_stream

_PAINT_CLSID = bytes.fromhex("0a00030000000000c000000000000046")
_SOURCE_STREAMS = ("\x01Ole10Native", "Package", "package", "Workbook", "Book", "WordDocument")


def paint_bitmap_dimensions(payload: bytes) -> tuple[int, int] | None:
    """Recognize Paint's unique native BI_RGB bitmap, excluding Office sources.

    The bitmap only proves the standalone image kind and physical dimensions.
    Its white-composited pixels are not used instead of Qt's alpha-preserving
    image data. Unknown, ambiguous and malformed OLE sources remain rejected.
    """
    try:
        name, native = _read_cfb_stream(payload, _SOURCE_STREAMS, expected_root_clsid=_PAINT_CLSID)
    except ClipboardOfficeProviderError:
        return None
    if name != "\x01Ole10Native" or len(native) < 58 or struct.unpack_from("<I", native)[0] != len(native) - 4:
        return None
    bitmap = native[4:]
    if bitmap[:2] != b"BM":
        return None
    bitmap_size = struct.unpack_from("<I", bitmap, 2)[0]
    if bitmap_size < 54 or len(bitmap) != (bitmap_size + 15) // 16 * 16 or any(bitmap[bitmap_size:]):
        return None
    if bitmap[6:10] != b"\0\0\0\0" or struct.unpack_from("<I", bitmap, 10)[0] != 54:
        return None
    header_size, width, height, planes, bits, compression, image_size = struct.unpack_from("<IiiHHII", bitmap, 14)
    if (
        header_size != 40
        or not 1 <= width <= MAX_CLIPBOARD_IMAGE_SIDE
        or not 1 <= abs(height) <= MAX_CLIPBOARD_IMAGE_SIDE
        or width * abs(height) > MAX_CLIPBOARD_IMAGE_PIXELS_PER_RESOURCE
        or planes != 1
        or bits != 32
        or compression != 0
        or image_size not in {0, width * abs(height) * 4}
        or bitmap_size != 54 + width * abs(height) * 4
    ):
        return None
    return width, abs(height)


__all__ = ["paint_bitmap_dimensions"]

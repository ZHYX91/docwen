"""Bounded image freezing for clipboard inputs.

Only PNG is admitted into structured clipboard resources. Standalone Qt
images are detached on the owner thread as straight RGBA8 and encoded once;
background work consumes only the frozen bytes.
"""

from __future__ import annotations

import hashlib
import io
import struct
from dataclasses import dataclass

from docwen_core.models.clipboard_document import (
    MAX_CLIPBOARD_IMAGE_PIXELS_PER_RESOURCE,
    MAX_CLIPBOARD_IMAGE_SIDE,
)

_PNG_SIGNATURE = b"\x89PNG\r\n\x1a\n"


class ClipboardImageBytesError(ValueError):
    """Stable bounded failure for clipboard image bytes."""

    def __init__(self, code: str, message: str) -> None:
        self.code = code
        super().__init__(message)


@dataclass(frozen=True, slots=True)
class FrozenPng:
    payload: bytes
    width: int
    height: int
    rgba_sha256: str
    payload_sha256: str
    device_pixel_ratio: float = 1.0


def _png_dimensions(payload: bytes) -> tuple[int, int]:
    if len(payload) < 33 or not payload.startswith(_PNG_SIGNATURE):
        raise ClipboardImageBytesError("clipboard.image_png_invalid", "Clipboard image is not a valid PNG.")
    length = struct.unpack_from(">I", payload, 8)[0]
    if length != 13 or payload[12:16] != b"IHDR":
        raise ClipboardImageBytesError("clipboard.image_png_invalid", "Clipboard PNG has an invalid header.")
    width, height = struct.unpack_from(">II", payload, 16)
    if (
        width < 1
        or height < 1
        or width > MAX_CLIPBOARD_IMAGE_SIDE
        or height > MAX_CLIPBOARD_IMAGE_SIDE
        or width * height > MAX_CLIPBOARD_IMAGE_PIXELS_PER_RESOURCE
    ):
        raise ClipboardImageBytesError(
            "clipboard.image_budget_exceeded",
            "Clipboard image exceeds the supported pixel budget.",
        )
    return width, height


def inspect_png_bytes(payload: bytes, *, device_pixel_ratio: float = 1.0) -> FrozenPng:
    """Decode one bounded PNG and hash its straight RGBA pixels."""

    if not isinstance(payload, bytes) or not payload:
        raise ClipboardImageBytesError("clipboard.image_png_invalid", "Clipboard PNG is empty or invalid.")
    width, height = _png_dimensions(payload)
    try:
        from PIL import Image

        with Image.open(io.BytesIO(payload)) as image:
            if image.format != "PNG" or image.size != (width, height):
                raise ClipboardImageBytesError(
                    "clipboard.image_png_invalid",
                    "Clipboard image is not a valid PNG.",
                )
            image.load()
            rgba = image.convert("RGBA").tobytes()
    except ClipboardImageBytesError:
        raise
    except Exception as exc:
        raise ClipboardImageBytesError(
            "clipboard.image_png_invalid",
            "Clipboard image could not be decoded safely.",
        ) from exc
    if len(rgba) != width * height * 4:
        raise ClipboardImageBytesError(
            "clipboard.image_png_invalid",
            "Clipboard image decoded to an invalid pixel buffer.",
        )
    return FrozenPng(
        payload=payload,
        width=width,
        height=height,
        rgba_sha256=hashlib.sha256(rgba).hexdigest(),
        payload_sha256=hashlib.sha256(payload).hexdigest(),
        device_pixel_ratio=float(device_pixel_ratio),
    )


def freeze_qimage(image: object) -> FrozenPng:
    """Detach a Qt image as physical-pixel straight RGBA8 and encode PNG once."""

    from PySide6.QtCore import QBuffer, QByteArray, QIODevice
    from PySide6.QtGui import QImage, QPixmap

    if isinstance(image, QPixmap):
        source = image.toImage()
    elif isinstance(image, QImage):
        source = image
    else:
        raise ClipboardImageBytesError(
            "clipboard.image_qt_invalid",
            "Clipboard image data is unavailable.",
        )
    if source.isNull():
        raise ClipboardImageBytesError("clipboard.image_qt_invalid", "Clipboard image data is empty.")

    width = int(source.width())
    height = int(source.height())
    if (
        width < 1
        or height < 1
        or width > MAX_CLIPBOARD_IMAGE_SIDE
        or height > MAX_CLIPBOARD_IMAGE_SIDE
        or width * height > MAX_CLIPBOARD_IMAGE_PIXELS_PER_RESOURCE
    ):
        raise ClipboardImageBytesError(
            "clipboard.image_budget_exceeded",
            "Clipboard image exceeds the supported pixel budget.",
        )

    dpr = float(source.devicePixelRatio())
    detached = source.convertToFormat(QImage.Format.Format_RGBA8888).copy()
    detached.setDevicePixelRatio(1.0)
    if detached.isNull() or detached.width() != width or detached.height() != height:
        raise ClipboardImageBytesError(
            "clipboard.image_qt_invalid",
            "Clipboard image could not be detached safely.",
        )

    raw = bytes(detached.constBits())
    row_bytes = width * 4
    stride = int(detached.bytesPerLine())
    if stride < row_bytes or len(raw) < stride * height:
        raise ClipboardImageBytesError(
            "clipboard.image_qt_invalid",
            "Clipboard image has an invalid pixel layout.",
        )
    rgba = b"".join(raw[row * stride : row * stride + row_bytes] for row in range(height))
    rgba_sha = hashlib.sha256(rgba).hexdigest()

    encoded = QByteArray()
    buffer = QBuffer(encoded)
    if not buffer.open(QIODevice.OpenModeFlag.WriteOnly):
        raise ClipboardImageBytesError(
            "clipboard.image_encode_failed",
            "Clipboard image could not be encoded.",
        )
    try:
        if not detached.save(buffer, "PNG"):
            raise ClipboardImageBytesError(
                "clipboard.image_encode_failed",
                "Clipboard image could not be encoded.",
            )
    finally:
        buffer.close()

    frozen = inspect_png_bytes(bytes(encoded), device_pixel_ratio=dpr)
    if frozen.width != width or frozen.height != height or frozen.rgba_sha256 != rgba_sha:
        raise ClipboardImageBytesError(
            "clipboard.image_encode_failed",
            "Clipboard image encoding changed its pixel content.",
        )
    return frozen


__all__ = [
    "ClipboardImageBytesError",
    "FrozenPng",
    "freeze_qimage",
    "inspect_png_bytes",
]

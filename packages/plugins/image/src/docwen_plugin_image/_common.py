"""Shared helpers for the image plugin."""

from __future__ import annotations

import io
import os
import uuid
from pathlib import Path
from typing import TYPE_CHECKING, Any

from PIL import Image, ImageOps, ImageSequence

from docwen_core.paths import input_stem

if TYPE_CHECKING:
    from docwen_core.protocols.execution_context import ConverterContext

IMAGE_MEDIA_TYPES = {
    "jpg": "image/jpeg",
    "jpeg": "image/jpeg",
    "png": "image/png",
    "gif": "image/gif",
    "bmp": "image/bmp",
    "tif": "image/tiff",
    "tiff": "image/tiff",
    "webp": "image/webp",
}

PIL_FORMATS = {
    "jpg": "JPEG",
    "jpeg": "JPEG",
    "png": "PNG",
    "gif": "GIF",
    "bmp": "BMP",
    "tif": "TIFF",
    "tiff": "TIFF",
    "webp": "WEBP",
}

SUPPORTED_IMAGE_FORMATS = ("jpg", "jpeg", "png", "gif", "bmp", "tif", "tiff", "webp")
HEIC_INPUT_FORMATS = ("heic", "heif")
FLAT_EXPORT_EXIF_TAGS = (271, 272, 305, 306)


def normalize_format(value: str) -> str:
    value = (value or "").lower().lstrip(".")
    if value == "jpeg":
        return "jpg"
    if value == "tiff":
        return "tif"
    return value


def source_format_from_context(context: ConverterContext) -> str:
    """Return the admitted concrete format for the active workspace input.

    The application admission boundary owns content inspection.  Image
    converters therefore consume ``FileRef.format`` and never infer parser
    selection from a path suffix.
    """
    input_refs = context.request.input_refs
    if not input_refs:
        raise ValueError("Image conversion requires an admitted input reference.")

    workspace_path = Path(context.workspace.input_path).resolve(strict=False)
    matching_ref = next(
        (ref for ref in input_refs if Path(ref.path).resolve(strict=False) == workspace_path),
        input_refs[0] if len(input_refs) == 1 else None,
    )
    if matching_ref is None:
        raise ValueError("The active image input has no matching admitted input reference.")

    source_format = normalize_format(matching_ref.format)
    if source_format in {"", "image", "unknown"}:
        raise ValueError("The admitted image input must retain its concrete detected format.")
    return source_format


def is_heic_format(source_format: str) -> bool:
    return normalize_format(source_format) in HEIC_INPUT_FORMATS


def preconvert_heic_to_png(input_path: str, staging_dir: str) -> str:
    """Convert HEIC/HEIF input to a PNG intermediate in the plugin workspace."""
    try:
        from pillow_heif import register_heif_opener
    except ImportError as exc:
        raise RuntimeError("HEIC/HEIF conversion requires the optional pillow-heif dependency.") from exc

    register_heif_opener()
    output_path = Path(staging_dir) / f"{input_stem(input_path)}.png"
    counter = 1
    while output_path.exists():
        output_path = Path(staging_dir) / f"{input_stem(input_path)}_{counter}.png"
        counter += 1
    output_path.parent.mkdir(parents=True, exist_ok=True)

    try:
        with Image.open(input_path) as img:
            converted = convert_mode_for_format(img, "png")
            try:
                converted.save(output_path, **save_params("png"))
            finally:
                converted.close()
    except Exception as exc:
        raise RuntimeError(f"HEIC/HEIF preprocessing failed: {exc}") from exc
    return str(output_path)


def media_type_for(format_name: str) -> str:
    return IMAGE_MEDIA_TYPES.get(format_name.lower(), "application/octet-stream")


def pil_format_for(format_name: str) -> str:
    normalized = normalize_format(format_name)
    return PIL_FORMATS.get(normalized, normalized.upper())


def new_artifact_id() -> str:
    return str(uuid.uuid4())


def file_size(path: str) -> int:
    return os.path.getsize(path) if os.path.isfile(path) else 0


def has_alpha(img: Image.Image) -> bool:
    return img.mode in ("RGBA", "LA", "PA") or (img.mode == "P" and img.info.get("transparency") is not None)


def paste_on_white(img: Image.Image) -> Image.Image:
    background = Image.new("RGB", img.size, (255, 255, 255))
    working = img.convert("RGBA") if img.mode in ("P", "PA") else img
    if working.mode in ("RGBA", "LA"):
        background.paste(working, mask=working.split()[-1])
    else:
        background.paste(working)
    return background


def convert_mode_for_format(img: Image.Image, target_format: str) -> Image.Image:
    target = normalize_format(target_format)
    if target in ("jpg", "bmp"):
        if has_alpha(img):
            return paste_on_white(img)
        supported_modes = ("RGB", "L") if target == "jpg" else ("RGB",)
        if img.mode not in supported_modes:
            return img.convert("RGB")
        return img.copy()
    if target == "gif":
        if img.mode not in ("P", "L"):
            return img.convert("P")
        return img.copy()
    if target == "png":
        if img.mode in ("P", "LA", "PA", "La", "RGBa"):
            return img.convert("RGBA")
        if img.mode not in ("1", "L", "I", "I;16", "RGB", "RGBA"):
            return img.convert("RGB")
        return img.copy()
    if target == "webp":
        if has_alpha(img) or img.mode in ("LA", "La", "RGBa"):
            return img.convert("RGBA")
        if img.mode not in ("RGB", "L"):
            return img.convert("RGB")
        return img.copy()
    if target == "tif":
        if img.mode in ("P", "LA"):
            return img.convert("RGB")
        return img.copy()
    return img.copy()


def save_params(target_format: str, quality: int = 95) -> dict[str, Any]:
    target = normalize_format(target_format)
    if target == "jpg":
        return {"format": "JPEG", "quality": quality, "optimize": True}
    if target == "webp":
        return {"format": "WEBP", "quality": quality, "method": 6}
    if target == "png":
        return {"format": "PNG", "compress_level": 9, "optimize": True}
    if target == "gif":
        return {"format": "GIF", "optimize": True}
    if target == "tif":
        return {"format": "TIFF", "compression": "tiff_lzw"}
    if target == "bmp":
        return {"format": "BMP"}
    return {"format": pil_format_for(target)}


def normalize_size_unit(unit: str) -> str:
    normalized = (unit or "KB").upper()
    if normalized not in ("KB", "MB"):
        raise ValueError(f"Invalid size unit: {unit}")
    return normalized


def target_size_bytes(size_limit: Any, unit: str) -> int:
    try:
        size_int = int(size_limit)
    except (TypeError, ValueError):
        raise ValueError("size_limit must be an integer") from None
    if size_int <= 0:
        raise ValueError("size_limit must be a positive integer")
    normalized = normalize_size_unit(unit)
    return size_int * 1024 if normalized == "KB" else size_int * 1024 * 1024


def prepare_flat_export(img: Image.Image) -> tuple[Image.Image, dict[str, Any]]:
    """Normalize display orientation and retain the supported flat-export metadata."""
    source_exif = img.getexif()
    output_exif = Image.Exif()
    for tag in FLAT_EXPORT_EXIF_TAGS:
        value = source_exif.get(tag)
        if value is not None:
            output_exif[tag] = value
    if source_exif.get(274) is not None:
        output_exif[274] = 1

    save_metadata: dict[str, Any] = {}
    if len(output_exif):
        save_metadata["exif"] = output_exif
    icc_profile = img.info.get("icc_profile")
    if isinstance(icc_profile, bytes):
        save_metadata["icc_profile"] = icc_profile
    return ImageOps.exif_transpose(img), save_metadata


def compress_to_size(
    img: Image.Image,
    output_path: str,
    target_format: str,
    limit_bytes: int,
    save_metadata: dict[str, Any] | None = None,
) -> bool:
    target = normalize_format(target_format)
    quality_min = 15
    quality_max = 95
    best_buffer: io.BytesIO | None = None
    metadata = save_metadata or {}

    while quality_min <= quality_max:
        quality = (quality_min + quality_max) // 2
        buffer = io.BytesIO()
        img.save(buffer, **save_params(target, quality=quality), **metadata)
        current_size = buffer.tell()
        if current_size <= limit_bytes:
            best_buffer = buffer
            quality_min = quality + 1
        else:
            quality_max = quality - 1

    if best_buffer is None:
        buffer = io.BytesIO()
        img.save(buffer, **save_params(target, quality=15), **metadata)
        best_buffer = buffer

    Path(output_path).write_bytes(best_buffer.getvalue())
    return file_size(output_path) <= limit_bytes


def save_image_with_options(
    img: Image.Image,
    output_path: str,
    target_format: str,
    options: dict[str, Any],
    *,
    save_metadata: dict[str, Any] | None = None,
) -> None:
    target = normalize_format(target_format)
    converted = convert_mode_for_format(img, target)
    metadata = save_metadata or {}
    try:
        compress_mode = options.get("compress_mode") or "lossless"
        if compress_mode == "limit_size":
            limit_bytes = target_size_bytes(options.get("size_limit"), options.get("size_unit", "KB"))
            if target in ("jpg", "webp"):
                ok = compress_to_size(converted, output_path, target, limit_bytes, metadata)
                if not ok:
                    raise ValueError(
                        f"Unable to compress image to target size: target={limit_bytes} bytes, actual={file_size(output_path)} bytes"
                    )
                return
            converted.save(output_path, **save_params(target), **metadata)
            if file_size(output_path) > limit_bytes:
                raise ValueError(
                    f"{target.upper()} cannot be reliably compressed to target size: target={limit_bytes} bytes, actual={file_size(output_path)} bytes"
                )
            return
        if compress_mode != "lossless":
            raise ValueError(f"Unknown compress_mode: {compress_mode}")
        converted.save(output_path, **save_params(target), **metadata)
    finally:
        converted.close()


def iter_frames(input_path: str) -> list[Image.Image]:
    frames: list[Image.Image] = []
    with Image.open(input_path) as img:
        for frame in ImageSequence.Iterator(img):
            copy = frame.copy()
            copy.load()
            frames.append(copy)
    return frames

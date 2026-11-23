"""Preserve all pages or animation frames when exporting a sequence."""

from __future__ import annotations

from contextlib import ExitStack
from typing import TYPE_CHECKING, Any

from PIL import Image

from docwen_plugin_image._common import (
    convert_mode_for_format,
    prepare_flat_export,
    rgba_for_export,
    save_image_with_options,
    target_size_bytes,
)

if TYPE_CHECKING:
    from collections.abc import Callable
    from typing import BinaryIO


def _loop_metadata(source: Image.Image, target: str) -> dict[str, int]:
    """Translate GIF repeats after the first play into WebP/APNG total plays."""
    raw_loop = source.info.get("loop")
    plays = 1 if raw_loop is None else int(raw_loop)
    if source.format == "GIF" and raw_loop is not None and plays > 0:
        plays += 1
    if target == "gif":
        if plays == 1:
            return {}  # No Netscape extension means one play, not infinite looping.
        loop = plays - 1 if plays else 0
    else:
        loop = plays
    maximum = 0xFFFFFFFF if target == "png" else 0xFFFF
    if not 0 <= loop <= maximum:
        raise ValueError(f"Animation loop count cannot be represented in {target.upper()}")
    return {"loop": loop}


def save_sequence(
    source: Image.Image,
    output_path: str,
    target: str,
    options: dict[str, Any],
    *,
    cancel_check: Callable[[], None],
    source_stream: BinaryIO,
) -> int:
    """Save composited frames, preserving playback timing and loop semantics."""
    if source.format == "PNG" and source.mode.startswith("I;16") and target == "png":
        # Pillow computes APNG frame differences via an 8-bit projection. Two
        # distinct wide samples can collapse to one frame even if mode is retained.
        count = getattr(source, "n_frames", 1)
        for index in range(count):
            cancel_check()
            source.seek(index)
            source.load()
        compress_mode = options.get("compress_mode") or "lossless"
        if compress_mode == "limit_size":
            limit = target_size_bytes(options.get("size_limit"), options.get("size_unit", "KB"))
            if source_stream.seek(0, 2) > limit:
                raise ValueError("The original 16-bit APNG exceeds the requested size limit.")
        elif compress_mode != "lossless":
            raise ValueError(f"Unknown compress_mode: {compress_mode}")
        source_stream.seek(0)
        with open(output_path, "wb") as destination:
            while chunk := source_stream.read(1024 * 1024):
                cancel_check()
                destination.write(chunk)
        return count
    with ExitStack() as stack:
        frames: list[Image.Image] = []
        durations: list[int] = []
        metadata: dict[str, Any] = {}
        poster = bool(source.info.get("default_image", False))
        start = 1 if poster and target != "png" else 0
        for index in range(start, getattr(source, "n_frames", 1)):
            cancel_check()
            source.seek(index)
            prepared, frame_metadata = prepare_flat_export(source)
            with prepared:
                # Seeking composites the input's disposal and blend operations.
                # Normalize every frame, including those passed as append_images.
                frame = (
                    convert_mode_for_format(prepared, target) if target in ("gif", "tif") else rgba_for_export(prepared)
                )
            stack.callback(frame.close)
            # Encoders may fall back to image.info when an explicit save option
            # is absent (notably a GIF that should play only once).
            frame.info.pop("loop", None)
            frames.append(frame)
            if not poster or index != 0:
                durations.append(int(source.info.get("duration", 0)))
            if not metadata:
                metadata.update(frame_metadata)
        metadata.update(save_all=True, append_images=frames[1:])
        if target in ("gif", "webp", "png"):
            metadata["duration"] = durations
            metadata.pop("loop", None)
            metadata.update(_loop_metadata(source, target))
            if target == "png":
                metadata.update(disposal=0, blend=0, default_image=poster)
            elif target == "gif":
                metadata["disposal"] = 2
        save_image_with_options(frames[0], output_path, target, options, save_metadata=metadata)
        return len(frames)

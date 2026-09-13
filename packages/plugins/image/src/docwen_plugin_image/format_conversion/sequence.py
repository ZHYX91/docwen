"""Preserve all pages or animation frames when exporting a sequence."""

from __future__ import annotations

from contextlib import ExitStack
from typing import TYPE_CHECKING, Any

from PIL import Image

from docwen_plugin_image._common import prepare_flat_export, save_image_with_options

if TYPE_CHECKING:
    from collections.abc import Callable


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
) -> int:
    """Save composited frames, preserving playback timing and loop semantics."""
    with ExitStack() as stack:
        frames: list[Image.Image] = []
        durations: list[int] = []
        metadata: dict[str, Any] = {}
        poster = bool(source.info.get("default_image", False))
        start = 1 if poster and target != "png" else 0
        for index in range(start, getattr(source, "n_frames", 1)):
            cancel_check()
            source.seek(index)
            frame, frame_metadata = prepare_flat_export(source)
            stack.callback(frame.close)
            if target != "tif":
                # Seeking through Pillow's decoder composites disposal/blend
                # operations. Export full frames with replacement semantics.
                rgba = frame.convert("RGBA")
                stack.callback(rgba.close)
                frame = rgba
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

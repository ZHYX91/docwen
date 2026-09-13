"""Preserve all pages or animation frames when exporting a sequence."""

from __future__ import annotations

from contextlib import ExitStack
from typing import TYPE_CHECKING, Any

from PIL import Image

from docwen_plugin_image._common import prepare_flat_export, save_image_with_options

if TYPE_CHECKING:
    from collections.abc import Callable


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
            frames.append(frame)
            if not poster or index != 0:
                durations.append(int(source.info.get("duration", 0)))
            if not metadata:
                metadata.update(frame_metadata)
        metadata.update(save_all=True, append_images=frames[1:])
        if target in ("gif", "webp", "png"):
            metadata["duration"] = durations
            if "loop" in source.info:
                metadata["loop"] = source.info["loop"]
            elif target in ("webp", "png"):
                metadata["loop"] = 1
            if target == "png":
                metadata.update(disposal=0, blend=0, default_image=poster)
            elif target == "gif":
                metadata["disposal"] = 2
        save_image_with_options(frames[0], output_path, target, options, save_metadata=metadata)
        return len(frames)

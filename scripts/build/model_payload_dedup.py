"""Share byte-identical bundled RapidOCR defaults with DocWen's model resources.

Both locations remain addressable: third-party RapidOCR defaults use its package
directory, while DocWen deliberately selects models/rapidocr by language.
Only POSIX onedir payloads use relative file links. Windows ZIP/MSIX layouts
remain unchanged, and any absent or differing file is retained verbatim.
"""

from __future__ import annotations

import hashlib
import os
import sys
from pathlib import Path

_DEFAULT_MODEL_FILES = (
    "ch_PP-OCRv4_det_infer.onnx",
    "ch_PP-OCRv4_rec_infer.onnx",
    "ch_ppocr_mobile_v2.0_cls_infer.onnx",
)


def _digest(path: Path) -> bytes:
    with path.open("rb") as stream:
        return hashlib.file_digest(stream, "sha256").digest()


def compact_rapidocr_model_copies(payload: Path, *, platform: str | None = None) -> dict[str, int]:
    """Replace identical embedded model *files* with internal relative links.

    The external and vendor model paths are both retained. Never link
    directories, never modify model bytes, and leave Windows binaries alone.
    A mismatched vendor model is a valid independent resource, not an error.
    """

    if (sys.platform if platform is None else platform) not in {"linux", "darwin"}:
        return {"linked": 0, "skipped": len(_DEFAULT_MODEL_FILES), "saved_bytes": 0}

    vendor_models = payload / "_internal" / "rapidocr_onnxruntime" / "models"
    public_models = payload / "models" / "rapidocr"
    linked = 0
    skipped = 0
    saved_bytes = 0

    for name in _DEFAULT_MODEL_FILES:
        bundled = vendor_models / name
        shared = public_models / name

        if bundled.is_symlink():
            # Idempotent if an earlier pass already linked to this exact file.
            if bundled.is_file() and bundled.resolve() == shared.resolve() and shared.is_file():
                continue
            # Do not override a vendor-provided link with unknown semantics.
            skipped += 1
            continue

        if (
            not bundled.is_file()
            or not shared.is_file()
            or shared.is_symlink()
            or bundled.stat().st_size != shared.stat().st_size
            or _digest(bundled) != _digest(shared)
        ):
            skipped += 1
            continue

        relative_target = os.path.relpath(shared, start=bundled.parent)
        temporary = bundled.with_name(f".{name}.docwen-dedup-link")
        if temporary.exists() or temporary.is_symlink():
            raise RuntimeError(f"rapidocr_dedup_staging_exists:{name}")
        try:
            temporary.symlink_to(relative_target)
            if not temporary.is_file() or temporary.resolve(strict=True) != shared.resolve(strict=True):
                raise RuntimeError(f"rapidocr_dedup_target_invalid:{name}")
            size = bundled.stat().st_size
            os.replace(temporary, bundled)
            linked += 1
            saved_bytes += size
        finally:
            temporary.unlink(missing_ok=True)

    return {"linked": linked, "skipped": skipped, "saved_bytes": saved_bytes}

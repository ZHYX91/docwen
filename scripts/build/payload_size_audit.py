"""Build-only physical byte inventory for frozen DocWen payloads.

Count regular files rather than following links, so POSIX model sharing is not
double-counted. Report relative paths only; the release CI runner's workspace
path is deliberately never emitted.
"""

from __future__ import annotations

from collections import defaultdict
from pathlib import Path
from typing import TypedDict


class PayloadSizeReport(TypedDict):
    physical_bytes: int
    regular_files: int
    links: int
    ocr_model_bytes: int
    other_model_bytes: int
    non_ocr_bytes: int
    non_model_bytes: int
    top_groups: list[tuple[str, int]]
    top_non_ocr_groups: list[tuple[str, int]]
    largest_files: list[tuple[str, int]]
    largest_non_ocr_files: list[tuple[str, int]]


def _model_kind(relative: Path) -> str | None:
    """Categorize ONNX bytes separately from non-model runtime dependencies."""

    name = relative.as_posix()
    if relative.suffix.lower() != ".onnx":
        return None
    if name.startswith(("models/rapidocr/", "_internal/rapidocr_onnxruntime/models/")):
        return "ocr"
    if name.startswith(("models/rapidtable/", "_internal/pymupdf/layout/resources/onnx/")):
        return "other"
    return None


def _size_group(relative: Path) -> str:
    parts = relative.parts
    if parts[0] != "_internal":
        return parts[0]
    if len(parts) == 2:
        return "_internal/(root files)"
    return "_internal/" + parts[1]


def audit_payload_size(root: Path, *, max_items: int = 10) -> PayloadSizeReport:
    """Summarize on-disk bytes without touching any payload contents."""

    if not root.is_dir() or root.is_symlink():
        raise ValueError("size_audit_requires_plain_directory")

    groups: dict[str, int] = defaultdict(int)
    non_ocr_groups: dict[str, int] = defaultdict(int)
    files: list[tuple[int, str]] = []
    non_ocr_files: list[tuple[int, str]] = []
    ocr_model_bytes = 0
    other_model_bytes = 0
    link_count = 0

    for path in root.rglob("*"):
        if path.is_symlink():
            link_count += 1
            continue
        if not path.is_file():
            continue
        relative = path.relative_to(root)
        size = path.lstat().st_size
        group = _size_group(relative)
        groups[group] += size
        files.append((size, relative.as_posix()))
        kind = _model_kind(relative)
        if kind == "ocr":
            ocr_model_bytes += size
        else:
            non_ocr_groups[group] += size
            non_ocr_files.append((size, relative.as_posix()))
            if kind == "other":
                other_model_bytes += size

    largest = sorted(files, key=lambda item: (-item[0], item[1]))[:max_items]
    grouped = sorted(groups.items(), key=lambda item: (-item[1], item[0]))
    non_ocr_grouped = sorted(non_ocr_groups.items(), key=lambda item: (-item[1], item[0]))
    largest_non_ocr = sorted(non_ocr_files, key=lambda item: (-item[0], item[1]))[:max_items]
    physical_bytes = sum(size for size, _ in files)
    return {
        "physical_bytes": physical_bytes,
        "regular_files": len(files),
        "links": link_count,
        "ocr_model_bytes": ocr_model_bytes,
        "other_model_bytes": other_model_bytes,
        "non_ocr_bytes": physical_bytes - ocr_model_bytes,
        "non_model_bytes": physical_bytes - ocr_model_bytes - other_model_bytes,
        "top_groups": grouped[:max_items],
        "top_non_ocr_groups": non_ocr_grouped[:max_items],
        "largest_files": [(path, size) for size, path in largest],
        "largest_non_ocr_files": [(path, size) for size, path in largest_non_ocr],
    }

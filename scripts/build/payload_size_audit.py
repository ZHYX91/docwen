"""Build-only physical byte inventory for frozen DocWen payloads.

Count regular files rather than following links, so POSIX model sharing is not
double-counted. Report relative paths only; the release CI runner's workspace
path is deliberately never emitted.
"""

from __future__ import annotations

from collections import defaultdict
from pathlib import Path


def _size_group(relative: Path) -> str:
    parts = relative.parts
    if parts[0] != "_internal":
        return parts[0]
    if len(parts) == 2:
        return "_internal/(root files)"
    return "_internal/" + parts[1]


def audit_payload_size(root: Path, *, max_items: int = 10) -> dict[str, object]:
    """Summarize on-disk bytes without touching any payload contents."""

    if not root.is_dir() or root.is_symlink():
        raise ValueError("size_audit_requires_plain_directory")

    groups: dict[str, int] = defaultdict(int)
    files: list[tuple[int, str]] = []
    link_count = 0

    for path in root.rglob("*"):
        if path.is_symlink():
            link_count += 1
            continue
        if not path.is_file():
            continue
        relative = path.relative_to(root)
        size = path.lstat().st_size
        groups[_size_group(relative)] += size
        files.append((size, relative.as_posix()))

    largest = sorted(files, key=lambda item: (-item[0], item[1]))[:max_items]
    grouped = sorted(groups.items(), key=lambda item: (-item[1], item[0]))
    return {
        "physical_bytes": sum(size for size, _ in files),
        "regular_files": len(files),
        "links": link_count,
        "top_groups": grouped[:max_items],
        "largest_files": [(path, size) for size, path in largest],
    }

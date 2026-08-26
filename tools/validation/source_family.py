"""Read one source file or an explicitly named source-file family."""

from __future__ import annotations

from pathlib import Path


def read_source_text(path: Path) -> str:
    """Read an exact path, or concatenate a wildcard family in stable order."""
    if not any(character in path.name for character in "*?["):
        return path.read_text(encoding="utf-8")

    members = sorted(path.parent.glob(path.name))
    if path.name.startswith("test_") and path.name.endswith("_*.py"):
        support_name = f"_{path.name.removeprefix('test_').removesuffix('_*.py')}_support.py"
        support = path.parent / support_name
        if support.is_file():
            members.insert(0, support)
    if not members:
        raise FileNotFoundError(path)
    return "\n".join(member.read_text(encoding="utf-8") for member in members)

"""Resource root discovery for runtime registries.

The runtime needs a small, dependency-free way to find bundled resources in
source checkouts and packaged layouts.  This module intentionally returns paths;
it does not load template, locale, or config semantics.
"""

from __future__ import annotations

import sys
from pathlib import Path


class ResourceRegistry:
    """Resolve known resource directories from a distribution root."""

    def __init__(self, root: Path | str | None = None) -> None:
        self.root = Path(root) if root is not None else find_project_root()

    @classmethod
    def default(cls) -> ResourceRegistry:
        return cls(find_project_root())

    def templates_dir(self) -> Path:
        return self.root / "templates"

    def configs_dir(self) -> Path:
        return self.root / "configs"

    def assets_dir(self) -> Path:
        return self.root / "assets"

    def locales_dir(self) -> Path:
        candidates = [
            self.root / "i18n" / "locales",
            self.root / "_internal" / "docwen" / "i18n" / "locales",
            self.root / "docwen" / "i18n" / "locales",
            self.root / "src" / "docwen" / "i18n" / "locales",
        ]
        for candidate in candidates:
            if candidate.exists():
                return candidate
        return candidates[0]


def find_project_root() -> Path:
    """Find the resource root for source and PyInstaller layouts."""

    meipass = getattr(sys, "_MEIPASS", None)
    if isinstance(meipass, str) and meipass:
        return _find_root_from_meipass(Path(meipass))

    cur = Path(__file__).resolve()
    return _find_root_from_module_path(cur)


def _find_root_from_module_path(module_file: Path) -> Path:
    cur = module_file.resolve()
    for parent in cur.parents:
        if _path_exists(parent / "templates") and _path_exists(parent / "configs"):
            return parent
        if _path_exists(parent / "pyproject.toml") and _path_exists(parent / "packages"):
            return parent

    package_root = cur.parents[1]
    for candidate in (
        package_root,
        package_root / "resources",
        package_root.parent,
    ):
        if _path_exists(candidate / "templates") or _path_exists(candidate / "configs"):
            return candidate

    return Path.cwd()


def _path_exists(path: Path) -> bool:
    """Treat inaccessible discovery candidates as absent."""

    try:
        return path.exists()
    except OSError:
        return False


def _find_root_from_meipass(meipass: Path) -> Path:
    """Resolve a packaged root from PyInstaller's internal extraction dir."""

    internal_root = meipass.resolve()
    deploy_root = internal_root.parent
    if (
        internal_root.name == "_internal"
        and (deploy_root / "templates").exists()
        and (deploy_root / "configs").exists()
    ):
        return deploy_root
    return internal_root

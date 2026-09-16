"""Resolve one coherent writable DocWen profile for every entry point.

Packaged archive builds are portable by default: their durable user data lives
beside the executable under ``data``.  Store/MSIX builds cannot write beside
the package and therefore use the platform user-data directory.  Source/dev
runs keep the historical platform defaults unless an explicit isolation hook is
provided.

The existing ``DOCWEN_*_DIR`` variables remain internal host/test hooks.  This
module only fills missing variables; explicit callers keep authority over an
injected profile.
"""

from __future__ import annotations

import ctypes
import hashlib
import os
import sys
from dataclasses import dataclass
from pathlib import Path

from platformdirs import user_data_dir

_APPMODEL_ERROR_NO_PACKAGE = 15700
_ERROR_INSUFFICIENT_BUFFER = 122


@dataclass(frozen=True, slots=True)
class ProfilePaths:
    root: Path
    config_dir: Path
    data_dir: Path
    log_root: Path
    portable: bool


def _windows_has_package_identity() -> bool:
    if sys.platform != "win32":
        return False
    try:
        get_name = ctypes.windll.kernel32.GetCurrentPackageFullName
    except (AttributeError, OSError):
        return False
    length = ctypes.c_uint32(0)
    result = int(get_name(ctypes.byref(length), None))
    if result == _ERROR_INSUFFICIENT_BUFFER:
        return True
    # An indeterminate package-identity probe must not authorize writes beside
    # a potentially protected installation.
    return result != _APPMODEL_ERROR_NO_PACKAGE


def _packaged_archive_root() -> Path | None:
    if not bool(getattr(sys, "frozen", False)):
        return None
    if sys.platform == "win32" and _windows_has_package_identity():
        return None
    executable = Path(sys.executable).resolve(strict=False)
    return executable.parent / "data"


def default_profile_paths() -> ProfilePaths | None:
    """Return automatic packaged paths, or ``None`` for source/dev runs."""

    portable_root = _packaged_archive_root()
    if portable_root is not None:
        root = portable_root.absolute()
        return ProfilePaths(root, root / "configs", root, root, True)
    if bool(getattr(sys, "frozen", False)):
        root = Path(user_data_dir("docwen", appauthor=False)).absolute()
        return ProfilePaths(root, root / "configs", root, root, False)
    return None


def configure_process_profile() -> ProfilePaths | None:
    """Bind Config, template data, and logs to one packaged profile.

    The logging subsystem interprets ``DOCWEN_LOG_DIR`` as a profile root and
    appends ``logs`` itself, while the configuration hook expects the exact
    writable ``configs`` directory.
    """

    automatic = default_profile_paths()
    if automatic is None:
        return None
    os.environ.setdefault("DOCWEN_CONFIG_DIR", str(automatic.config_dir))
    os.environ.setdefault("DOCWEN_DATA_DIR", str(automatic.data_dir))
    os.environ.setdefault("DOCWEN_LOG_DIR", str(automatic.log_root))
    return automatic


def effective_profile_root() -> Path:
    """Return the durable profile identity used to separate local instances."""

    configure_process_profile()
    explicit = os.environ.get("DOCWEN_DATA_DIR", "").strip()
    if explicit:
        return Path(explicit).expanduser().absolute()
    return Path(user_data_dir("docwen", appauthor=False)).absolute()


def profile_instance_name(app_name: str = "docwen") -> str:
    """Derive a stable local-control namespace from the active data profile."""

    root = os.path.normcase(str(effective_profile_root().resolve(strict=False)))
    digest = hashlib.sha256(root.encode("utf-8", errors="strict")).hexdigest()[:16]
    return f"{app_name}-{digest}"


__all__ = [
    "ProfilePaths",
    "configure_process_profile",
    "default_profile_paths",
    "effective_profile_root",
    "profile_instance_name",
]

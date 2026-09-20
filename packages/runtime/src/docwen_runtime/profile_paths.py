"""Resolve and bind one immutable user profile for a process lifetime.

DATA selects a whole profile. CONFIG and LOG are explicit component overrides;
LOG retains its documented root-plus-logs spelling. Resolution never creates
files or mutates environment variables. Entry points bind before composition,
so threads, configuration reloads, template discovery and IPC share the same
paths even if the environment later changes.
"""

from __future__ import annotations

import ctypes
import hashlib
import os
import sys
import tempfile
import threading
from collections.abc import Iterator, Mapping
from contextlib import contextmanager
from dataclasses import dataclass
from pathlib import Path
from types import MappingProxyType
from typing import Literal

import platformdirs

_APPMODEL_ERROR_NO_PACKAGE = 15700
_ERROR_INSUFFICIENT_BUFFER = 122


class ProfileSelectionError(OSError):
    """The selected startup profile cannot be resolved without changing it."""


@dataclass(frozen=True, slots=True)
class ProfilePaths:
    root: Path
    config_dir: Path
    data_dir: Path
    log_dir: Path
    temporary_log_dir: Path
    source: Literal["explicit", "portable", "package", "platform"]
    config_override: bool
    log_override: Literal["DOCWEN_LOG_DIR", "DOCWEN_LOG_TO_TEMP"] | None

    @property
    def portable(self) -> bool:
        return self.source == "portable"


@dataclass(frozen=True, slots=True)
class _ProcessProfile:
    paths: ProfilePaths
    environment: Mapping[str, str]


_bound_profile: _ProcessProfile | None = None
_profile_bindings = 0
_profile_lock = threading.RLock()


def _windows_has_package_identity() -> bool:
    if sys.platform != "win32":
        return False
    try:
        get_name = ctypes.windll.kernel32.GetCurrentPackageFullName
        length = ctypes.c_uint32(0)
        result = int(get_name(ctypes.byref(length), None))
    except (AttributeError, OSError):
        # An unknown package identity never authorizes installation-directory writes.
        return True
    return result == _ERROR_INSUFFICIENT_BUFFER or result != _APPMODEL_ERROR_NO_PACKAGE


def _absolute(value: str | Path) -> Path:
    return Path(value).expanduser().resolve(strict=False)


def _profile_directory(path: Path) -> Path:
    """Reject an existing file in a required directory path without creating it."""
    for candidate in (path, *path.parents):
        if candidate.exists():
            if not candidate.is_dir():
                raise NotADirectoryError(f"Selected DocWen profile path is not a directory: {candidate}")
            break
    return path


def resolve_profile_paths(environment: Mapping[str, str] | None = None) -> ProfilePaths:
    """Resolve effective paths without modifying disk or the supplied environment."""
    environment = os.environ if environment is None else environment
    data = environment.get("DOCWEN_DATA_DIR", "").strip()
    config = environment.get("DOCWEN_CONFIG_DIR", "").strip()
    log = environment.get("DOCWEN_LOG_DIR", "").strip()
    frozen = bool(getattr(sys, "frozen", False))
    source: Literal["explicit", "portable", "package", "platform"]
    if data:
        root, source = _absolute(data), "explicit"
    elif frozen and not _windows_has_package_identity():
        root, source = _absolute(Path(sys.executable).parent / "data"), "portable"
    else:
        root = _absolute(platformdirs.user_data_dir("docwen", appauthor=False))
        source = "package" if frozen else "platform"
    config_dir = (
        _absolute(platformdirs.user_config_dir("docwen", appauthor=False)) / "configs"
        if source == "platform"
        else root / "configs"
    )
    log_dir = _absolute(platformdirs.user_log_dir("docwen", appauthor=False)) if source == "platform" else root / "logs"
    temporary_log_dir = _absolute(tempfile.gettempdir()) / "docwen" / "logs"
    log_override: Literal["DOCWEN_LOG_DIR", "DOCWEN_LOG_TO_TEMP"] | None = None
    if log:
        log_dir, log_override = _absolute(log) / "logs", "DOCWEN_LOG_DIR"
    elif environment.get("DOCWEN_LOG_TO_TEMP", "").strip().lower() in {"1", "true", "yes", "on"}:
        log_dir, log_override = temporary_log_dir, "DOCWEN_LOG_TO_TEMP"
    return ProfilePaths(
        root=root,
        config_dir=_profile_directory(_absolute(config) if config else config_dir),
        data_dir=_profile_directory(root),
        log_dir=log_dir,
        temporary_log_dir=temporary_log_dir,
        source=source,
        config_override=bool(config),
        log_override=log_override,
    )


def current_profile_paths() -> ProfilePaths:
    """Read the bound startup profile; standalone library callers resolve locally."""
    return _bound_profile.paths if _bound_profile is not None else resolve_profile_paths()


@contextmanager
def bind_process_profile() -> Iterator[ProfilePaths]:
    """Bind one profile across the composed entry point, including its threads."""
    global _bound_profile, _profile_bindings
    with _profile_lock:
        if _bound_profile is None:
            environment = dict(os.environ)
            try:
                paths = resolve_profile_paths(environment)
            except (OSError, ValueError) as exc:
                raise ProfileSelectionError(f"Cannot open the selected DocWen profile: {exc}") from exc
            # Child processes may start from a different working directory.
            for key, value in (
                ("DOCWEN_DATA_DIR", paths.data_dir),
                ("DOCWEN_CONFIG_DIR", paths.config_dir),
                ("DOCWEN_LOG_DIR", paths.log_dir.parent),
            ):
                if environment.get(key, "").strip():
                    environment[key] = str(value)
            _bound_profile = _ProcessProfile(paths, MappingProxyType(environment))
        _profile_bindings += 1
        profile = _bound_profile.paths
    try:
        yield profile
    finally:
        with _profile_lock:
            _profile_bindings -= 1
            if _profile_bindings == 0:
                _bound_profile = None


def profile_process_environment() -> dict[str, str]:
    """Launch the matching GUI with the same startup environment, not later edits."""
    return dict(_bound_profile.environment if _bound_profile is not None else os.environ)


def profile_instance_name(app_name: str = "docwen") -> str:
    """Separate instances by the canonical configuration and template stores."""
    profile = current_profile_paths()
    identity = "\0".join(os.path.normcase(str(path)) for path in (profile.config_dir, profile.data_dir))
    digest = hashlib.sha256(identity.encode("utf-8", errors="strict")).hexdigest()[:16]
    return f"{app_name}-{digest}"


__all__ = [
    "ProfilePaths",
    "ProfileSelectionError",
    "bind_process_profile",
    "current_profile_paths",
    "profile_instance_name",
    "profile_process_environment",
    "resolve_profile_paths",
]

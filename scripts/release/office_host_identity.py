"""Read the installed executable/version identity for a reported Office backend."""

from __future__ import annotations

import hashlib
import os
import re
import subprocess
from pathlib import Path
from typing import Any

_PROGIDS = {
    "msoffice_word": "Word.Application",
    "Microsoft Word": "Word.Application",
    "msoffice_excel": "Excel.Application",
    "Microsoft Excel": "Excel.Application",
    "wps_writer": "Kwps.Application",
    "WPS Writer": "Kwps.Application",
    "wps_spreadsheets": "Ket.Application",
    "WPS Spreadsheets": "Ket.Application",
}


def _com_executable(prog_id: str) -> Path:
    import winreg

    command = ""
    for view in (winreg.KEY_WOW64_64KEY, winreg.KEY_WOW64_32KEY):
        try:
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, f"{prog_id}\\CLSID", 0, winreg.KEY_READ | view) as key:
                class_id = str(winreg.QueryValueEx(key, None)[0])
            with winreg.OpenKey(
                winreg.HKEY_CLASSES_ROOT, f"CLSID\\{class_id}\\LocalServer32", 0, winreg.KEY_READ | view
            ) as key:
                command = os.path.expandvars(str(winreg.QueryValueEx(key, None)[0])).strip()
            break
        except FileNotFoundError:
            continue
    match = re.match(r'^"([^"]+\.exe)"|^(.+?\.exe)(?:\s|$)', command, re.IGNORECASE)
    if match is None:
        raise RuntimeError("office_registered_executable_unrecognized")
    return Path(match.group(1) or match.group(2)).resolve(strict=True)


def office_host_identity(backend: str) -> dict[str, Any]:
    """Do not launch COM or modify Office state merely to identify its version."""
    if backend.casefold() == "libreoffice":
        from docwen_core.office_bridge import find_soffice_path

        executable = find_soffice_path()
        if executable is None:
            raise RuntimeError("office_version_backend_missing")
        executable = Path(executable).resolve(strict=True)
        process = subprocess.run(
            [str(executable), "--version"], capture_output=True, text=True, errors="replace", check=True, timeout=20
        )
        version = process.stdout.strip()
        basis = "selected LibreOffice executable --version"
    else:
        if os.name != "nt" or backend not in _PROGIDS:
            raise RuntimeError(f"office_version_backend_unknown:{backend}")
        import win32api

        executable = _com_executable(_PROGIDS[backend])
        metadata = win32api.GetFileVersionInfo(str(executable), "\\")
        high, low = metadata["FileVersionMS"], metadata["FileVersionLS"]
        version = ".".join(str(value) for value in (high >> 16, high & 65535, low >> 16, low & 65535))
        basis = "registered COM server executable FileVersion"
    if not version:
        raise RuntimeError("office_version_empty")
    return {
        "backend": backend,
        "version": version,
        "versionSource": basis,
        "executable": str(executable),
        "bytes": executable.stat().st_size,
        "sha256": hashlib.sha256(executable.read_bytes()).hexdigest(),
    }

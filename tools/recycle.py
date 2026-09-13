"""Recycle a previously validated Windows target, with no deletion fallback."""

from __future__ import annotations

import os
import subprocess
from pathlib import Path


def recycle_directory(path: Path) -> None:
    if os.name != "nt":
        raise OSError("recycle_requires_windows")
    script = """
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName Microsoft.VisualBasic
[Microsoft.VisualBasic.FileIO.FileSystem]::DeleteDirectory(
    $env:DOCWEN_RECYCLE_TARGET,
    [Microsoft.VisualBasic.FileIO.UIOption]::OnlyErrorDialogs,
    [Microsoft.VisualBasic.FileIO.RecycleOption]::SendToRecycleBin)
"""
    environment = {**os.environ, "DOCWEN_RECYCLE_TARGET": str(path)}
    subprocess.run(
        ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", script],
        env=environment,
        check=True,
        creationflags=subprocess.CREATE_NO_WINDOW,
    )
    if path.exists():
        raise OSError(f"recycle_target_remains:{path}")

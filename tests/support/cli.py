"""Shared source-checkout CLI subprocess helpers."""

from __future__ import annotations

import os
import sys
from pathlib import Path


def bundle_cli_command() -> list[str]:
    """Return the installed composition-root command for the active environment."""

    script_name = "docwen.exe" if os.name == "nt" else "docwen"
    script_path = Path(sys.executable).parent / script_name
    if script_path.is_file():
        return [str(script_path)]
    return [
        sys.executable,
        "-c",
        "import sys; from docwen_bundle.cli_entry import main; raise SystemExit(main(sys.argv[1:]))",
    ]

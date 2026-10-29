"""Shared CLI subprocess and controller-projection test helpers."""

from __future__ import annotations

import os
import sys
from pathlib import Path
from unittest.mock import MagicMock


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


def make_controller_mock_with_cancellation() -> MagicMock:
    """Delegate scope ownership to Application while keeping execution mocked.

    This supports CLI request-projection tests only. It does not install a
    Runtime or turn their mocked conversion results into publication evidence.
    """
    from docwen_application.controller import ApplicationController

    controller = MagicMock()
    cancellation_owner = ApplicationController()
    controller.prepare_execution_cancellation.side_effect = cancellation_owner.prepare_execution_cancellation
    controller.release_execution_cancellation.side_effect = cancellation_owner.release_execution_cancellation
    controller.cancel.side_effect = cancellation_owner.cancel
    return controller

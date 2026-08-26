"""Shared terminal and JSON error rendering for execution commands."""

from __future__ import annotations

import argparse
import sys
from typing import Any

from docwen_cli.commands import execution_request
from docwen_cli.exit_codes import ExitCode, exit_code_from_error_code
from docwen_cli.i18n import cli_t


def print_invalid_input(
    action: str,
    args: argparse.Namespace,
    message: str,
    *,
    input_file: str = "",
    error_code: str = "invalid_input",
    details: Any = None,
    hint: str | None = None,
) -> int:
    """Render an input-domain failure and return its stable exit code."""

    del action
    prefix = cli_t("cli.messages.error_prefix")
    full_msg = message if str(message).startswith(str(prefix)) else f"{prefix}: {message}"

    if getattr(args, "json", False):
        from docwen_cli.presenters.json_presenter import JsonPresenter

        presenter = JsonPresenter()
        presenter.present_error(
            execution_request.public_command(args),
            full_msg,
            error_code=error_code,
            details=details,
            hint=hint,
        )
        return int(exit_code_from_error_code(error_code))

    if input_file:
        full_msg = f"{full_msg} ({input_file})"
    print(full_msg, file=sys.stderr)
    if hint:
        print(f"Hint: {hint}", file=sys.stderr)
    return int(exit_code_from_error_code(error_code))


def print_unavailable(
    action: str,
    args: argparse.Namespace,
    message: str,
) -> int:
    """Render a runtime-dependency failure and return its stable exit code."""

    del action
    if getattr(args, "json", False):
        from docwen_cli.presenters.json_presenter import JsonPresenter

        presenter = JsonPresenter()
        presenter.present_error(
            execution_request.public_command(args),
            message,
            error_code="dependency_missing",
        )
    return int(ExitCode.DEPENDENCY_MISSING)

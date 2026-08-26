"""Protocol 3 GUI control command adapter."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Any

from docwen_cli.exit_codes import ExitCode, exit_code_from_error_code
from docwen_cli.gui_control_port import GUI_SETTINGS_SECTIONS, GuiControlError, GuiControlPort
from docwen_cli.parser import bounded_integer, get_common_parser
from docwen_cli.presenters.json_presenter import JsonPresenter


def _present_success(args: argparse.Namespace, command: str, data: dict[str, Any]) -> int:
    if getattr(args, "json", False):
        JsonPresenter(include_timing=getattr(args, "timing", False)).present_data(command, data)
    elif command == "gui status":
        print("running" if data.get("running") else "stopped")
    else:
        print("accepted")
    return int(ExitCode.OK)


def _present_error(args: argparse.Namespace, command: str, error: GuiControlError) -> int:
    if getattr(args, "json", False):
        JsonPresenter(include_timing=getattr(args, "timing", False)).present_error(
            command,
            str(error),
            error_code=error.code,
            details=error.details,
        )
    else:
        print(f"Error: {error}", file=sys.stderr)
    return int(exit_code_from_error_code(error.code))


def execute_gui_control(args: argparse.Namespace, control: GuiControlPort | None = None) -> int:
    command = f"gui {args.gui_command}"
    if control is None:
        return _present_error(
            args,
            command,
            GuiControlError("capability_unavailable", "GUI control is unavailable in this CLI assembly."),
        )

    timeout = float(args.timeout)
    try:
        if args.gui_command == "status":
            data = control.status(timeout=timeout)
        elif args.gui_command == "activate":
            data = control.activate(timeout=timeout)
        elif args.gui_command == "open-settings":
            data = control.open_settings(str(args.section), timeout=timeout)
        else:
            raw_file = getattr(args, "file", None)
            if raw_file is None:
                data = control.open(None, timeout=timeout)
            else:
                path = Path(str(raw_file))
                if not path.is_absolute():
                    raise GuiControlError("invalid_path", "GUI open requires an absolute file path.")
                data = control.open(str(path.resolve()), timeout=timeout)
    except GuiControlError as exc:
        return _present_error(args, command, exc)
    return _present_success(args, command, data)


def register_gui_control_parser(subparsers: Any) -> argparse.ArgumentParser:
    parser = subparsers.add_parser("gui", parents=[get_common_parser()], help="Control the DocWen GUI.")
    commands = parser.add_subparsers(dest="gui_command", required=True)

    open_parser = commands.add_parser("open", parents=[get_common_parser()])
    open_parser.add_argument("file", nargs="?")
    open_parser.add_argument("--timeout", type=bounded_integer(1, 1800, label="timeout"), default=30, metavar="SECONDS")

    activate = commands.add_parser("activate", parents=[get_common_parser()])
    activate.add_argument("--timeout", type=bounded_integer(1, 1800, label="timeout"), default=10, metavar="SECONDS")

    open_settings = commands.add_parser("open-settings", parents=[get_common_parser()])
    open_settings.add_argument("--section", choices=GUI_SETTINGS_SECTIONS, required=True)
    open_settings.add_argument(
        "--timeout",
        type=bounded_integer(1, 1800, label="timeout"),
        default=30,
        metavar="SECONDS",
    )

    status = commands.add_parser("status", parents=[get_common_parser()])
    status.add_argument("--timeout", type=bounded_integer(1, 1800, label="timeout"), default=5, metavar="SECONDS")
    return parser


__all__ = ["execute_gui_control", "register_gui_control_parser"]

"""Schema export derived from the real protocol 3 argument parser."""

from __future__ import annotations

import argparse
import sys
from typing import Any

from docwen_cli.exit_codes import ExitCode
from docwen_cli.parser import get_common_parser
from docwen_cli.presenters.json_presenter import JsonPresenter


def _action_schema(action: argparse.Action) -> dict[str, Any]:
    choices = list(action.choices) if action.choices is not None else None
    value_type = getattr(action.type, "__name__", None) if action.type is not None else None
    if isinstance(action, argparse._StoreTrueAction):
        value_type = "boolean"
    elif isinstance(action, argparse._AppendAction):
        value_type = "array"
    return {
        "name": action.dest,
        "flags": list(action.option_strings),
        "required": bool(action.required),
        "nargs": action.nargs,
        "type": value_type or "string",
        "choices": choices,
        "default": None if action.default is argparse.SUPPRESS else action.default,
    }


def _parser_schema(parser: argparse.ArgumentParser, command: str) -> dict[str, Any]:
    arguments: list[dict[str, Any]] = []
    subcommands: list[str] = []
    for action in parser._actions:
        if isinstance(action, argparse._HelpAction):
            continue
        if isinstance(action, argparse._SubParsersAction):
            subcommands.extend(action.choices)
            continue
        arguments.append(_action_schema(action))
    return {"command": command, "arguments": arguments, "subcommands": subcommands}


def _resolve_parser(root: argparse.ArgumentParser, path: list[str]) -> tuple[argparse.ArgumentParser, str]:
    parser = root
    consumed: list[str] = []
    for segment in path:
        subparsers = next(
            (action for action in parser._actions if isinstance(action, argparse._SubParsersAction)),
            None,
        )
        if subparsers is None or segment not in subparsers.choices:
            raise KeyError(" ".join(path))
        parser = subparsers.choices[segment]
        consumed.append(segment)
    return parser, " ".join(consumed)


def execute_schema_v3(args: argparse.Namespace, controller: Any | None = None) -> int:
    del controller
    from docwen_cli.main import _build_parser

    root = _build_parser()
    path = list(getattr(args, "schema_path", []) or [])
    try:
        parser, command = _resolve_parser(root, path)
    except KeyError:
        requested = " ".join(path)
        message = f"Unknown command path: {requested}"
        if getattr(args, "json", False):
            JsonPresenter().present_error("schema", message, error_code="resource_not_found")
        else:
            print(f"Error: {message}", file=sys.stderr)
        return int(ExitCode.NOT_FOUND)

    data = _parser_schema(parser, command)
    if getattr(args, "json", False):
        JsonPresenter().present_data("schema", data)
    else:
        print(parser.format_help(), end="")
    return int(ExitCode.OK)


def register_schema_v3_parser(subparsers: Any) -> argparse.ArgumentParser:
    parser = subparsers.add_parser(
        "schema",
        parents=[get_common_parser()],
        help="Export a command schema derived from the active parser.",
    )
    parser.add_argument("schema_path", nargs="*", metavar="COMMAND_PATH")
    return parser


__all__ = ["execute_schema_v3", "register_schema_v3_parser"]

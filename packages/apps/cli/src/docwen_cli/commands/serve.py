"""Register and execute the DocWen Machine Protocol stdio server."""

from __future__ import annotations

import argparse
from typing import Any

from docwen_cli.parser import get_common_parser


def register_serve_parser(subparsers: Any) -> None:
    parser = subparsers.add_parser(
        "serve",
        parents=[get_common_parser()],
        help="Run DocWen Machine Protocol v2.",
        description="Run DocWen Machine Protocol v2 over framed stdio.",
    )
    parser.add_argument(
        "--stdio",
        action="store_true",
        required=True,
        help="Use Content-Length framed stdin/stdout transport.",
    )
    parser.add_argument("--managed", action="store_true", help="Use the Windows CLI Job supervisor.")
    parser.add_argument("--parent-pid", type=int, help="Keep the managed worker within this caller's lifetime.")
    parser.set_defaults(command_path="serve")


def execute_serve(args: argparse.Namespace, server: Any) -> int:
    if getattr(args, "managed", False) or getattr(args, "parent_pid", None) is not None:
        raise ValueError("Managed stdio must use DocWenCLI serve --stdio --managed --parent-pid PID")
    if not args.stdio:
        raise ValueError("serve requires --stdio")
    if server is None:
        raise RuntimeError("Machine Protocol server is not configured")
    return int(server.run())


__all__ = ["execute_serve", "register_serve_parser"]

"""Register and execute the DocWen Machine Protocol stdio server."""

from __future__ import annotations

import argparse
from typing import Any

from docwen_cli.parser import get_common_parser


def register_serve_parser(subparsers: Any) -> None:
    parser = subparsers.add_parser(
        "serve",
        parents=[get_common_parser()],
        help="Run DocWen Machine Protocol v1.",
        description="Run DocWen Machine Protocol v1 over framed stdio.",
    )
    parser.add_argument(
        "--stdio",
        action="store_true",
        required=True,
        help="Use Content-Length framed stdin/stdout transport.",
    )
    parser.set_defaults(command_path="serve")


def execute_serve(args: argparse.Namespace, server: Any) -> int:
    if not args.stdio:
        raise ValueError("serve requires --stdio")
    if server is None:
        raise RuntimeError("Machine Protocol server is not configured")
    return int(server.run())


__all__ = ["execute_serve", "register_serve_parser"]

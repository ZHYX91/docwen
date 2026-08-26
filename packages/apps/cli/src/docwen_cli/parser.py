"""Common argument parser — shared parent parser for all CLI commands.

Defines the 9 global parameters listed in CLI契约规范.md §4.
Every command inherits these via ``parents=[get_common_parser()]``.
"""

from __future__ import annotations

import argparse
from typing import Any

from docwen_core.version import PRODUCT_VERSION

# Available locale codes when the runtime locale registry is unavailable.
_AVAILABLE_LOCALES: list[str] = [
    "zh_CN",
    "en_US",
    "de_DE",
    "es_ES",
    "fr_FR",
    "ja_JP",
    "ko_KR",
    "pt_BR",
    "ru_RU",
    "vi_VN",
    "zh_TW",
]

_common_parser: argparse.ArgumentParser | None = None
_common_defaults_parser: argparse.ArgumentParser | None = None

_COMMON_DEFAULTS: dict[str, object] = {
    "lang": None,
    "json": False,
    "quiet": False,
    "verbose": False,
    "timing": False,
}


class CliUsageError(ValueError):
    """Raised for parser usage errors so the entry point owns presentation."""

    def __init__(self, message: str, usage: str) -> None:
        super().__init__(message)
        self.usage = usage


class CliArgumentParser(argparse.ArgumentParser):
    """Production parser with an explicit, stable long-option contract."""

    def __init__(self, *args: Any, **kwargs: Any) -> None:
        # argparse abbreviations are sensitive to whatever options happen to be
        # registered on a level.  The public contract lists exact long names,
        # so reject unstable prefixes and keep pre-parsing consistent.
        kwargs.setdefault("allow_abbrev", False)
        super().__init__(*args, **kwargs)

    def error(self, message: str) -> None:
        """Raise a typed error instead of printing and exiting from argparse."""

        raise CliUsageError(message, self.format_usage())


def get_available_locale_codes() -> list[str]:
    """Return available locale codes for ``--lang`` choices."""
    return list(_AVAILABLE_LOCALES)


def bounded_integer(minimum: int, maximum: int, *, label: str):
    """Return an argparse converter enforcing an inclusive integer range."""

    def convert(raw: str) -> int:
        try:
            value = int(raw)
        except ValueError as exc:
            raise argparse.ArgumentTypeError(f"{label} must be an integer") from exc
        if not minimum <= value <= maximum:
            raise argparse.ArgumentTypeError(f"{label} must be between {minimum} and {maximum}")
        return value

    return convert


def get_common_parser() -> argparse.ArgumentParser:
    """Return the shared parent parser (cached).

    All commands must include this parser via ``parents=[get_common_parser()]``
    so that ``--json``, ``--quiet``, ``--verbose`` etc. are available
    regardless of which command or sub-command is invoked.
    """
    global _common_parser
    if _common_parser is not None:
        return _common_parser

    codes = get_available_locale_codes()
    # A common option is registered at the root and at each nested command so
    # callers can place it at any level.  Suppress absent values here: argparse
    # parses a sub-command into a fresh Namespace and then merges it into its
    # parent, so ordinary False/None/1 defaults would erase an explicit value
    # parsed at an earlier level.  The root-only defaults parent below applies
    # the canonical values exactly once.
    _common_parser = CliArgumentParser(
        add_help=False,
        argument_default=argparse.SUPPRESS,
    )
    _common_parser.add_argument(
        "--version",
        action="version",
        version=f"DocWen {PRODUCT_VERSION} (CLI protocol 3)",
        help="Show the DocWen product and CLI protocol versions.",
    )
    _common_parser.add_argument(
        "--lang",
        choices=codes,
        help="Override interface language (e.g. zh_CN, en_US).",
    )
    _common_parser.add_argument(
        "--json",
        action="store_true",
        help="Produce one protocol 3 JSON document.",
    )
    verbosity = _common_parser.add_mutually_exclusive_group()
    verbosity.add_argument(
        "--quiet",
        "-q",
        action="store_true",
        help="Minimal output — suppress progress messages.",
    )
    verbosity.add_argument(
        "--verbose",
        "-v",
        action="store_true",
        help="Verbose output — show per-file progress.",
    )
    _common_parser.add_argument(
        "--timing",
        action="store_true",
        help="Include per-file timing info in JSON output.",
    )
    return _common_parser


def get_root_common_defaults_parser() -> argparse.ArgumentParser:
    """Return a root-only parent that applies common option defaults once.

    Nested command parsers must inherit only :func:`get_common_parser`.  This
    defaults-only parent deliberately owns no actions, because calling
    ``set_defaults`` on a parser containing the shared actions would mutate
    those action objects and restore the cross-level overwrite bug.
    """
    global _common_defaults_parser
    if _common_defaults_parser is None:
        _common_defaults_parser = CliArgumentParser(add_help=False)
        _common_defaults_parser.set_defaults(**_COMMON_DEFAULTS)
    return _common_defaults_parser

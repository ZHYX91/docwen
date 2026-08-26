"""config command — settings management.

Provides ``config reset`` to restore configuration to defaults via the
registry-driven three-layer config system.
"""

from __future__ import annotations

import argparse
import sys
from typing import Any

from docwen_cli.exit_codes import ExitCode
from docwen_cli.i18n import cli_t
from docwen_cli.parser import get_common_parser

# Maps tab names → the registry's logical reset group.  Runtime owns the
# physical file/key plan so this CLI entry point stays aligned with the GUI.
_TAB_GROUP_MAP: dict[str, str] = {
    "general": "general",
    "output": "output",
    "logging": "logging",
    "export": "export",
    "other": "other",
    "formatting": "formatting",
    "document": "document",
    "text": "text",
    "spreadsheet": "spreadsheet",
    "image": "image",
    "layout": "layout",
    "link": "link",
    "proofread": "proofread",
}

_TABS = list(_TAB_GROUP_MAP.keys())


def register_config_parser(subparsers: Any) -> argparse.ArgumentParser:
    """Register the ``config`` command (container)."""
    p = subparsers.add_parser(
        "config",
        parents=[get_common_parser()],
        help=cli_t("cli.help.settings"),
    )
    sub = p.add_subparsers(dest="config_command", required=True)
    _register_reset_parser(sub)
    return p


def _register_reset_parser(sub: Any) -> None:
    """Register ``reset`` sub-command."""
    sp = sub.add_parser(
        "reset",
        parents=[get_common_parser()],
        help=cli_t("cli.help.settings_reset"),
    )
    sp.add_argument("group", choices=[*_TABS, "all"], metavar="GROUP")
    sp.add_argument("--yes", "-y", action="store_true", required=True, help="Confirm the destructive reset.")


def execute_config(args: argparse.Namespace, controller: Any = None) -> int:
    """Execute config sub-commands."""
    sub_cmd = getattr(args, "config_command", None)

    if sub_cmd is None:
        return int(ExitCode.INTERNAL_ERROR)

    if sub_cmd == "reset":
        return _execute_reset(args, controller)
    else:
        return int(ExitCode.INTERNAL_ERROR)


def _execute_reset(args: argparse.Namespace, controller: Any = None) -> int:
    """Execute ``config reset``.

    When ``--tab`` is given, restores that logical registry group.  Without
    ``--tab``, resets all non-excluded files.
    """
    json_mode = bool(getattr(args, "json", False))
    yes = bool(getattr(args, "yes", False))
    tab = getattr(args, "group", None)
    if tab == "all":
        tab = None
    quiet = bool(getattr(args, "quiet", False))

    group: str | None = None
    if tab is not None:
        group = _TAB_GROUP_MAP.get(tab)
        if group is None:
            message = cli_t("cli.messages.invalid_tab", tab=tab)
            if json_mode:
                from docwen_cli.presenters.json_presenter import JsonPresenter

                presenter = JsonPresenter()
                presenter.present_error(
                    "config reset",
                    message,
                    error_code="invalid_input",
                )
            elif not quiet:
                print(f"{cli_t('cli.messages.error_prefix')}: {message}", file=sys.stderr)
            return int(ExitCode.INVALID_INPUT)

    assert yes  # argparse requires explicit confirmation in every output mode

    config_port = getattr(controller, "config_port", None)
    if config_port is None:
        message = "Configuration service is unavailable in this CLI assembly."
        if json_mode:
            from docwen_cli.presenters.json_presenter import JsonPresenter

            JsonPresenter().present_error(
                "config reset",
                message,
                error_code="config_unavailable",
            )
        elif not quiet:
            print(message, file=sys.stderr)
        return int(ExitCode.UNAVAILABLE)

    if tab is not None:
        assert group is not None
        ok = config_port.reset_group(group)
        success_message = cli_t("settings.reset.tab_success", tab_name=tab)
    else:
        ok = config_port.reset_all()
        success_message = cli_t("settings.reset.success")

    if ok:
        if json_mode:
            from docwen_cli.presenters.json_presenter import JsonPresenter

            presenter = JsonPresenter()
            presenter.present_data(
                "config reset",
                {"action": "reset", "tab": tab, "message": success_message},
                success=True,
            )
        elif not quiet:
            print(success_message)
        return int(ExitCode.OK)
    else:
        failure_message = cli_t("settings.reset.failed")
        if tab is not None:
            failure_message = f"{failure_message} [tab={tab}]"
        if json_mode:
            from docwen_cli.presenters.json_presenter import JsonPresenter

            presenter = JsonPresenter()
            presenter.present_error(
                "config reset",
                failure_message,
                error_code="reset_failed",
                details={"action": "reset", "tab": tab},
            )
        elif not quiet:
            print(failure_message, file=sys.stderr)
        return int(ExitCode.INTERNAL_ERROR)

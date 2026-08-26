"""Focused tests split from test_cli_contract.py."""

from __future__ import annotations

import pytest

from ._cli_contract_support import (
    _make_execution_args,
    argparse,
)
from ._cli_contract_support import (
    _runtime_route_contract as _runtime_route_contract,
)

pytestmark = pytest.mark.unit


class TestResolveExecutionAction:
    """``resolve_cli_action`` is the single source of truth for action resolution."""

    def test_no_action_resolves_empty(self) -> None:
        """An execution request with no internal action resolves to an empty string."""
        from docwen_cli.commands.execution_request import resolve_cli_action

        args = _make_execution_args()
        assert resolve_cli_action(args) == ""

    def test_action_gongwen(self) -> None:
        """The normalized gongwen execution resolves to the gongwen action."""
        from docwen_cli.commands.execution_request import resolve_cli_action

        args = _make_execution_args(action="gongwen")
        assert resolve_cli_action(args) == "gongwen"

    def test_action_validate(self) -> None:
        """The normalized validation execution resolves to the validate action."""
        from docwen_cli.commands.execution_request import resolve_cli_action

        args = _make_execution_args(action="validate")
        assert resolve_cli_action(args) == "validate"

    def test_action_merge_pdfs(self) -> None:
        """The normalized PDF merge execution resolves to merge_pdfs."""
        from docwen_cli.commands.execution_request import resolve_cli_action

        args = _make_execution_args(action="merge_pdfs")
        assert resolve_cli_action(args) == "merge_pdfs"

    def test_canonical_numbering_action_is_unchanged(self) -> None:
        """The parser supplies the canonical runtime action without an alias layer."""
        from docwen_cli.commands.execution_request import resolve_cli_action

        args = _make_execution_args(action="process_md_numbering")
        assert resolve_cli_action(args) == "process_md_numbering"

    def test_action_empty_string_if_arg_missing(self) -> None:
        """Action resolution returns empty string when --action is not set."""
        from docwen_cli.commands.execution_request import resolve_cli_action

        ns = argparse.Namespace()
        assert resolve_cli_action(ns) == ""

    def test_action_whitespace_resolves_empty(self) -> None:
        """Whitespace-only action resolves to empty instead of leaking to runtime routing."""
        from docwen_cli.commands.execution_request import resolve_cli_action

        args = _make_execution_args(action="   ")
        assert resolve_cli_action(args) == ""

"""Named PDF CLI actions must retain the target needed for layout composition."""

from __future__ import annotations

from argparse import Namespace
from types import SimpleNamespace

import pytest

from docwen_application.runtime_capability_catalog import parse_runtime_capability_catalog
from docwen_cli.commands import convert
from docwen_cli.commands.execution_v3 import _prepare_args

from .capability_fixtures import bundled_available_runtime_projection

pytestmark = pytest.mark.contract


@pytest.mark.parametrize(
    ("command", "subcommand", "action"),
    [("merge", "pdf", "merge_pdfs"), ("split", "pdf", "split_pdf")],
)
def test_named_pdf_operation_keeps_concrete_route_target(command: str, subcommand: str, action: str) -> None:
    args = Namespace(
        command=command,
        **{f"{command}_command": subcommand},
        files=["first.xps", "second.pdf"] if command == "merge" else ["first.xps"],
        output="merged.pdf" if command == "merge" else None,
        output_dir="pages" if command == "split" else None,
    )
    _prepare_args(args)
    assert args.action == action
    assert args.to == "pdf"


@pytest.mark.parametrize("source", ["xps", "ofd", "pdf"])
@pytest.mark.parametrize("action", ["split_pdf", "merge_pdfs"])
def test_pdf_actions_resolve_composed_layout_routes(
    monkeypatch: pytest.MonkeyPatch, source: str, action: str,
) -> None:
    catalog = parse_runtime_capability_catalog(bundled_available_runtime_projection())
    monkeypatch.setattr(
        convert,
        "_inspection_for",
        lambda *_args: SimpleNamespace(detected_format=source, workflow_category="layout"),
    )
    routes = convert._resolve_runtime_routes(
        catalog,
        files=["input"],
        inspections={},
        action=action,
        target_format="pdf",
    )
    assert routes["input"].source == "pdf"
    assert routes["input"].target == "pdf"
    assert routes["input"].action_name == action


def test_other_named_actions_do_not_acquire_an_unrelated_pdf_target() -> None:
    args = Namespace(command="validate", file="sample.docx", report=None)
    _prepare_args(args)
    assert args.action == "validate"
    assert args.to == ""

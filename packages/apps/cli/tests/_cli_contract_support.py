"""Contract tests for unified CLI action resolution and JSON output.

Covers the ``plan-cli-contract-polish`` requirements:
- ``resolve_cli_action`` returns ``""`` for no ``--action`` and the action name otherwise
- All execution paths (single/batch/aggregate/dry-run) use the same resolved action
- JSON output separates envelope ``command`` from data ``action_name``
- ``validate`` routing is preserved and tested
"""

from __future__ import annotations

import argparse
from pathlib import Path
from unittest.mock import MagicMock

import pytest
from packages.apps.cli.tests.capability_fixtures import bundled_available_runtime_catalog

from docwen_cli.exit_codes import ExitCode

pytestmark = pytest.mark.unit


@pytest.fixture(autouse=True)
def _runtime_route_contract(monkeypatch: pytest.MonkeyPatch) -> None:
    """Use exact bundled manifest routes instead of permissive controller mocks."""

    catalog = bundled_available_runtime_catalog()
    monkeypatch.setattr("docwen_cli.commands.convert.runtime_route_catalog", lambda _controller: catalog)


def _make_execution_args(**overrides) -> argparse.Namespace:
    """Build a minimal argparse.Namespace mimicking normalized protocol 3 execution."""
    ns = argparse.Namespace()
    ns.command = "convert"
    ns.json = False
    ns.quiet = False
    ns.verbose = False
    ns.timing = False
    ns.batch = False
    ns.jobs = 1
    ns.continue_on_error = False
    ns.output = None
    ns.dry_run = False
    ns.to = "md"
    ns.template = None
    ns.check = []
    ns.extract_img = False
    ns.no_extract_img = False
    ns.ocr = False
    ns.image_mode = None
    ns.image_link_style = None
    ns.table_merge_strategy = None
    ns.ocr_placement = None
    ns.ocr_language = None
    ns.action = ""
    ns.clean_numbering = None
    ns.add_numbering = None
    ns.heading_numbering_render_mode = None
    ns.heading_merge_mode = None
    ns.files = ["/test/doc.md"]
    ns.file = None
    ns.pages = None
    ns.dpi = None
    ns.mode = None
    ns.keep_alpha = False
    for k, v in overrides.items():
        setattr(ns, k, v)
    return ns


def _write_ooxml(path: Path) -> None:
    """Write a structurally valid OOXML package for admission tests."""

    if path.suffix.lower() == ".docx":
        from docx import Document

        Document().save(str(path))
        return
    if path.suffix.lower() == ".xlsx":
        from openpyxl import Workbook

        Workbook().save(str(path))
        return
    raise AssertionError(f"Unsupported OOXML fixture: {path}")


__all__ = (
    "ExitCode",
    "MagicMock",
    "Path",
    "_make_execution_args",
    "_runtime_route_contract",
    "_write_ooxml",
    "argparse",
    "pytest",
    "pytestmark",
)

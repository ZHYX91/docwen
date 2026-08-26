"""Tests for typed option normalization in convert command.

Covers the three normalization helpers and their integration with
``build_execution_options()``:

- ``parse_pages`` — page range string → ``list[int]`` (F-C4-026)
- ``normalize_proofread_options`` — ``--check`` values → boolean flags (F-C4-024, F-C4-030)
- ``normalize_numbering_options`` — clean/add numbering → plugin-compatible booleans (F-C4-023, F-C4-035)
"""

from __future__ import annotations

import argparse

import pytest

pytestmark = pytest.mark.unit


def _fake_convert_args(extra: dict | None = None) -> argparse.Namespace:
    """Build a minimal namespace for normalized protocol 3 execution."""
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
    # Convert-specific
    ns.to = "md"
    ns.template = None
    ns.check = []
    ns.extract_img = False
    ns.no_extract_img = False
    ns.ocr = False
    ns.image_mode = None
    ns.ocr_placement = None
    ns.optimize_for = None
    ns.clean_numbering = None
    ns.add_numbering = None
    ns.heading_numbering_render_mode = None
    ns.heading_merge_mode = None
    ns.files = ["/test/doc.md"]
    ns.file = None
    # Action-specific
    ns.pages = None
    ns.dpi = None
    ns.mode = None
    ns.keep_alpha = False
    ns.spreadsheet_password_prompt = False
    ns.allow_spreadsheet_protection_loss = False

    if extra:
        for k, v in extra.items():
            setattr(ns, k, v)
    return ns


__all__ = (
    "_fake_convert_args",
    "argparse",
    "pytest",
    "pytestmark",
)

"""Repo checks for invoice maintenance/debug tooling."""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

import pytest

pytestmark = pytest.mark.unit

_INVOICE_TOOL_PATHS = (
    Path("tools/invoice_ocr_eval.py"),
    Path("tools/debug_invoice_fitz_insert.py"),
    Path("tools/debug_invoice_multi_page.py"),
)

_FORBIDDEN_SNIPPETS = (
    "docwen.converter",
    "docwen.utils.ocr_utils",
    "layout2md",
    "invoice_cn_ocr",
    "extract_text_simple",
)


def test_invoice_tools_use_current_plugin_entrypoints() -> None:
    project_root = Path(__file__).resolve().parents[2]

    for rel_path in _INVOICE_TOOL_PATHS:
        text = (project_root / rel_path).read_text(encoding="utf-8")
        for forbidden in _FORBIDDEN_SNIPPETS:
            assert forbidden not in text, f"{rel_path} still references {forbidden}"


def test_invoice_ocr_eval_help_is_available() -> None:
    project_root = Path(__file__).resolve().parents[2]
    script = project_root / "tools" / "invoice_ocr_eval.py"

    result = subprocess.run(
        [sys.executable, str(script), "--help"],
        capture_output=True,
        text=True,
        timeout=20,
    )

    assert result.returncode == 0
    assert "usage:" in result.stdout

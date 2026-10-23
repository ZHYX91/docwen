"""Real CLI layout normalization composes with PDF merge and split operations."""

from __future__ import annotations

import hashlib
import json
import os
from pathlib import Path

import fitz
import pytest
from tests.support.cli import bundle_cli_command
from tests.support.subprocess_runner import run_subprocess
from tests.support.xps import create_minimal_xps, pdf_visual_projection

pytestmark = [pytest.mark.e2e, pytest.mark.pr_gate, pytest.mark.release_gate]


@pytest.mark.parametrize("operation", ["split", "merge-before", "merge-after"])
def test_cli_composes_real_xps_with_pdf_action(tmp_path: Path, operation: str) -> None:
    source = tmp_path / "layout 输入.xps"
    create_minimal_xps(source)
    normalized = tmp_path / "expected.pdf"
    with fitz.open(source) as document:
        normalized.write_bytes(document.convert_to_pdf())
    expected_pages = pdf_visual_projection(normalized)["pages"]
    extra = tmp_path / "extra.pdf"
    with fitz.open() as document:
        document.new_page().insert_text((72, 72), "EXTRA PDF PAGE")
        document.save(extra)
    extra_pages = pdf_visual_projection(extra)["pages"]
    before = {path: hashlib.sha256(path.read_bytes()).hexdigest() for path in (source, extra)}
    environment = os.environ.copy()
    environment.update(
        {
            "DOCWEN_DATA_DIR": str(tmp_path / "profile"),
            "DOCWEN_CONFIG_DIR": str(tmp_path / "config"),
            "DOCWEN_LOG_DIR": str(tmp_path / "logs"),
            "DOCWEN_RUNTIME_ROOT": str(tmp_path / "runtime"),
            "PYTHONIOENCODING": "utf-8",
        }
    )
    if operation == "split":
        argv = ["split", "pdf", str(source), "--pages", "1", "--output-dir", str(tmp_path / "pages")]
    else:
        inputs = [extra, source] if operation == "merge-before" else [source, extra]
        argv = ["merge", "pdf", *map(str, inputs), "--output", str(tmp_path / "merged.pdf")]
    result = run_subprocess([*bundle_cli_command(), *argv, "--json"], env=environment, timeout=120)
    assert result.returncode == 0, f"{result.stdout[-2000:]}\n{result.stderr[-2000:]}"
    payload = json.loads(result.stdout)
    assert payload["success"] is True
    artifacts = payload["data"]["artifacts"]
    actual_pages = [page for item in artifacts for page in pdf_visual_projection(item["path"])["pages"]]
    if operation == "split":
        assert len(artifacts) == 2
        assert all(pdf_visual_projection(item["path"])["page_count"] == 1 for item in artifacts)
    else:
        assert len(artifacts) == 1
        expected_pages = extra_pages + expected_pages if operation == "merge-before" else expected_pages + extra_pages
    assert actual_pages == expected_pages
    assert {path: hashlib.sha256(path.read_bytes()).hexdigest() for path in before} == before
    assert not list((tmp_path / "runtime").glob("dw-*"))

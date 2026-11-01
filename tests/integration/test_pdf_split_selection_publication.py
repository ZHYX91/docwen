"""Generated PDF selection contracts through the real CLI/Runtime publisher.

These are newly synthesized files, not copies of a missing review sample.
Independent expected page text and dimensions identify every source page.
"""

from __future__ import annotations

import hashlib
import json
from pathlib import Path
from typing import Any

import fitz
import pytest

from docwen_application.bundle_mapping import build_bundle_draft
from docwen_application.controller import ApplicationController
from docwen_bundle.runtime_factory import create_runtime_port
from docwen_cli.main import main
from docwen_core.models.file_ref import FileRef
from docwen_core.models.request import ConversionRequest, OutputPolicy
from docwen_runtime.output.artifact_bundle import ArtifactBundleCommitter

pytestmark = [pytest.mark.integration, pytest.mark.pr_gate]


def _new_pdf(path: Path) -> bytes:
    with fitz.open() as document:
        for number in range(1, 5):
            page = document.new_page(width=200 + number * 10, height=300 + number * 10)
            page.insert_text((25, 45), f"ORIGINAL PAGE {number}", fontsize=10)
        document.save(path)
    original = path.read_bytes()
    assert _page_facts(path) == _expected([1, 2, 3, 4])
    return original


def _page_facts(path: Path) -> list[tuple[str, float, float]]:
    with fitz.open(path) as document:
        return [(page.get_text().strip(), page.rect.width, page.rect.height) for page in document]


def _expected(numbers: list[int]) -> list[tuple[str, float, float]]:
    return [(f"ORIGINAL PAGE {number}", 200 + number * 10, 300 + number * 10) for number in numbers]


@pytest.mark.parametrize("pages", [[1, 999], [0, 1], [-1, 1], [1, 1.5], [True, 1], [1, "3"], [], [1, 2, 3, 4]])
def test_runtime_rejects_the_whole_invalid_pdf_selection(
    tmp_path: Path, round_trip_runtime: Any, pages: list[object]
) -> None:
    source = tmp_path / "source.pdf"
    original = _new_pdf(source)
    output = tmp_path / "published"
    output.mkdir()
    sentinel = output / "already-published.txt"
    sentinel.write_bytes(b"previous output")
    request = ConversionRequest(
        request_id="strict-page-selection-negative",
        input_refs=[FileRef(path=str(source), format="pdf", category="layout", size_bytes=len(original))],
        target_format="pdf",
        action_name="split_pdf",
        options={"split_mode": "custom", "pages": pages},
        output_policy=OutputPolicy(output_dir=str(output)),
    )
    result = round_trip_runtime.execute(request)
    assert not result.success
    assert result.error is not None
    assert result.artifacts == []
    assert sorted(path.name for path in output.iterdir()) == [sentinel.name]
    assert sentinel.read_bytes() == b"previous output"
    assert source.read_bytes() == original


@pytest.mark.parametrize(
    ("mode", "pages", "groups"),
    [
        ("custom", [1, 3], [[1, 3], [2, 4]]),
        ("custom", [3, 1, 3], [[1, 3], [2, 4]]),
        ("every_page", [], [[1], [2], [3], [4]]),
        ("odd_even", [], [[1, 3], [2, 4]]),
    ],
)
def test_runtime_pdf_groups_match_independent_source_pages_after_bundle_publication(
    tmp_path: Path, round_trip_runtime: Any, mode: str, pages: list[int], groups: list[list[int]]
) -> None:
    source = tmp_path / "source.pdf"
    original = _new_pdf(source)
    output = tmp_path / "published"
    request = ConversionRequest(
        request_id="strict-page-selection-positive",
        input_refs=[FileRef(path=str(source), format="pdf", category="layout", size_bytes=len(original))],
        target_format="pdf",
        action_name="split_pdf",
        options={"split_mode": mode, "pages": pages},
        output_policy=OutputPolicy(output_dir=str(output)),
    )
    result = round_trip_runtime.execute(request)
    assert result.success, result.error
    profile = "section_documents" if mode == "every_page" else "partition_documents"
    bundle = ArtifactBundleCommitter().commit(
        task_id=request.request_id,
        staging_root=str(output),
        draft=build_bundle_draft(profile=profile, output_media_type="application/pdf", artifacts=result.artifacts),
    )
    ordered = sorted(bundle.entries, key=lambda item: item.ordinal)
    by_id = {item.artifact_id: item for item in bundle.artifacts}
    facts = []
    for entry in ordered:
        artifact = by_id[entry.artifact_id]
        path = output / artifact.locator
        content = path.read_bytes()
        assert hashlib.sha256(content).hexdigest() == artifact.sha256
        assert len(content) == artifact.size_bytes
        facts.append(_page_facts(path))
    assert facts == [_expected(numbers) for numbers in groups]
    assert source.read_bytes() == original


@pytest.mark.parametrize("expression", ["1,3x", "1,999", "1,,3", "0,1", "1-3-4"])
def test_cli_invalid_pdf_selection_publishes_nothing(
    tmp_path: Path, capsys: pytest.CaptureFixture[str], expression: str
) -> None:
    source = tmp_path / "source.pdf"
    original = _new_pdf(source)
    output = tmp_path / "published"
    controller = ApplicationController(runtime_port=create_runtime_port())
    controller.start()
    try:
        exit_code = main(
            ["split", "pdf", str(source), "--pages", expression, "--output-dir", str(output), "--json"],
            controller=controller,
        )
    finally:
        controller.stop()
    captured = capsys.readouterr()
    payload = json.loads(captured.out)
    assert exit_code != 0
    assert payload["success"] is False
    assert not list(output.rglob("*.pdf"))
    assert source.read_bytes() == original


def test_cli_valid_pdf_selection_keeps_each_whole_selected_page(
    tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    source = tmp_path / "source.pdf"
    original = _new_pdf(source)
    output = tmp_path / "published"
    controller = ApplicationController(runtime_port=create_runtime_port())
    controller.start()
    try:
        exit_code = main(
            ["split", "pdf", str(source), "--pages", "3,1,3", "--output-dir", str(output), "--json"],
            controller=controller,
        )
    finally:
        controller.stop()
    payload = json.loads(capsys.readouterr().out)
    assert exit_code == 0, payload
    assert payload["success"] is True
    assert [_page_facts(path) for path in sorted(output.rglob("*.pdf"))] == [_expected([1, 3]), _expected([2, 4])]
    assert source.read_bytes() == original

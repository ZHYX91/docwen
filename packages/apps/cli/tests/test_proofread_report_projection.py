"""Proofread report projection is identical for every JSON presenter path."""

from __future__ import annotations

import json
import os
from pathlib import Path

import pytest

pytestmark = pytest.mark.contract


def _report() -> dict[str, object]:
    return {
        "schema": "docwen.proofread_report.v2",
        "file": "input.md",
        "source": {
            "content_sha256": "a" * 64,
            "encoding": "utf-8",
            "decode_errors": "replace",
        },
        "location_contract": {
            "id": "docwen.proofread-text-range",
            "version": 1,
            "coordinate_system": "unicode_code_point",
            "offset_base": 0,
            "line_base": 0,
            "column_base": 0,
            "range_end": "exclusive",
        },
        "checks_enabled": {
            "symbol_pairing": False,
            "symbol_correction": False,
            "typos_rule": True,
            "sensitive_word": False,
        },
        "issues": [],
        "summary": {},
    }


def _result(report_path: Path, *, inline: bool):
    from docwen_core.models.artifact import ArtifactManifest
    from docwen_core.models.result import ConversionMetrics, ConversionResult

    report = _report()
    from docwen_runtime.path_io import filesystem_path

    filesystem_path(report_path).write_text(json.dumps(report, ensure_ascii=False), encoding="utf-8")
    artifact = ArtifactManifest(
        artifact_id="proofread-report",
        kind="primary",
        staging_path=str(report_path),
        suggested_name=report_path.name,
        media_type="application/json",
        metadata={"issues_found": 0},
        is_primary=True,
    )
    metrics = ConversionMetrics(extra={"proofread_report": report} if inline else {})
    return ConversionResult(
        task_id="proofread-report",
        success=True,
        artifacts=[artifact],
        metrics=metrics,
    )


def test_inline_file_and_batch_projection_preserve_the_same_report(
    tmp_path: Path,
    capsys: pytest.CaptureFixture[str],
) -> None:
    from docwen_cli.presenters.json_presenter import JsonPresenter

    inline_result = _result(tmp_path / "inline.json", inline=True)
    file_result = _result(tmp_path / "file.json", inline=False)

    JsonPresenter().present_single(inline_result, command="validate markdown")
    inline_payload = json.loads(capsys.readouterr().out)
    JsonPresenter().present_single(file_result, command="validate markdown")
    file_payload = json.loads(capsys.readouterr().out)
    JsonPresenter().present_batch(
        [file_result],
        command="batch validate",
        input_files=[str(tmp_path / "input.md")],
    )
    batch_payload = json.loads(capsys.readouterr().out)

    expected = _report()
    assert inline_payload["data"]["details"]["proofread"] == expected
    assert file_payload["data"]["details"]["proofread"] == expected
    assert batch_payload["data"]["results"][0]["details"]["proofread"] == expected


@pytest.mark.skipif(os.name != "nt", reason="Win32 extended-path boundary")
def test_exact_260_unit_report_projects_in_single_and_batch_protocols(
    tmp_path: Path,
    capsys: pytest.CaptureFixture[str],
) -> None:
    from docwen_cli.presenters.json_presenter import JsonPresenter
    from docwen_runtime.path_io import filesystem_path, windows_utf16_units

    root_units = windows_utf16_units(os.path.abspath(tmp_path))
    suffix = ".json"
    filename_units = 260 - root_units - 1
    if filename_units < len(suffix) + 1 or filename_units > 255:
        pytest.skip("pytest temp root cannot express an exact 260-unit report path")
    report_path = tmp_path / f"{'r' * (filename_units - len(suffix))}{suffix}"
    result = _result(report_path, inline=False)

    assert windows_utf16_units(os.path.abspath(report_path)) == 260
    assert filesystem_path(report_path).is_file()
    assert not result.artifacts[0].staging_path.startswith("\\\\?\\")

    JsonPresenter().present_single(result, command="validate markdown")
    single_payload = json.loads(capsys.readouterr().out)
    JsonPresenter().present_batch(
        [result],
        command="batch validate",
        input_files=[str(tmp_path / "input.md")],
    )
    batch_payload = json.loads(capsys.readouterr().out)

    expected = _report()
    assert single_payload["data"]["details"]["proofread"] == expected
    assert batch_payload["data"]["results"][0]["details"]["proofread"] == expected

"""Direct contracts for the optional packaged proofread report 2.0 gate."""

from __future__ import annotations

import json
import subprocess
import sys
from hashlib import sha256
from pathlib import Path
from typing import Any

import pytest

pytestmark = pytest.mark.unit


def _position(text: str, offset: int) -> dict[str, int]:
    starts = [0]
    index = 0
    while index < len(text):
        if text[index] == "\r":
            if index + 1 < len(text) and text[index + 1] == "\n":
                starts.append(index + 2)
                index += 2
                continue
            starts.append(index + 1)
        elif text[index] == "\n":
            starts.append(index + 1)
        index += 1
    line = max(index for index, start in enumerate(starts) if start <= offset)
    return {"offset": offset, "line": line, "column": offset - starts[line]}


def _issue(
    text: str,
    matched: str,
    *,
    source: str,
    rule_key: str,
    replacement: str | None = None,
) -> dict[str, Any]:
    start = text.index(matched)
    end = start + len(matched)
    start_position = _position(text, start)
    end_position = _position(text, end)
    issue: dict[str, Any] = {
        "range": {"start": start_position, "end": end_position},
        "matched_text": matched,
        "error_text": matched,
        "suggestion": replacement or "suggestion only",
        "error_type": rule_key,
        "source": source,
        "rule_key": rule_key,
    }
    if replacement is not None:
        issue["fix"] = {
            "kind": "replace_text",
            "replacement": replacement,
            "applicable": True,
        }
    return issue


def _report(source_path: Path, *, empty: bool) -> dict[str, Any]:
    source_bytes = source_path.read_bytes()
    text = source_bytes.decode("utf-8", errors="replace")
    issues = (
        []
        if empty
        else [
            _issue(text, "１", source="symbol", rule_key="symbol_correct", replacement="1"),
            _issue(text, "２", source="symbol", rule_key="symbol_correct", replacement="2"),
            _issue(text, "（", source="pairing", rule_key="symbol_pair"),
        ]
    )
    return {
        "schema": "docwen.proofread_report.v2",
        "file": source_path.name,
        "source": {
            "content_sha256": sha256(source_bytes).hexdigest(),
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
            "symbol_pairing": not empty,
            "symbol_correction": not empty,
            "typos_rule": False,
            "sensitive_word": False,
        },
        "issues": issues,
        "summary": {} if empty else {"symbol_correct": 2, "symbol_pair": 1},
    }


def test_packaged_proofread_report_smoke_uses_real_validate_syntax_and_preserves_bytes(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from scripts.release import verify_packaged_cli

    binary_path = tmp_path / "DocWenCLI.exe"
    binary_path.write_bytes(b"binary")
    calls: list[tuple[str, ...]] = []

    def fake_run(binary: Path, *args: str, cwd: Path) -> subprocess.CompletedProcess[str]:
        calls.append(args)
        assert binary == binary_path
        assert cwd == tmp_path
        source_path = Path(args[1])
        report_path = Path(args[args.index("--report") + 1])
        report = _report(source_path, empty=args[args.index("--check") + 1] == "none")
        report_path.write_text(json.dumps(report, ensure_ascii=False), encoding="utf-8")
        payload = {
            "protocol_version": 3,
            "success": True,
            "command": "validate",
            "data": {"output": str(report_path), "details": {"proofread": report}},
            "error": None,
        }
        return subprocess.CompletedProcess([str(binary), *args], 0, stdout=json.dumps(payload), stderr="")

    monkeypatch.setattr(verify_packaged_cli, "_run", fake_run)

    populated, empty = verify_packaged_cli._run_optional_proofread_report_smoke(binary_path, work_dir=tmp_path)

    source = tmp_path / "proofread-report-2.0-coordinates.md"
    raw = source.read_bytes()
    assert raw == verify_packaged_cli._PROOFREAD_REPORT_FIXTURE_TEXT.encode("utf-8")
    assert raw.startswith(b"\xef\xbb\xbf")
    assert b"\r\n" in raw
    decoded = raw.decode("utf-8")
    for marker in ("😀", "e\u0301", "👩\u200d💻", "１２", "（"):
        assert marker in decoded
    assert populated.is_file()
    assert empty.is_file()
    assert len(calls) == 2
    assert calls[0][:2] == ("validate", str(source))
    assert calls[0][2:6] == ("--check", "symbol", "--check", "punct")
    assert calls[1][2:4] == ("--check", "none")
    assert all(call[-2:] == ("--json", "--quiet") for call in calls)


def test_packaged_proofread_report_rejects_range_or_fix_drift(tmp_path: Path) -> None:
    from scripts.release import verify_packaged_cli

    source = tmp_path / "fixture.md"
    source_bytes = verify_packaged_cli._write_proofread_report_fixture(source)
    report = _report(source, empty=False)
    report["issues"][0]["range"]["end"]["offset"] += 1

    with pytest.raises(RuntimeError, match="position_mismatch"):
        verify_packaged_cli._verify_proofread_report(
            report,
            source_path=source,
            source_bytes=source_bytes,
            expected_checks=report["checks_enabled"],
            expect_fixture_issues=True,
        )

    report = _report(source, empty=False)
    report["issues"][2]["fix"] = {"kind": "replace_text", "replacement": "）", "applicable": True}
    with pytest.raises(RuntimeError, match="fix_invalid"):
        verify_packaged_cli._verify_proofread_report(
            report,
            source_path=source,
            source_bytes=source_bytes,
            expected_checks=report["checks_enabled"],
            expect_fixture_issues=True,
        )


def test_packaged_proofread_report_does_not_fold_cli_failure_into_empty_success(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from scripts.release import verify_packaged_cli

    binary_path = tmp_path / "DocWenCLI.exe"
    binary_path.write_bytes(b"binary")
    failure = {
        "protocol_version": 3,
        "success": False,
        "command": "validate",
        "data": {},
        "error": {"category": "unavailable", "code": "capability_unavailable", "message": "unavailable"},
    }
    monkeypatch.setattr(
        verify_packaged_cli,
        "_run",
        lambda *_args, **_kwargs: subprocess.CompletedProcess(
            [str(binary_path), "validate"],
            6,
            stdout=json.dumps(failure),
            stderr="",
        ),
    )

    with pytest.raises(RuntimeError, match="failed with exit code 6"):
        verify_packaged_cli._run_optional_proofread_report_smoke(binary_path, work_dir=tmp_path)


def test_packaged_proofread_report_smoke_is_documented_and_in_help() -> None:
    process = subprocess.run(
        [sys.executable, "scripts/release/verify_packaged_cli.py", "--help"],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
    )
    packaging = Path("docs/packaging.md").read_text(encoding="utf-8")
    testing = Path("docs/testing.md").read_text(encoding="utf-8")

    assert process.returncode == 0, process.stderr
    assert "--proofread-report-smoke" in process.stdout
    assert "--proofread-report-smoke" in packaging
    assert "--proofread-report-smoke" in testing
    assert "Unicode code-point" in packaging
    assert "成功空报告" in testing

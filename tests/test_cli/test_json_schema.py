"""CLI protocol 3 JSON schema and entry-point integration checks."""

from __future__ import annotations

import json
import os
from pathlib import Path

import pytest
from tests.support.cli import bundle_cli_command
from tests.support.subprocess_runner import run_subprocess


@pytest.fixture(autouse=True)
def _isolate_cli_process_data(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    """Keep source CLI subprocesses away from the real user's data roots."""

    runtime_root = tmp_path.parent / f"{tmp_path.name}-runtime"
    monkeypatch.setenv("DOCWEN_CONFIG_DIR", str(runtime_root / "config_home"))
    monkeypatch.setenv("DOCWEN_LOG_DIR", str(runtime_root / "log_home"))


def _repo_root() -> Path:
    return Path(__file__).resolve().parents[2]


def _load_schema() -> dict:
    path = _repo_root() / "docs" / "specs" / "json-contracts.schema.json"
    return json.loads(path.read_text(encoding="utf-8"))


def _extract_json_code_blocks(markdown: str) -> list[str]:
    blocks: list[str] = []
    in_json = False
    buffer: list[str] = []
    for line in markdown.splitlines():
        if not in_json and line.strip() == "```json":
            in_json = True
            buffer = []
            continue
        if in_json and line.strip() == "```":
            blocks.append("\n".join(buffer).strip())
            in_json = False
            continue
        if in_json:
            buffer.append(line)
    return [block for block in blocks if block]


def _assert_protocol_3_envelope(payload: dict[str, object]) -> None:
    schema = _load_schema()
    required = set(schema["required"])
    assert set(payload) == required
    assert payload["protocol_version"] == 3
    assert isinstance(payload["product_version"], str)
    assert isinstance(payload["success"], bool)
    assert isinstance(payload["command"], str)
    assert isinstance(payload["data"], dict)
    assert isinstance(payload["warnings"], list)
    assert isinstance(payload["meta"], dict)
    error = payload["error"]
    if error is not None:
        assert isinstance(error, dict)
        assert set(error) == {"category", "code", "message", "details", "hint"}
        categories = set(schema["properties"]["error"]["anyOf"][1]["properties"]["category"]["enum"])
        assert error["category"] in categories


def _run_cli(*args: str) -> dict[str, object]:
    repo_root = _repo_root()
    env = os.environ.copy()
    env["PYTHONUTF8"] = "1"
    env["PYTHONIOENCODING"] = "utf-8"
    proc = run_subprocess(
        [*bundle_cli_command(), *args],
        cwd=str(repo_root),
        env=env,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
    )
    assert proc.returncode != 0
    assert proc.stderr == ""
    payload = json.loads(proc.stdout)
    assert isinstance(payload, dict)
    return payload


@pytest.mark.contract
def test_json_contracts_doc_exists() -> None:
    path = _repo_root() / "docs" / "specs" / "json-contracts.md"
    assert path.is_file()
    assert "JSON contracts / JSON 契约" in path.read_text(encoding="utf-8")


@pytest.mark.contract
def test_cli_json_schema_is_exact_protocol_3_envelope() -> None:
    schema = _load_schema()
    assert schema["type"] == "object"
    assert schema["additionalProperties"] is False
    assert schema["properties"]["protocol_version"]["const"] == 3
    assert set(schema["required"]) == {
        "protocol_version",
        "product_version",
        "success",
        "command",
        "data",
        "error",
        "warnings",
        "meta",
    }


@pytest.mark.contract
def test_json_contracts_md_examples_conform_to_protocol_3_shape() -> None:
    markdown = (_repo_root() / "docs" / "specs" / "json-contracts.md").read_text(encoding="utf-8")
    blocks = _extract_json_code_blocks(markdown)
    assert len(blocks) >= 2
    for raw in blocks:
        payload = json.loads(raw)
        assert isinstance(payload, dict)
        _assert_protocol_3_envelope(payload)


@pytest.mark.contract
def test_protocol_3_golden_fixtures_use_the_exact_envelope() -> None:
    golden_dir = _repo_root() / "tests" / "fixtures" / "golden"
    paths = sorted(golden_dir.glob("cli_protocol_v3_*.json"))
    assert [path.name for path in paths] == [
        "cli_protocol_v3_batch.json",
        "cli_protocol_v3_error.json",
        "cli_protocol_v3_success.json",
    ]
    for path in paths:
        payload = json.loads(path.read_text(encoding="utf-8"))
        assert isinstance(payload, dict)
        _assert_protocol_3_envelope(payload)


@pytest.mark.integration
@pytest.mark.pr_gate
def test_invalid_input_uses_protocol_3_schema(tmp_path: Path) -> None:
    payload = _run_cli(
        "convert",
        str(tmp_path / "missing.docx"),
        "--to",
        "md",
        "--output",
        str(tmp_path / "out.md"),
        "--json",
    )
    _assert_protocol_3_envelope(payload)
    assert payload["command"] == "convert"
    assert payload["error"]["category"] == "invalid_input"


@pytest.mark.integration
def test_invalid_option_combo_uses_protocol_3_schema(tmp_path: Path) -> None:
    payload = _run_cli(
        "convert",
        str(tmp_path / "missing.docx"),
        "--to",
        "md",
        "--output",
        str(tmp_path / "out.md"),
        "--ocr",
        "--no-extract-img",
        "--json",
    )
    _assert_protocol_3_envelope(payload)
    assert payload["error"]["category"] == "invalid_input"


@pytest.mark.integration
def test_invalid_format_uses_protocol_3_schema(tmp_path: Path) -> None:
    payload = _run_cli(
        "convert",
        str(tmp_path / "missing.docx"),
        "--to",
        "nope",
        "--output",
        str(tmp_path / "out.nope"),
        "--json",
    )
    _assert_protocol_3_envelope(payload)
    assert payload["error"]["category"] == "invalid_input"

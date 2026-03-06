"""CLI JSON 输出 schema 的单元测试与契约校验。"""

from __future__ import annotations

import json
import os
import subprocess
import sys
from pathlib import Path

import pytest

pytestmark = pytest.mark.unit


def _repo_root() -> Path:
    return Path(__file__).resolve().parents[2]


def _load_schema() -> dict:
    path = _repo_root() / "doc" / "cli_json_output_schema.json"
    return json.loads(path.read_text(encoding="utf-8"))


def _load_json(path: Path) -> object:
    with path.open("r", encoding="utf-8") as f:
        return json.load(f)


def _extract_json_code_blocks(markdown: str) -> list[str]:
    blocks: list[str] = []
    in_json = False
    buf: list[str] = []

    for line in markdown.splitlines():
        if not in_json and line.strip() == "```json":
            in_json = True
            buf = []
            continue
        if in_json and line.strip() == "```":
            blocks.append("\n".join(buf).strip())
            in_json = False
            buf = []
            continue
        if in_json:
            buf.append(line)

    return [b for b in blocks if b]


def _is_type(value: object, schema_type: str) -> bool:
    if schema_type == "null":
        return value is None
    if schema_type == "boolean":
        return isinstance(value, bool)
    if schema_type == "string":
        return isinstance(value, str)
    if schema_type == "object":
        return isinstance(value, dict)
    if schema_type == "array":
        return isinstance(value, list)
    if schema_type == "integer":
        return isinstance(value, int) and not isinstance(value, bool)
    if schema_type == "number":
        return isinstance(value, (int, float)) and not isinstance(value, bool)
    return False


def _validate_required_and_types(schema_def: dict, instance: dict) -> None:
    required = schema_def.get("required") or []
    props = schema_def.get("properties") or {}

    for key in required:
        assert key in instance, f"missing required key: {key}"

    for key, prop in props.items():
        if key not in instance:
            continue
        expected = prop.get("type")
        if not expected:
            continue
        types = expected if isinstance(expected, list) else [expected]
        assert any(_is_type(instance[key], t) for t in types), f"type mismatch for {key}: {instance[key]!r} not in {types}"


def test_cli_json_output_schema_doc_exists() -> None:
    repo_root = _repo_root()
    path = repo_root / "doc" / "cli_json_output_schema.md"
    assert path.is_file()
    text = path.read_text(encoding="utf-8")
    assert "CLI JSON 输出契约" in text


def test_cli_json_schema_file_exists_and_has_expected_defs() -> None:
    schema = _load_schema()

    assert schema.get("$schema")
    assert schema.get("type") == "object"
    props = schema.get("properties") or {}
    assert isinstance(props, dict)
    assert props.get("schema_version", {}).get("const") == 2
    for key in ["schema_version", "success", "command", "data", "error", "warnings", "timing"]:
        assert key in props


def test_json_schema_required_keys_match_golden_fixtures() -> None:
    repo_root = _repo_root()
    schema_path = repo_root / "doc" / "cli_json_output_schema.json"
    golden_dir = repo_root / "tests" / "fixtures" / "golden"

    schema = _load_json(schema_path)
    assert isinstance(schema, dict)
    required = set(schema.get("required") or [])
    assert required

    for path in sorted(golden_dir.glob("*.json")):
        data = _load_json(path)
        assert isinstance(data, dict)
        keys = set(data.keys())
        assert required.issubset(keys), f"{path.name} missing {sorted(required - keys)}"


def test_cli_json_output_schema_md_examples_conform_to_schema() -> None:
    repo_root = _repo_root()
    schema = _load_schema()

    md_path = repo_root / "doc" / "cli_json_output_schema.md"
    markdown = md_path.read_text(encoding="utf-8")

    blocks = _extract_json_code_blocks(markdown)
    assert blocks

    for raw in blocks:
        data = json.loads(raw)
        assert isinstance(data, dict)
        _validate_required_and_types(schema, data)
        assert data.get("schema_version") == 2
        error = data.get("error")
        assert (error is None) or (
            isinstance(error, dict) and {"error_code", "message", "details"}.issubset(set(error.keys()))
        )


@pytest.mark.integration
def test_cli_json_output_conforms_to_schema_on_invalid_input(tmp_path: Path) -> None:
    repo_root = _repo_root()
    schema = _load_schema()

    missing_file = tmp_path / "missing.docx"
    assert not missing_file.exists()

    env = os.environ.copy()
    env["PYTHONUTF8"] = "1"
    env["PYTHONIOENCODING"] = "utf-8"
    env["PYTHONPATH"] = f"{repo_root / 'src'}{os.pathsep}{env.get('PYTHONPATH', '')}"

    cmd = [
        sys.executable,
        "-m",
        "docwen.cli_run",
        "convert",
        str(missing_file),
        "--to",
        "md",
        "--json",
        "--quiet",
    ]
    proc = subprocess.run(
        cmd,
        cwd=str(repo_root),
        env=env,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
    )

    assert proc.returncode != 0
    stdout = proc.stdout.strip()
    assert stdout

    data = json.loads(stdout)
    assert isinstance(data, dict)
    _validate_required_and_types(schema, data)


@pytest.mark.integration
def test_cli_json_output_conforms_to_schema_on_invalid_option_combo(tmp_path: Path) -> None:
    repo_root = _repo_root()
    schema = _load_schema()

    missing_file = tmp_path / "missing.docx"
    assert not missing_file.exists()

    env = os.environ.copy()
    env["PYTHONUTF8"] = "1"
    env["PYTHONIOENCODING"] = "utf-8"
    env["PYTHONPATH"] = f"{repo_root / 'src'}{os.pathsep}{env.get('PYTHONPATH', '')}"

    cmd = [
        sys.executable,
        "-m",
        "docwen.cli_run",
        "convert",
        str(missing_file),
        "--to",
        "md",
        "--ocr",
        "--no-extract-img",
        "--json",
        "--quiet",
    ]
    proc = subprocess.run(
        cmd,
        cwd=str(repo_root),
        env=env,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
    )

    assert proc.returncode != 0
    stdout = proc.stdout.strip()
    assert stdout

    data = json.loads(stdout)
    assert isinstance(data, dict)
    _validate_required_and_types(schema, data)


@pytest.mark.integration
def test_cli_json_output_conforms_to_schema_on_convert_invalid_to(tmp_path: Path) -> None:
    repo_root = _repo_root()
    schema = _load_schema()

    missing_file = tmp_path / "missing.docx"
    assert not missing_file.exists()

    env = os.environ.copy()
    env["PYTHONUTF8"] = "1"
    env["PYTHONIOENCODING"] = "utf-8"
    env["PYTHONPATH"] = f"{repo_root / 'src'}{os.pathsep}{env.get('PYTHONPATH', '')}"

    cmd = [
        sys.executable,
        "-m",
        "docwen.cli_run",
        "convert",
        str(missing_file),
        "--to",
        "nope",
        "--json",
        "--quiet",
    ]
    proc = subprocess.run(
        cmd,
        cwd=str(repo_root),
        env=env,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
    )

    assert proc.returncode != 0
    stdout = proc.stdout.strip()
    assert stdout

    data = json.loads(stdout)
    assert isinstance(data, dict)
    _validate_required_and_types(schema, data)

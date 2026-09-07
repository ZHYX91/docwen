"""Direct end-to-end smoke tests for the current CLI entry point."""

from __future__ import annotations

import json
import os
import subprocess
import sys
from pathlib import Path

import pytest

pytestmark = pytest.mark.e2e

REPO_ROOT = Path(__file__).resolve().parents[2]
BUNDLE_SRC = REPO_ROOT / "packages" / "bundle" / "src"
SAMPLE_MD = REPO_ROOT / "samples" / "sample.md"


def _run_cli(args: list[str], *, runtime_root: Path, timeout: int = 120) -> subprocess.CompletedProcess[str]:
    encoded_args = json.dumps(args)
    script = (
        "import json,sys;"
        f"sys.path.insert(0, {str(BUNDLE_SRC)!r});"
        "from docwen_bundle.cli_entry import main;"
        "sys.exit(main(json.loads(sys.argv[1])))"
    )
    environment = os.environ.copy()
    environment.update(
        {
            "DOCWEN_CONFIG_DIR": str(runtime_root / "config"),
            "DOCWEN_LOG_DIR": str(runtime_root / "logs"),
            "PYTHONIOENCODING": "utf-8",
        }
    )
    return subprocess.run(
        [sys.executable, "-c", script, encoded_args],
        capture_output=True,
        text=True,
        encoding="utf-8",
        timeout=timeout,
        cwd=REPO_ROOT,
        env=environment,
        check=False,
    )


@pytest.mark.parametrize(
    "args",
    [
        pytest.param(["--help"], id="help"),
        pytest.param(["doctor", "--json"], id="doctor-json"),
        pytest.param(["inspect", str(SAMPLE_MD)], id="inspect-markdown"),
        pytest.param(["resources", "list", "formats"], id="formats"),
        pytest.param(["resources", "list", "formats", "--json"], id="formats-json"),
    ],
)
def test_cli_commands_start_and_succeed(args: list[str], tmp_path: Path) -> None:
    completed = _run_cli(args, runtime_root=tmp_path)

    assert completed.returncode == 0, completed.stderr[-4000:]


def test_cli_markdown_docx_roundtrip(tmp_path: Path) -> None:
    output_docx_parent = tmp_path / "generated"
    output_directory = tmp_path / "roundtrip"

    converted = _run_cli(
        ["convert", str(SAMPLE_MD), "--to", "docx", "--output-dir", str(output_docx_parent)],
        runtime_root=tmp_path,
    )
    assert converted.returncode == 0, converted.stderr[-4000:]
    [output_docx] = output_docx_parent.glob("*/*.docx")
    assert output_docx.parent.stem == output_docx.stem

    roundtripped = _run_cli(
        ["convert", str(output_docx), "--to", "md", "--output-dir", str(output_directory)],
        runtime_root=tmp_path,
    )
    assert roundtripped.returncode == 0, roundtripped.stderr[-4000:]
    assert output_directory.is_dir()
    document_nodes = [path for path in output_directory.iterdir() if path.is_dir()]
    assert len(document_nodes) == 1
    assert (document_nodes[0] / "docwen-node.json").is_file()
    assert len(list(document_nodes[0].glob("*.md"))) == 1

from __future__ import annotations

import hashlib
import subprocess
import sys
from pathlib import Path

import pytest

pytestmark = [pytest.mark.contract, pytest.mark.golden]


PROJECT_ROOT = Path(__file__).resolve().parents[2]


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()


def test_fa11_probe_uses_isolated_project_workers_and_file_backed_dictionaries() -> None:
    probe = (PROJECT_ROOT / "tools" / "validation" / "probe_proofread_numbering_parity.py").read_text(encoding="utf-8")
    for project in ("docwen-ref-tk", "docwen-ref-pyside6", "docwen-current"):
        assert project in probe
    assert "proofread_typos.toml" in probe
    assert '"proofread" / "typos.toml"' in probe
    assert "subprocess.run" in probe
    assert "shutil.copytree" in probe
    assert "_augment_anchor_evidence" in probe
    assert "get_config_loader" not in probe
    assert "create_runtime_port(config_loader=loader)" in probe


def test_fa11_probe_entrypoint_is_executable() -> None:
    completed = subprocess.run(
        [
            sys.executable,
            str(PROJECT_ROOT / "tools" / "validation" / "probe_proofread_numbering_parity.py"),
            "--help",
        ],
        cwd=PROJECT_ROOT,
        check=False,
        capture_output=True,
        text=True,
        timeout=30,
    )

    assert completed.returncode == 0, completed.stderr
    assert "docwen-current" in completed.stdout

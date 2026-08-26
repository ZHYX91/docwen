"""Release-gate source Machine proof for the canonical physical-page corpus."""

from __future__ import annotations

import sys
from pathlib import Path

import pytest
from scripts.release.verify_packaged_cli import _run_machine_protocol_smoke
from tests.support.cli import bundle_cli_command

pytestmark = [
    pytest.mark.e2e,
    pytest.mark.release_gate,
    pytest.mark.slow,
    pytest.mark.skipif(sys.platform == "darwin", reason="Physical-page source capabilities are unavailable on macOS"),
]


def test_source_machine_runs_the_canonical_physical_page_release_corpus(tmp_path: Path) -> None:
    """Run the same P=4/K=5, four-combination verifier used for frozen candidates."""

    work_dir = tmp_path / "work"
    work_dir.mkdir()
    (work_dir / "config_home").mkdir()
    (work_dir / "log_home").mkdir()

    output = _run_machine_protocol_smoke(bundle_cli_command(), work_dir=work_dir)

    assert output.is_file()

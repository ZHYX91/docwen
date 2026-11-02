"""Exercise the package-gate harness against a source CLI, without claiming package evidence."""

from __future__ import annotations

import sys
from pathlib import Path

import pytest
from scripts.release import conversion_fidelity_smoke as smoke
from tools.source_imports import source_import_environment

pytestmark = pytest.mark.integration


def test_fidelity_harness_on_current_source_cli(tmp_path, monkeypatch):
    repo = Path(__file__).resolve().parents[2]
    original = smoke.subprocess.Popen

    def spawn(command, **kwargs):
        # Only this source self-test substitutes an entry point. The installed
        # acceptance script itself accepts exactly one native packaged binary.
        assert command[0] == sys.executable
        kwargs["env"] = source_import_environment(repo, kwargs["env"])
        return original([sys.executable, "-m", "docwen_bundle.cli_entry", *command[1:]], **kwargs)

    monkeypatch.setattr(smoke.subprocess, "Popen", spawn)
    report = smoke.run_smoke(
        Path(sys.executable), tmp_path / "source-harness", candidate_id="source-selftest-NOT-package"
    )
    assert report["status"] == "passed"
    assert len(report["cases"]) == 9

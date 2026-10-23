"""Release-gate source Machine proof for the canonical physical-page corpus."""

from __future__ import annotations

import json
import sys
from pathlib import Path
from textwrap import dedent

import pytest
from scripts.release.verify_packaged_cli import _run_machine_protocol_smoke

pytestmark = [
    pytest.mark.e2e,
    pytest.mark.release_gate,
    pytest.mark.slow,
]


def test_source_machine_runs_the_canonical_physical_page_release_corpus(tmp_path: Path) -> None:
    """Run the same P=4/K=5, four-combination verifier used for frozen candidates."""

    work_dir = tmp_path / "work"
    work_dir.mkdir()
    (work_dir / "config_home").mkdir()
    (work_dir / "log_home").mkdir()

    # A terminal event is emitted before TaskManager's finally cleans its
    # workspace. Exercise one deliberately delayed real PDF cleanup without
    # substituting conversion or Bundle output. Machine completion must still
    # leave the next task's workspace precondition satisfied.
    command = [
        sys.executable,
        "-c",
        dedent("""
            import json
            import sys
            import time
            from pathlib import Path
            from docwen_runtime.workspace.manager import WorkspaceManager

            original_cleanup = WorkspaceManager.cleanup
            delayed = False

            def delayed_cleanup(self, task_id):
                global delayed
                handle = self.get(task_id)
                if not delayed and handle is not None and handle.input_path.endswith('.pdf'):
                    delayed = True
                    Path('cleanup-delay.json').write_text(json.dumps({'delaySeconds': 1}), encoding='utf-8')
                    time.sleep(1)
                return original_cleanup(self, task_id)

            WorkspaceManager.cleanup = delayed_cleanup
            from docwen_bundle.cli_entry import main
            raise SystemExit(main(sys.argv[1:]))
        """),
    ]
    output = _run_machine_protocol_smoke(command, work_dir=work_dir)

    assert output.is_file()
    assert json.loads((work_dir / "cleanup-delay.json").read_text(encoding="utf-8")) == {"delaySeconds": 1}

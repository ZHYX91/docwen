"""Launch source GUI/CLI with a leased runtime and explicit checkout imports."""

from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

if __package__ in {None, ""}:
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from tools.run_lease import lease_payload, transition
from tools.workspace_root import resolve_workspace_root


def source_environment(repo: Path, run: Path) -> dict[str, str]:
    environment = os.environ.copy()
    sources = sorted(path for path in (repo / "packages").rglob("src") if path.is_dir())
    environment.update(
        PYTHONPATH=os.pathsep.join(map(str, sources)),
        PYTHONDONTWRITEBYTECODE="1",
        DOCWEN_RUNTIME_ROOT=str(run / "runtime"),
        TEMP=str(run / "temp"),
        TMP=str(run / "temp"),
        TMPDIR=str(run / "temp"),
    )
    return environment


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--workspace-root", type=Path)
    parser.add_argument("entry", choices=("gui", "cli"))
    parser.add_argument("arguments", nargs=argparse.REMAINDER)
    args = parser.parse_args(argv)
    repo = Path(__file__).resolve().parents[1]
    workspace = resolve_workspace_root(repo, explicit=args.workspace_root)
    run = Path(tempfile.mkdtemp(prefix="source-", dir=workspace / "temp")).resolve()
    (run / "temp").mkdir()
    lease = lease_payload(run, owner="docwen.tools.run-source", kind="source-launch")
    marker = run / ".docwen-temp-lease.json"
    marker.write_text(json.dumps(lease), encoding="utf-8")
    state = "retained-failure"
    try:
        arguments = args.arguments[1:] if args.arguments[:1] == ["--"] else args.arguments
        module = "docwen_bundle.gui_entry" if args.entry == "gui" else "docwen_bundle.cli_entry"
        result = subprocess.run(
            [sys.executable, "-m", module, *arguments], cwd=repo, env=source_environment(repo, run), check=False
        )
        state = "completed-success" if result.returncode == 0 else "retained-failure"
        return result.returncode
    except KeyboardInterrupt:
        state = "retained-interrupted"
        raise
    finally:
        transition(lease, root=run, owner="docwen.tools.run-source", state=state)
        marker.write_text(json.dumps(lease), encoding="utf-8")
        if state == "completed-success":
            try:
                shutil.rmtree(run)
            except OSError:
                transition(lease, root=run, owner="docwen.tools.run-source", state="retained-cleanup-failure")
                marker.write_text(json.dumps(lease), encoding="utf-8")
                raise


if __name__ == "__main__":
    raise SystemExit(main())

"""Launch source GUI/CLI with a leased runtime and explicit checkout imports."""

from __future__ import annotations

import argparse
import os
import subprocess
import sys
from pathlib import Path

if __package__ in {None, ""}:
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from tools.run_lease import managed_run
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
    with managed_run(
        workspace / "temp", prefix="source-", owner="docwen.tools.run-source", kind="source-launch"
    ) as run:
        (run.root / "temp").mkdir()
        arguments = args.arguments[1:] if args.arguments[:1] == ["--"] else args.arguments
        module = "docwen_bundle.gui_entry" if args.entry == "gui" else "docwen_bundle.cli_entry"
        result = subprocess.run(
            [sys.executable, "-m", module, *arguments], cwd=repo, env=source_environment(repo, run.root), check=False
        )
        run.state = "completed-success" if result.returncode == 0 else "retained-failure"
        return result.returncode


if __name__ == "__main__":
    raise SystemExit(main())

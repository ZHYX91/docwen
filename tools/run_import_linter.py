from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path

_IMPORT_LINTER_BOOTSTRAP = (
    "from importlinter.cli import lint_imports_command; lint_imports_command(prog_name='lint-imports')"
)


def _workspace_src_paths(repo_root: Path) -> list[str]:
    paths = [str(repo_root / "src")]
    for src_dir in sorted((repo_root / "packages").glob("**/src")):
        if src_dir.is_dir():
            paths.append(str(src_dir))
    return paths


def main(argv: list[str]) -> int:
    repo_root = Path(__file__).resolve().parents[1]
    env = os.environ.copy()
    extra_paths = _workspace_src_paths(repo_root)
    existing = env.get("PYTHONPATH", "").strip()
    env["PYTHONPATH"] = os.pathsep.join([*extra_paths, existing] if existing else extra_paths)
    proc = subprocess.run(
        [sys.executable, "-c", _IMPORT_LINTER_BOOTSTRAP, *argv],
        cwd=repo_root,
        env=env,
    )
    return int(proc.returncode)


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))

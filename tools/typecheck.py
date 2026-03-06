from __future__ import annotations

import subprocess
import sys
from pathlib import Path


def _run(args: list[str]) -> int:
    proc = subprocess.run(args, cwd=Path(__file__).resolve().parents[1])
    return int(proc.returncode)


def main() -> int:
    steps: list[tuple[list[str], bool]] = [
        ([sys.executable, "-m", "pyright", "--level", "error"], True),
    ]

    exit_code = 0
    for args, gate in steps:
        code = _run(args)
        if code != 0:
            if gate:
                return code
            exit_code = code

    return exit_code


if __name__ == "__main__":
    raise SystemExit(main())

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

if __package__ in {None, ""}:
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from tools.source_checks import typecheck_steps


def _run(args: list[str]) -> int:
    proc = subprocess.run(args, cwd=Path(__file__).resolve().parents[1])
    return int(proc.returncode)


def main() -> int:
    exit_code = 0
    for _name, args, gate in typecheck_steps():
        code = _run(args)
        if code != 0:
            if gate:
                return code
            exit_code = code

    return exit_code


if __name__ == "__main__":
    raise SystemExit(main())

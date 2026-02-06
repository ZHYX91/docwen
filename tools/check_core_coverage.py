from __future__ import annotations

import argparse
import subprocess
import sys


CORE_INCLUDE = ",".join(
    [
        "*/src/docwen/cli/executor.py",
        "*/src/docwen/services/strategies/__init__.py",
        "*/src/docwen/converter/smart_converter.py",
        "*/src/docwen/table_merger/core.py",
    ]
)


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--fail-under", type=float, default=40.0)
    args, pytest_args = parser.parse_known_args(argv)

    steps = [
        [sys.executable, "-m", "coverage", "erase"],
        [sys.executable, "-m", "coverage", "run", "-m", "pytest", *pytest_args],
        [
            sys.executable,
            "-m",
            "coverage",
            "report",
            f"--include={CORE_INCLUDE}",
            f"--fail-under={args.fail_under}",
        ],
        [sys.executable, "-m", "coverage", "html"],
    ]

    for cmd in steps:
        completed = subprocess.run(cmd)
        if completed.returncode != 0:
            return completed.returncode

    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))


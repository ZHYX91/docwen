from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path


def _run(args: list[str]) -> int:
    proc = subprocess.run(args, cwd=Path(__file__).resolve().parents[1])
    return int(proc.returncode)


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--suite", choices=["fast", "full"], default="fast")
    parser.add_argument("--phase5", action="store_true")
    parser.add_argument("--skip-ruff", action="store_true")
    parser.add_argument("--skip-pyright", action="store_true")
    parser.add_argument("--skip-pytest", action="store_true")
    args = parser.parse_args(argv)

    steps: list[tuple[str, list[str], bool]] = []
    if not args.skip_ruff:
        steps += [
            ("ruff-format", [sys.executable, "-m", "ruff", "format", "--check", "."], True),
            ("ruff-check", [sys.executable, "-m", "ruff", "check", "."], True),
        ]
        if args.phase5:
            steps += [
                (
                    "ruff-phase5",
                    [
                        sys.executable,
                        "-m",
                        "ruff",
                        "check",
                        ".",
                        "--select",
                        "PTH,T20",
                        "--statistics",
                        "--exit-zero",
                    ],
                    False,
                )
            ]
    if not args.skip_pyright:
        steps += [
            ("pyright", [sys.executable, "-m", "pyright", "--level", "error"], True),
        ]

    exit_code = 0
    for name, cmd, gate in steps:
        print(f"==> {name}")
        code = _run(cmd)
        if code != 0:
            if gate:
                return code
            exit_code = code

    if not args.skip_pytest:
        print("==> pytest")
        pytest_cmd: list[str] = [sys.executable, "-m", "pytest"]
        if args.suite == "fast":
            pytest_cmd += ["-m", "not integration and not slow"]
        return_code = _run(pytest_cmd)
        if return_code != 0:
            return return_code
    return exit_code


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))

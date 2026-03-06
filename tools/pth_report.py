from __future__ import annotations

import argparse
import json
import subprocess
import sys
from collections import Counter
from pathlib import Path


def _iter_dir_buckets(path: str, *, depth: int) -> str:
    parts = Path(path).as_posix().split("/")
    return "/".join(parts[:depth]) if len(parts) >= depth else parts[0]


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--root", default="src/docwen")
    parser.add_argument("--top-files", type=int, default=30)
    parser.add_argument("--top-dirs", type=int, default=30)
    parser.add_argument("--dir-depth", type=int, default=4)
    args = parser.parse_args(argv)

    proc = subprocess.run(
        [
            sys.executable,
            "-m",
            "ruff",
            "check",
            args.root,
            "--select",
            "PTH",
            "--exit-zero",
            "--output-format",
            "json",
        ],
        capture_output=True,
        text=True,
    )
    data = json.loads(proc.stdout or "[]")

    files = Counter()
    dirs = Counter()
    codes = Counter()

    for item in data:
        filename = item.get("filename") or item.get("path") or item.get("location", {}).get("path")
        code = item.get("code")
        if filename:
            files[filename] += 1
            dirs[_iter_dir_buckets(filename, depth=args.dir_depth)] += 1
        if code:
            codes[code] += 1

    total = len(data)
    print(f"root={args.root} total={total}")
    print("")
    print("Top codes:")
    for code, n in codes.most_common(20):
        print(f"{n:5} {code}")
    print("")
    print("Top dirs:")
    for d, n in dirs.most_common(args.top_dirs):
        print(f"{n:5} {d}")
    print("")
    print("Top files:")
    for f, n in files.most_common(args.top_files):
        print(f"{n:5} {f}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))

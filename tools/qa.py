from __future__ import annotations

import argparse
import ast
import re
import subprocess
import sys
from pathlib import Path


def _run(args: list[str]) -> int:
    proc = subprocess.run(args, cwd=Path(__file__).resolve().parents[1])
    return int(proc.returncode)


def _scan_private_symbol_usage(repo_root: Path) -> int:
    deny_import = re.compile(r"^\s*from\s+docwen(?:\.[\w]+)*\s+import\s+.*\b_\w+", re.MULTILINE)
    deny_string_path = re.compile(r"docwen(?:\.[\w]+)*\._[A-Za-z]\w*")

    def _chain(node: ast.AST) -> list[str] | None:
        if isinstance(node, ast.Name):
            return [node.id]
        if isinstance(node, ast.Attribute):
            parent = _chain(node.value)
            if not parent:
                return None
            return [*parent, node.attr]
        return None

    def _scan_ast_for_docwen_private_access(text: str) -> list[tuple[int, str]]:
        try:
            tree = ast.parse(text)
        except SyntaxError:
            return []

        aliases: dict[str, str] = {}
        issues: list[tuple[int, str]] = []

        for node in ast.walk(tree):
            if isinstance(node, ast.Import):
                for item in node.names:
                    if (item.name == "docwen" or item.name.startswith("docwen.")) and item.asname:
                        aliases[item.asname] = item.name
            elif isinstance(node, ast.ImportFrom):
                mod = node.module or ""
                if not (mod == "docwen" or mod.startswith("docwen.")):
                    continue
                for item in node.names:
                    if item.name == "*":
                        continue
                    if item.name.startswith("_") and not item.name.startswith("__"):
                        issues.append((getattr(node, "lineno", 1), "import-private-from-docwen"))
                        continue
                    bound = item.asname or item.name
                    aliases[bound] = f"{mod}.{item.name}"
            elif isinstance(node, ast.Attribute):
                if not (node.attr.startswith("_") and not node.attr.startswith("__")):
                    continue

                dotted = _chain(node)
                if not dotted:
                    continue
                root = dotted[0]
                if root == "docwen":
                    issues.append((getattr(node, "lineno", 1), "attr-private-on-docwen"))
                    continue
                if root in aliases:
                    issues.append((getattr(node, "lineno", 1), "attr-private-on-docwen-import"))
                    continue

        return issues

    offenders: dict[Path, set[str]] = {}
    for rel_root in ("tests", "tools"):
        base = repo_root / rel_root
        if not base.is_dir():
            continue
        for path in base.rglob("*.py"):
            if path.name == "qa.py":
                continue
            try:
                text = path.read_text(encoding="utf-8")
            except Exception:
                continue

            if deny_import.search(text):
                offenders.setdefault(path, set()).add("import-private-from-docwen")
            if deny_string_path.search(text):
                offenders.setdefault(path, set()).add("string-refers-docwen-private")

            for lineno, kind in _scan_ast_for_docwen_private_access(text):
                offenders.setdefault(path, set()).add(f"{kind}:L{lineno}")

    if not offenders:
        return 0

    print("==> private-symbol-boundary")
    print("Found disallowed private symbol usage in tests/tools. Fix by adding a public API in src and importing that.")
    for path, reasons in sorted(offenders.items(), key=lambda x: str(x[0])):
        rel = path.relative_to(repo_root)
        print(f"  - {rel} ({','.join(sorted(reasons))})")
    return 1


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--suite", choices=["fast", "full"], default="fast")
    parser.add_argument("--phase5", action="store_true")
    parser.add_argument("--skip-ruff", action="store_true")
    parser.add_argument("--skip-pyright", action="store_true")
    parser.add_argument("--skip-pytest", action="store_true")
    args = parser.parse_args(argv)

    repo_root = Path(__file__).resolve().parents[1]
    code = _scan_private_symbol_usage(repo_root)
    if code != 0:
        return code

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

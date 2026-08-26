#!/usr/bin/env python3
"""Architecture cleanliness scanner — detects violations in the repo.

Scans source, configuration, workflow, and current Markdown documentation files
in the repository for architecture rule violations. Generated output and
archived documentation are excluded.

Categories:
  1-old-project-path      – References to the old project root (docwen旧).
  2-deprecated-deep-import – Deprecated deep imports from the legacy monolith
                            (from docwen.converter.*).
  3-monkey-patch          – Runtime monkey-patching of third-party libraries.
  4-shim-wrapper          – Shim / wrapper / compat / forward / alias adapter
                            files that indicate incomplete migration.
  5-cross-package-import   – One plugin importing from another plugin's
                            internal modules (detected via docwen_plugin_X
                            import in a different docwen_plugin_Y source dir).
  6-legacy-monolith-import – packages/ code importing from legacy monolith
                            namespace (docwen.application, docwen.cli, …).

Output:
  One violation per line:  "[category] file:line: content"

Exit codes:
  0 – No violations.
  1 – One or more violations found.
"""

from __future__ import annotations

import argparse
import ast
import re
import sys
from pathlib import Path

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

EXCLUDE_DIRS = frozenset(
    {
        ".git",
        ".venv",
        "__pycache__",
        "htmlcov",
        "build",
        "dist",
        "tmp",
        ".mypy_cache",
        ".pytest_cache",
        ".ruff_cache",
        "node_modules",
    }
)

# Full path prefixes (relative to repo root) to exclude.
# e.g. "tools/generated" excludes tools/generated/ and everything under it.
EXCLUDE_PATH_PREFIXES: tuple[str, ...] = (
    "docs/archive/",
    "tests/fixtures/golden/",
)

# Files whose own content explains what they detect and should not trigger
# self-detection.
_SKIP_FILES = frozenset(
    {
        "tools/validation/check_architecture_cleanliness.py",
    }
)

SCAN_EXTENSIONS = frozenset({".md", ".py", ".toml", ".yml", ".yaml"})

ERROR_CATEGORIES = frozenset(
    {
        "1-old-project-path",
        "2-deprecated-deep-import",
        "3-monkey-patch",
        "4-shim-wrapper",
        "5-cross-package-import",
        "6-legacy-monolith-import",
    }
)

# Legacy monolith namespaces that packages/ must not import.
_LEGACY_NS = (
    "docwen.application",
    "docwen.bootstrap",
    "docwen.cli",
    "docwen.config",
    "docwen.converter",
    "docwen.core",
    "docwen.docx_spell",
    "docwen.errors",
    "docwen.formats",
    "docwen.gui",
    "docwen.i18n",
    "docwen.ipc",
    "docwen.md_spell",
    "docwen.security",
    "docwen.services",
    "docwen.template",
    "docwen.text_rules",
    "docwen.utils",
)

# Patterns
_RE_DOCWEN_OLD = re.compile(r"docwen旧|OneDrive[/\\]Projects[/\\]docwen旧")

_RE_DEEP_IMPORT_DEPRECATED = re.compile(r"(?:from\s+docwen\.converter(?:\.|\s+import)|import\s+docwen\.converter\.)")

# Monkey-patch detection: only flag module-level (non-test) code that
# patches third-party libraries at runtime.  pytest's ``monkeypatch``
# fixture is legitimate test infrastructure — not a violation.
_RE_MONKEY_PATCH_DOCSTRING = re.compile(
    r'"""Runtime monkey-patches? for',
    re.IGNORECASE,
)

_APPROVED_RUNTIME_PATCH_FILES = frozenset({"packages/core/src/docwen_core/ofd.py"})
_APPROVED_LEGACY_PARITY_FILES = frozenset(
    {
        "tools/golden_parity_runner.py",
        "tools/validation/probe_merge_tables_parity.py",
    }
)

_SHIMMY_NAMES = ("shim", "wrapper", "compat", "forward", "alias")
_SHIMMY_PATTERN = re.compile(
    r"(?:^|[\\/])(?:" + "|".join(_SHIMMY_NAMES) + r")(?:[_.]|$)",
    re.IGNORECASE,
)
_SHIMMY_CONTENT_PATTERN = re.compile(
    r"\b(?:compatibility|legacy)\s+(?:shim|wrapper)|\bforwarding\s+(?:module|wrapper)",
    re.IGNORECASE,
)

_CROSS_PLUGIN_RE = re.compile(r"(?:from|import)\s+(docwen_plugin_\w+)")

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _should_skip(rel_path: str) -> bool:
    """Return True if the relative path should be excluded."""
    parts = tuple(rel_path.replace("\\", "/").split("/"))
    # Exclude hidden / build directories by component name.
    if EXCLUDE_DIRS.intersection(parts):
        return True
    # Exclude by path prefix.
    normalized = "/".join(parts) + "/"
    return any(normalized.startswith(prefix) for prefix in EXCLUDE_PATH_PREFIXES)


def _collect_files(repo_root: Path) -> list[Path]:
    """Walk the repo and collect all scannable files."""
    files: list[Path] = []
    for ext in SCAN_EXTENSIONS:
        for path in repo_root.rglob(f"*{ext}"):
            rel = str(path.relative_to(repo_root)).replace("\\", "/")
            if _should_skip(rel):
                continue
            if rel in _SKIP_FILES:
                continue
            files.append(path)
    return sorted(files)


def _determine_plugin(file_path: Path, repo_root: Path) -> str | None:
    """If *file_path* lives inside a plugin src tree, return its package name."""
    parts = file_path.relative_to(repo_root).parts
    if parts[0] == "packages" and len(parts) >= 4 and parts[1] == "plugins":
        # packages/plugins/<name>/src/docwen_plugin_<name>/...
        plugin_src = Path(repo_root, *parts[:4])
        if plugin_src.is_dir():
            # The plugin package name is the directory under src/
            src_dir = Path(repo_root, *parts[:3], "src")
            if src_dir.is_dir():
                for child in src_dir.iterdir():
                    if child.is_dir() and child.name.startswith("docwen_plugin_"):
                        return child.name
    return None


# ---------------------------------------------------------------------------
# Scanners
# ---------------------------------------------------------------------------


def _scan_old_project_path(content: str, file_path: Path, repo_root: Path) -> list[str]:
    """Scan for references to the old project root (docwen旧)."""
    violations: list[str] = []
    rel_str = str(file_path.relative_to(repo_root)).replace("\\", "/")
    for match in _RE_DOCWEN_OLD.finditer(content):
        line_no = content.count("\n", 0, match.start()) + 1
        violations.append(f"[1-old-project-path] {rel_str}:{line_no}: {match.group()}")
    return violations


def _scan_deprecated_deep_import(content: str, file_path: Path, repo_root: Path) -> list[str]:
    """Scan for deprecated deep imports from docwen.converter.*."""
    violations: list[str] = []
    rel = file_path.relative_to(repo_root)
    rel_str = rel.as_posix()
    if rel_str in _APPROVED_LEGACY_PARITY_FILES:
        return []
    for line_no, line in enumerate(content.splitlines(), 1):
        if _RE_DEEP_IMPORT_DEPRECATED.search(line):
            # Skip comments and docstrings that might reference these
            stripped = line.strip()
            if stripped.startswith(("#", '"""', "'''")):
                continue
            violations.append(f"[2-deprecated-deep-import] {rel_str}:{line_no}: {stripped[:120]}")
    return violations


def _scan_monkey_patch(content: str, file_path: Path, repo_root: Path) -> list[str]:
    """Scan for runtime monkey-patching of third-party libraries.

    We only flag production source code (non-test) that patches external
    libraries at import time.  pytest's ``monkeypatch`` fixture usage in
    test files is normal and not flagged.
    """
    violations: list[str] = []
    rel = file_path.relative_to(repo_root)

    # Skip test files — pytest monkeypatch fixture is legitimate.
    rel_str = rel.as_posix()
    if "/tests/" in rel_str or "/test_" in rel_str or "test_" in file_path.stem:
        return []

    # Runtime patching is governed only for production package code.
    if file_path.suffix != ".py" or not rel_str.startswith("packages/"):
        return []
    if rel_str in _APPROVED_RUNTIME_PATCH_FILES:
        return []

    # An explicit patching module is always actionable unless allowlisted.
    if _RE_MONKEY_PATCH_DOCSTRING.search(content):
        line_no = content[: content.index('"""Runtime monkey-patches')].count("\n") + 1
        violations.append(f"[3-monkey-patch] {rel_str}:{line_no}: runtime monkey-patch docstring detected")
        return violations

    try:
        tree = ast.parse(content, filename=str(file_path))
    except SyntaxError:
        return []

    imported_bindings: dict[str, str] = {}
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            for alias in node.names:
                imported_bindings[alias.asname or alias.name.split(".", 1)[0]] = alias.name.split(".", 1)[0]
        elif isinstance(node, ast.ImportFrom):
            module_root = str(node.module or "").split(".", 1)[0]
            for alias in node.names:
                if alias.name != "*":
                    imported_bindings[alias.asname or alias.name] = module_root

    for node in ast.walk(tree):
        targets: list[ast.expr] = []
        if isinstance(node, ast.Assign):
            targets = list(node.targets)
        elif isinstance(node, ast.AnnAssign):
            targets = [node.target]
        for target in targets:
            root = target
            while isinstance(root, ast.Attribute):
                root = root.value
            module_root = imported_bindings.get(root.id) if isinstance(root, ast.Name) else None
            if (
                isinstance(root, ast.Name)
                and isinstance(target, ast.Attribute)
                and module_root
                and module_root not in sys.stdlib_module_names
                and not module_root.startswith("docwen")
            ):
                violations.append(
                    f"[3-monkey-patch] {rel_str}:{node.lineno}: assignment mutates imported object {root.id}"
                )

    return violations


def _scan_shim_wrapper(content: str, file_path: Path, repo_root: Path) -> list[str]:
    """Scan for shim/wrapper/compat/forward/alias adapter files."""
    violations: list[str] = []
    rel = file_path.relative_to(repo_root)
    rel_str = rel.as_posix()
    # Only flag files inside packages/, tools/, scripts/ — not in tests/
    if rel.parts[0] not in ("packages", "tools", "scripts"):
        return []
    # Exclude test files
    if "test" in file_path.stem.lower() or "tests" in rel.parts:
        return []
    module_docstring = ""
    if file_path.suffix == ".py":
        try:
            module_docstring = ast.get_docstring(ast.parse(content, filename=str(file_path))) or ""
        except SyntaxError:
            module_docstring = ""
    if _SHIMMY_PATTERN.search(file_path.stem) or _SHIMMY_CONTENT_PATTERN.search(module_docstring):
        violations.append(f"[4-shim-wrapper] {rel_str}:1: compatibility shim/wrapper content or file name detected")
    return violations


def _scan_cross_package_import(content: str, file_path: Path, repo_root: Path) -> list[str]:
    """Scan for cross-plugin imports (plugin X importing from plugin Y)."""
    violations: list[str] = []
    rel = file_path.relative_to(repo_root)
    rel_str = rel.as_posix()
    current_plugin = _determine_plugin(file_path, repo_root)
    if current_plugin is None:
        return []

    for line_no, line in enumerate(content.splitlines(), 1):
        stripped = line.strip()
        if stripped.startswith("#"):
            continue
        for match in _CROSS_PLUGIN_RE.finditer(stripped):
            imported_plugin = match.group(1)
            if imported_plugin != current_plugin:
                violations.append(
                    f"[5-cross-package-import] {rel_str}:{line_no}: {current_plugin} imports {imported_plugin}"
                )
    return violations


def _scan_legacy_monolith_import(content: str, file_path: Path, repo_root: Path) -> list[str]:
    """Scan for legacy monolith namespace imports inside packages/."""
    violations: list[str] = []
    rel = file_path.relative_to(repo_root)
    rel_str = rel.as_posix()

    # Only scan files under packages/
    if rel.parts[0] != "packages":
        return []

    for line_no, line in enumerate(content.splitlines(), 1):
        stripped = line.strip()
        if stripped.startswith("#"):
            continue
        for ns in _LEGACY_NS:
            # Match 'from docwen.xxx' or 'import docwen.xxx'
            if f"from {ns}" in stripped or f"import {ns}" in stripped:
                violations.append(
                    f"[6-legacy-monolith-import] {rel_str}:{line_no}: "
                    f"packages code imports legacy monolith module: {stripped[:120]}"
                )
                break  # One violation per line is enough
    return violations


# ---------------------------------------------------------------------------
# Orchestration
# ---------------------------------------------------------------------------

SCANNERS = [
    (_scan_old_project_path, True),  # uses content
    (_scan_deprecated_deep_import, True),  # uses content
    (_scan_monkey_patch, True),  # uses content
    (_scan_shim_wrapper, True),
    (_scan_cross_package_import, True),  # uses content
    (_scan_legacy_monolith_import, True),  # uses content
]


def scan(repo_root: Path) -> list[str]:
    """Run all scanners and return sorted violation lines."""
    files = _collect_files(repo_root)
    all_violations: list[str] = []

    for file_path in files:
        # Read content once for content-based scanners
        content: str | None = None
        for scanner_fn, needs_content in SCANNERS:
            if needs_content:
                if content is None:
                    try:
                        content = file_path.read_text(encoding="utf-8")
                    except (UnicodeDecodeError, OSError):
                        break
                violations = scanner_fn(content, file_path, repo_root)
            else:
                violations = scanner_fn(file_path, repo_root)
            all_violations.extend(violations)

    return sorted(all_violations)


def _has_error(violations: list[str]) -> bool:
    """Check whether any violation is an error-level category."""
    for v in violations:
        cat = v.split("]", 1)[0].lstrip("[")
        if cat in ERROR_CATEGORIES:
            return True
    return False


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Architecture cleanliness scanner")
    parser.add_argument(
        "--repo-root",
        type=Path,
        default=Path(__file__).resolve().parents[2],
        help="Repository root (default: auto-detect from script location)",
    )
    args = parser.parse_args(argv)

    repo_root = args.repo_root.resolve()
    if not repo_root.is_dir():
        print(f"error: not a directory: {repo_root}", file=sys.stderr)
        return 2

    violations = scan(repo_root)

    if violations:
        for v in violations:
            print(v)

    has_err = _has_error(violations)
    if has_err:
        summary_parts = []
        for cat in sorted(ERROR_CATEGORIES):
            count = sum(1 for v in violations if v.startswith(f"[{cat}]"))
            if count:
                summary_parts.append(f"{cat}={count}")
        print(
            f"\n==> architecture cleanliness: {len(violations)} violation(s) ({' | '.join(summary_parts)})",
            file=sys.stderr,
        )
        return 1

    print("==> architecture cleanliness: ok", file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

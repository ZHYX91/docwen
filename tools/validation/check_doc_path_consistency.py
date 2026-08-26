#!/usr/bin/env python3
"""Doc-path consistency checker — gate item 6 of 架构收口方案 (核验版).

Enforces the documentation path-reference rule defined in
``docs/testing.md`` §4.1 (历史路径一致性门禁):

  1. Current-architecture docs must NOT describe deleted old-path tokens
     (``src/docwen/``, ``gui_tk``, ``combobox_adapter``, ``theme_styles``)
     as if they are the current structure.
  2. Historical / migration / archive docs MAY keep old paths, provided the
     document carries a historical-context label.
  3. The checker must not false-positive on whitelisted mapping docs.

Exemption model (a file is exempt — never flagged — if ANY of these hold):
  - It lives under ``docs/archive/`` (archived).
  - It is a structurally-historical doc: ``docs/CHANGELOG.md`` or
    ``docs/superpowers/**``.
  - It is one of the named migration-mapping docs, or the rule-defining
    spec itself (which necessarily mentions the tokens).
  - It carries a historical-context marker anywhere in its text
    (已删除 / 已移除 / 已归档 / OBSOLETE / 历史对照 / 迁移来源 / 迁移 /
    旧路径 / 旧 `src/docwen/` / "no longer" / "deprecated").

The marker-based exemption is deliberately file-granular and conservative:
the rule's intent is that a *document as a whole* identify itself as
historical, not that every individual line be annotated.

Exit codes:
  0 – No unlabeled old-path references in current-architecture docs.
  1 – One or more current-architecture docs reference old paths without a
      historical-context label.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

# Deleted / old-structure tokens that must not appear as current in
# current-architecture docs. Sourced from the §4.1 rule.
OLD_PATH_TOKENS: tuple[str, ...] = (
    "src/docwen/",
    "gui_tk",
    "combobox_adapter",
    "theme_styles",
)

# Structurally-exempt locations (historical by nature of what they are).
_EXEMPT_SUBPATHS: tuple[str, ...] = (
    "docs/archive/",
    "docs/superpowers/",
)

# Specifically-exempt files (migration-mapping docs + the rule-defining spec).
_EXEMPT_FILES: frozenset[str] = frozenset(
    {
        "docs/CHANGELOG.md",
        "docs/packaging.md",
        "docs/specs/gui-behavior.md",
        "docs/configuration.md",
        "docs/testing.md",
    }
)

# A file is considered "labeled historical" if any of these markers appears
# in its text. Covers Chinese and English historical/migration framing.
_HISTORICAL_MARKERS: tuple[str, ...] = (
    "已删除",
    "已移除",
    "已归档",
    "OBSOLETE",
    "obsolete",
    "历史对照",
    "迁移来源",
    "迁移",
    "归档",
    "旧路径",
    "旧 `src/docwen/",
    "旧src/docwen/",
    "历史核对",
    "no longer",
    "deprecated",
    "removed",
    "migrated",
)

_SCAN_EXTENSIONS = frozenset({".md"})


# ---------------------------------------------------------------------------
# Core logic
# ---------------------------------------------------------------------------


def _is_exempt_path(rel_path: str) -> bool:
    """True if the file's location is structurally historical/exempt."""
    if rel_path in _EXEMPT_FILES:
        return True
    normalized = rel_path.replace("\\", "/")
    return any(normalized.startswith(p) for p in _EXEMPT_SUBPATHS)


def _has_historical_marker(text: str) -> bool:
    """True if the document text carries a historical-context label."""
    return any(marker in text for marker in _HISTORICAL_MARKERS)


def find_violations(repo_root: Path) -> list[str]:
    """Return a list of violation strings for unlabeled old-path references.

    A violation is reported per file (not per line): a current-architecture
    doc that contains any old-path token and lacks any historical-context
    marker.
    """
    docs_root = repo_root / "docs"
    if not docs_root.is_dir():
        return []

    violations: list[str] = []

    for path in sorted(docs_root.rglob("*")):
        if not path.is_file():
            continue
        if path.suffix not in _SCAN_EXTENSIONS:
            continue

        rel_path = path.relative_to(repo_root).as_posix()

        if _is_exempt_path(rel_path):
            continue

        try:
            text = path.read_text(encoding="utf-8")
        except (OSError, UnicodeDecodeError):
            continue

        # Does the doc reference any old-path token at all?
        referenced_tokens = [tok for tok in OLD_PATH_TOKENS if tok in text]
        if not referenced_tokens:
            continue

        # If it does, it must carry a historical-context label.
        if _has_historical_marker(text):
            continue

        violations.append(
            f"[old-path-unlabeled] {rel_path}: references "
            f"{referenced_tokens} without a historical/migration/archive label "
            f"(see 代码门禁与检查规范.md §4.1)"
        )

    return violations


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def main(argv: list[str] | None = None) -> int:
    _doc = __doc__ or ""
    parser = argparse.ArgumentParser(description=_doc.splitlines()[0] if _doc else "")
    parser.add_argument(
        "--repo-root",
        type=Path,
        default=Path(__file__).resolve().parents[2],
        help="Repository root to scan (default: this repo).",
    )
    args = parser.parse_args(argv)

    repo_root: Path = args.repo_root.resolve()
    violations = find_violations(repo_root)

    if violations:
        print(
            "Doc-path consistency check FAILED — current-architecture docs "
            "reference deleted old paths without a historical label:\n",
            file=sys.stdout,
        )
        for v in violations:
            print(v, file=sys.stdout)
        print(
            "\nFix: either remove the old-path reference, or add a "
            "historical/migration/archive label to the doc "
            "(已归档 / 历史对照 / 迁移来源 / 已删除 …).",
            file=sys.stdout,
        )
        return 1

    print("Doc-path consistency check passed — no unlabeled old-path references in current-architecture docs.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

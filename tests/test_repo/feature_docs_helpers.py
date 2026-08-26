"""Shared helpers for feature inventory repository checks."""

from __future__ import annotations

import re
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[2]
FEATURE_DOC = PROJECT_ROOT / "docs" / "specs" / "功能与能力清单.md"
FEATURE_STATUS_COLUMNS = (
    "mapped",
    "implemented",
    "not_started",
    "partial",
    "intentionally_removed",
    "tested",
    "verified",
)
FEATURE_CATEGORY_LABELS = {
    "gui": "GUI",
    "cli": "CLI",
    "conversion": "转换",
    "config": "配置",
    "output": "输出",
    "resource": "资源",
    "arch": "架构",
    "detection": "检测",
}


def markdown_row_cells(line: str) -> list[str]:
    row = line.strip()
    if row.startswith("|"):
        row = row[1:]
    if row.endswith("|"):
        row = row[:-1]
    return [cell.replace(r"\|", "|").strip() for cell in re.split(r"(?<!\\)\|", row)]


def feature_row(feature_id: str) -> list[str]:
    text = FEATURE_DOC.read_text(encoding="utf-8")
    row = next(line for line in text.splitlines() if line.startswith(f"| {feature_id} |"))
    return markdown_row_cells(row)


def feature_rows() -> list[tuple[str, list[str]]]:
    text = FEATURE_DOC.read_text(encoding="utf-8")
    rows: list[tuple[str, list[str]]] = []
    for line in text.splitlines():
        if line.startswith("| FEAT-"):
            cells = markdown_row_cells(line)
            rows.append((cells[0], cells))
    return rows


def repo_python_refs(cell: str) -> list[str]:
    return sorted(
        set(
            re.findall(
                r"(?:(?:packages|tests|scripts|tools|configs|assets|i18n|templates)/[^ `，；、)]+?\.py)(?:::[A-Za-z0-9_]+)?",
                cell,
            )
        )
    )


def repo_path_refs(cell: str) -> list[str]:
    refs: list[str] = []
    for token in sorted(
        set(re.findall(r"(?:(?:packages|src|tests|scripts|tools|configs|assets|i18n|templates)/[^ `，；、)]+)", cell))
    ):
        clean = token.split("::", 1)[0].rstrip(".,;:")
        if clean.endswith((".py", ".toml", ".md", ".json", ".csv", ".png", ".svg", ".ico", ".docx", ".xlsx")) or (
            "/src/" in clean or "/tests/" in clean
        ):
            refs.append(clean)
    return refs


def owner_scoped_roots(owner_cell: str) -> list[Path]:
    roots: list[Path] = []
    static_roots = {
        "docwen_application": PROJECT_ROOT / "packages" / "application" / "src" / "docwen_application",
        "docwen_bundle": PROJECT_ROOT / "packages" / "bundle" / "src" / "docwen_bundle",
        "docwen_cli": PROJECT_ROOT / "packages" / "apps" / "cli" / "src" / "docwen_cli",
        "docwen_core": PROJECT_ROOT / "packages" / "core" / "src" / "docwen_core",
        "docwen_gui": PROJECT_ROOT / "packages" / "apps" / "gui" / "src" / "docwen_gui",
        "docwen_runtime": PROJECT_ROOT / "packages" / "runtime" / "src" / "docwen_runtime",
    }
    for owner, root in static_roots.items():
        if owner in owner_cell:
            roots.append(root)

    for plugin_name in sorted(set(re.findall(r"docwen_plugin_([A-Za-z0-9_]+)", owner_cell))):
        roots.append(PROJECT_ROOT / "packages" / "plugins" / plugin_name / "src" / f"docwen_plugin_{plugin_name}")

    return roots


def owner_scoped_refs(cell: str) -> list[str]:
    refs: set[str] = set()
    path_pattern = r"\b(?:[A-Za-z_][A-Za-z0-9_]*/)+[A-Za-z0-9_./-]+\.py\b"
    for token in re.findall(path_pattern, cell):
        clean = token.split("::", 1)[0].rstrip(".,;:")
        if clean.startswith(
            ("packages/", "tests/", "scripts/", "tools/", "configs/", "assets/", "i18n/", "templates/")
        ):
            continue
        refs.add(clean)

    slash_basenames = {Path(ref).name for ref in refs}
    basename_pattern = r"(?<![/\\])\b[A-Za-z_][A-Za-z0-9_]*\.py\b"
    for token in re.findall(basename_pattern, cell):
        if not token.startswith("test_") and token not in slash_basenames:
            refs.add(token)

    return sorted(refs)


def owner_scoped_ref_exists(relative_ref: str, owner_roots: list[Path]) -> bool:
    for root in owner_roots:
        candidates = [root / relative_ref]
        parts = Path(relative_ref).parts
        if parts and parts[0] in {"application", "bundle", "cli", "core", "gui", "runtime"}:
            candidates.append(root / Path(*parts[1:]))
        for candidate in candidates:
            if candidate.exists():
                return True
        if "/" not in relative_ref:
            matches = [path for path in root.rglob(relative_ref) if path.is_file()]
            if len(matches) == 1:
                return True
    return False

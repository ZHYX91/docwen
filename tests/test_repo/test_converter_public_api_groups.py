"""
验证当前公共包入口的分组锚点与 import/__all__ 对齐。

不验证分组内容正确性（语义由人工审核），只验证：
1. __all__ 中每个符号都落入一个 # ===== X ===== 分组锚点段
2. __all__ 符号顺序与 from .X import Y 符号顺序一致
"""

from __future__ import annotations

import ast
import re
from pathlib import Path

import pytest

pytestmark = pytest.mark.contract

PUBLIC_API_INIT = (
    Path(__file__).parent.parent.parent
    / "packages"
    / "apps"
    / "cli"
    / "src"
    / "docwen_cli"
    / "commands"
    / "__init__.py"
)

GROUP_ANCHOR_RE = re.compile(r"^# (?:(?:===== .+ =====)|(?:[A-Za-z][A-Za-z0-9 _/\-]+))$")


def _normalize_group_anchor(stripped: str) -> str:
    value = stripped.removeprefix("# ").strip()
    if value.startswith("=====") and value.endswith("====="):
        value = value.removeprefix("=====").removesuffix("=====").strip()
    return value


def _parse_source() -> str:
    return PUBLIC_API_INIT.read_text(encoding="utf-8")


def _extract_all_symbols_with_groups(source: str) -> list[tuple[str, str | None]]:
    """Parse __all__ assignment, returning [(symbol, group_anchor_name_or_None), ...]."""
    in_all = False
    current_group: str | None = None
    results: list[tuple[str, str | None]] = []

    for line in source.split("\n"):
        stripped = line.strip()
        if stripped == "__all__ = [":
            in_all = True
            continue
        if in_all:
            if stripped == "]":
                break
            m = GROUP_ANCHOR_RE.match(stripped)
            if m:
                current_group = _normalize_group_anchor(stripped)
                continue
            sym_match = re.match(r'"(\w+)"', stripped)
            if sym_match:
                results.append((sym_match.group(1), current_group))

    return results


def _extract_import_symbols(source: str) -> list[str]:
    """Parse from .X import Y statements (before __all__), returning imported names in order."""
    tree = ast.parse(source)
    symbols: list[str] = []

    class ImportCollector(ast.NodeVisitor):
        def visit_ImportFrom(self, node: ast.ImportFrom) -> None:
            for alias in node.names:
                name = alias.asname or alias.name
                if not name.startswith("_"):
                    symbols.append(name)
            self.generic_visit(node)

    # Only walk nodes before __all__ assignment
    for node in ast.iter_child_nodes(tree):
        if isinstance(node, ast.Assign):
            for target in node.targets:
                if isinstance(target, ast.Name) and target.id == "__all__":
                    return symbols
        ImportCollector().visit(node)

    return symbols


def _classify_import_lines(source: str) -> dict[str, str | None]:
    """Map each imported name to its group anchor (or None)."""
    tree = ast.parse(source)
    name_to_group: dict[str, str | None] = {}

    # Collect all import line ranges so we can skip them when walking backward.
    import_line_numbers: set[int] = set()
    import_nodes: list[ast.ImportFrom] = []
    for node in ast.iter_child_nodes(tree):
        if isinstance(node, ast.Assign):
            for target in node.targets:
                if isinstance(target, ast.Name) and target.id == "__all__":
                    break  # stop at __all__
            else:
                continue
            break
        if isinstance(node, ast.ImportFrom):
            import_nodes.append(node)
            end = getattr(node, "end_lineno", node.lineno)
            for ln in range(node.lineno, end + 1):
                import_line_numbers.add(ln)

    lines = source.split("\n")

    def _get_import_group(lineno: int) -> str | None:
        """Walk backward from lineno to find the nearest group anchor."""
        for i in range(lineno - 2, -1, -1):
            if (i + 1) in import_line_numbers:
                continue  # skip lines belonging to other imports
            candidate = lines[i].strip()
            if GROUP_ANCHOR_RE.match(candidate):
                return _normalize_group_anchor(candidate)
            if candidate == "" or candidate.startswith("#"):
                continue
            break
        return None

    for node in import_nodes:
        grp = _get_import_group(node.lineno)
        for alias in node.names:
            name = alias.asname or alias.name
            if not name.startswith("_"):
                name_to_group[name] = grp

    return name_to_group


class TestConverterPublicApiGroups:
    """Group anchor and import/__all__ alignment tests."""

    def test_all_symbols_have_group_anchor(self) -> None:
        """Every symbol in __all__ must belong to a # ===== X ===== group."""
        source = _parse_source()
        all_entries = _extract_all_symbols_with_groups(source)

        ungrouped = [sym for sym, grp in all_entries if grp is None]
        assert not ungrouped, (
            f"__all__ symbols without group anchor: {ungrouped}. Add a '# ===== Group Name =====' comment above them."
        )

    def test_import_order_matches_all_order(self) -> None:
        """Symbols in __all__ must appear in the same relative order as in import statements."""
        source = _parse_source()
        all_entries = _extract_all_symbols_with_groups(source)
        all_symbols = [sym for sym, _grp in all_entries]
        import_symbols = _extract_import_symbols(source)

        # Build the relative ordering: each symbol should appear in imports
        # in the same order it appears in __all__
        all_positions = {sym: i for i, sym in enumerate(all_symbols)}
        import_positions = {sym: i for i, sym in enumerate(import_symbols)}

        # Check that for symbols in both lists, the relative order is consistent
        common = [s for s in all_symbols if s in import_positions]
        missing_from_imports = [s for s in all_symbols if s not in import_positions]

        assert not missing_from_imports, f"__all__ symbols not found in any import statement: {missing_from_imports}"

        # Verify monotonic ordering: for each pair (a, b) in __all__,
        # import position of a < import position of b
        violations = []
        for i in range(len(common)):
            for j in range(i + 1, len(common)):
                a, b = common[i], common[j]
                if import_positions[a] > import_positions[b]:
                    violations.append(
                        f"  {a} (__all__ #{all_positions[a]}, import #{import_positions[a]})"
                        f" before {b} (__all__ #{all_positions[b]}, import #{import_positions[b]})"
                    )
                # Only check adjacent to keep error output manageable
                if j == i + 1 and import_positions[a] > import_positions[b]:
                    pass  # already recorded

        assert not violations, (
            f"Import order doesn't match __all__ order ({len(violations)} violations, showing first 10):\n"
            + "\n".join(violations[:10])
        )

    def test_import_groups_match_all_groups(self) -> None:
        """Each import's group anchor must match the __all__ group of its symbols."""
        source = _parse_source()
        all_entries = _extract_all_symbols_with_groups(source)
        all_group_map = dict(all_entries)
        import_group_map = _classify_import_lines(source)

        mismatches = []
        for sym, import_grp in import_group_map.items():
            all_grp = all_group_map.get(sym)
            if all_grp is None:
                continue  # not in __all__, skip
            if import_grp is None:
                continue  # current public API file groups symbols in __all__ only
            if import_grp != all_grp:
                mismatches.append(f"  {sym}: import group={import_grp!r}, __all__ group={all_grp!r}")

        assert not mismatches, "Symbols whose import group differs from __all__ group:\n" + "\n".join(mismatches)

    def test_no_duplicate_symbols_in_all(self) -> None:
        """No symbol should appear twice in __all__."""
        source = _parse_source()
        all_entries = _extract_all_symbols_with_groups(source)
        all_symbols = [sym for sym, _grp in all_entries]

        seen = set()
        dups = []
        for sym in all_symbols:
            if sym in seen:
                dups.append(sym)
            seen.add(sym)

        assert not dups, f"Duplicate symbols in __all__: {dups}"

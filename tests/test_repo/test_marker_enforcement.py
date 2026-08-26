from __future__ import annotations

import ast
import tomllib
from pathlib import Path

import pytest
from tests.support.marker_taxonomy import PRIMARY_TEST_MARKERS
from tools.qa import PYTEST_PRIMARY_MARKER_DEBT_LIMIT, PYTEST_PRIMARY_MARKER_OVERLAP_LIMIT

pytestmark = pytest.mark.contract

_REPO_ROOT = Path(__file__).resolve().parents[2]


def _collect_marker_names(node: ast.AST) -> set[str]:
    marker_names: set[str] = set()

    def _walk(current: ast.AST) -> None:
        if (
            isinstance(current, ast.Attribute)
            and isinstance(current.value, ast.Attribute)
            and isinstance(current.value.value, ast.Name)
            and current.value.value.id == "pytest"
            and current.value.attr == "mark"
        ):
            marker_names.add(current.attr)
            return
        if isinstance(current, (ast.List, ast.Tuple, ast.Set)):
            for element in current.elts:
                _walk(element)
            return
        if isinstance(current, ast.Call):
            _walk(current.func)

    _walk(node)
    return marker_names


def _file_level_primary_markers(module: ast.Module) -> set[str]:
    markers: set[str] = set()
    for statement in module.body:
        if not isinstance(statement, ast.Assign):
            continue
        if not any(isinstance(target, ast.Name) and target.id == "pytestmark" for target in statement.targets):
            continue
        markers |= _collect_marker_names(statement.value)
    return markers & PRIMARY_TEST_MARKERS


def _decorator_primary_markers(node: ast.FunctionDef | ast.AsyncFunctionDef | ast.ClassDef) -> set[str]:
    markers: set[str] = set()
    for decorator in node.decorator_list:
        markers |= _collect_marker_names(decorator)
    return markers & PRIMARY_TEST_MARKERS


def _class_level_primary_markers(node: ast.ClassDef) -> set[str]:
    markers = _decorator_primary_markers(node)
    for statement in node.body:
        if not isinstance(statement, ast.Assign):
            continue
        if any(isinstance(target, ast.Name) and target.id == "pytestmark" for target in statement.targets):
            markers |= _collect_marker_names(statement.value) & PRIMARY_TEST_MARKERS
    return markers


def _test_node_primary_markers(
    statements: list[ast.stmt],
    inherited: set[str],
    *,
    owner: str = "",
) -> list[tuple[str, set[str]]]:
    nodes: list[tuple[str, set[str]]] = []
    for statement in statements:
        if isinstance(statement, ast.ClassDef):
            class_owner = f"{owner}::{statement.name}" if owner else statement.name
            nodes.extend(
                _test_node_primary_markers(
                    statement.body,
                    inherited | _class_level_primary_markers(statement),
                    owner=class_owner,
                )
            )
            continue
        if not isinstance(statement, (ast.FunctionDef, ast.AsyncFunctionDef)) or not statement.name.startswith("test_"):
            continue
        node_name = f"{owner}::{statement.name}" if owner else statement.name
        nodes.append((node_name, inherited | _decorator_primary_markers(statement)))
    return nodes


def _configured_test_roots() -> list[Path]:
    pyproject = tomllib.loads((_REPO_ROOT / "pyproject.toml").read_text(encoding="utf-8"))
    raw_roots = pyproject["tool"]["pytest"]["ini_options"]["testpaths"]
    return [_REPO_ROOT / str(root) for root in raw_roots]


def _test_files() -> list[Path]:
    paths: set[Path] = set()
    for root in _configured_test_roots():
        paths.update(root.rglob("test_*.py"))
    return sorted(paths)


def test_every_test_file_has_primary_marker_coverage() -> None:
    missing: list[str] = []
    overlapping: list[str] = []
    for path in _test_files():
        module = ast.parse(path.read_text(encoding="utf-8"))
        rel_path = path.relative_to(_REPO_ROOT).as_posix()
        for node_name, markers in _test_node_primary_markers(module.body, _file_level_primary_markers(module)):
            if not markers:
                missing.append(f"{rel_path}::{node_name}")
            elif len(markers) > 1:
                overlapping.append(f"{rel_path}::{node_name}: {', '.join(sorted(markers))}")

    assert len(overlapping) <= PYTEST_PRIMARY_MARKER_OVERLAP_LIMIT, (
        f"overlapping primary marker debt exceeded: actual={len(overlapping)}, "
        f"limit={PYTEST_PRIMARY_MARKER_OVERLAP_LIMIT}; first={overlapping[:20]}"
    )
    assert len(missing) <= PYTEST_PRIMARY_MARKER_DEBT_LIMIT, (
        f"primary marker debt exceeded: actual={len(missing)}, "
        f"limit={PYTEST_PRIMARY_MARKER_DEBT_LIMIT}; first={missing[:20]}"
    )

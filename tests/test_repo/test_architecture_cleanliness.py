"""Focused contracts for the architecture cleanliness scanner.

The scanner reports blocking violations for old-project-path references,
deprecated deep imports, unapproved runtime monkey-patches, compatibility
shims, and cross-package deep imports.

Rationale (per arch_audit findings):
    - Historical parity probes are explicitly excluded from Category 2;
      every live deprecated deep import is a blocking violation.
    - ``packages/core/src/docwen_core/ofd.py`` is the sole approved runtime
      patch module because its easyofd workarounds are explicit and tested.
"""

from __future__ import annotations

import ast
import sys
import tomllib
from pathlib import Path

import pytest

pytestmark = pytest.mark.unit


def test_audit_fix_import_scripts_do_not_return() -> None:
    """Historical one-shot import rewrite scripts should not be live tools."""
    project_root = Path(__file__).resolve().parents[2]
    audit_dir = project_root / "tools" / "audit"
    forbidden = sorted(path.name for path in audit_dir.glob("fix_*_imports.py"))

    assert forbidden == []


def test_plugin_manifests_do_not_duplicate_route_format_catalogs() -> None:
    """Built-in manifests use routes as the only supported-format catalog."""
    project_root = Path(__file__).resolve().parents[2]
    forbidden_keys = {
        "supported_actions",
        "supported_source_formats",
        "supported_target_formats",
    }
    violations: dict[str, list[str]] = {}

    for manifest_path in sorted((project_root / "packages" / "plugins").glob("**/manifest.py")):
        tree = ast.parse(manifest_path.read_text(encoding="utf-8"), filename=str(manifest_path))
        found = sorted(
            {
                node.value
                for node in ast.walk(tree)
                if isinstance(node, ast.Constant) and isinstance(node.value, str) and node.value in forbidden_keys
            }
        )
        if found:
            violations[str(manifest_path.relative_to(project_root))] = found

    assert violations == {}


def test_active_configuration_does_not_publish_reserved_ordered_list_setting() -> None:
    """Future numbering controls must return as complete features, not dead settings."""
    project_root = Path(__file__).resolve().parents[2]
    conversion = tomllib.loads((project_root / "configs" / "conversion.toml").read_text(encoding="utf-8"))

    assert "ordered_list" not in conversion["syntax"]
    for locale_path in sorted((project_root / "i18n" / "locales").glob("*.toml")):
        locale = tomllib.loads(locale_path.read_text(encoding="utf-8"))
        formatting = locale["settings"]["formatting"]
        assert "ordered_list_label" not in formatting, locale_path.name
        assert "ordered_list_tooltip" not in formatting, locale_path.name


def test_architecture_scanner_ignores_tmp_evidence_artifacts(tmp_path: Path) -> None:
    """Generated tmp evidence scripts must not pollute source architecture warnings."""
    project_root = Path(__file__).resolve().parents[2]
    sys.path.insert(0, str(project_root / "tools" / "validation"))
    try:
        import check_architecture_cleanliness as scanner
    finally:
        sys.path.remove(str(project_root / "tools" / "validation"))

    tmp_script = tmp_path / "tmp" / "old-system-probe.py"
    tmp_script.parent.mkdir(parents=True)
    legacy_import = "from docwen." + "converter import convert_docx_to_md\n"
    tmp_script.write_text(legacy_import, encoding="utf-8")
    tool_script = tmp_path / "tools" / "probe.py"
    tool_script.parent.mkdir(parents=True)
    tool_script.write_text(legacy_import, encoding="utf-8")

    violations = scanner.scan(tmp_path)
    normalized = [violation.replace("\\", "/") for violation in violations]

    assert not any("tmp/old-system-probe.py" in violation for violation in normalized)
    assert any("tools/probe.py" in violation for violation in normalized)


def test_architecture_scanner_covers_current_docs_and_content_named_shims(tmp_path: Path) -> None:
    project_root = Path(__file__).resolve().parents[2]
    sys.path.insert(0, str(project_root / "tools" / "validation"))
    try:
        import check_architecture_cleanliness as scanner
    finally:
        sys.path.remove(str(project_root / "tools" / "validation"))

    docs = tmp_path / "docs"
    docs.mkdir()
    retired_root = "docwen" + "旧"
    (docs / "current.md").write_text(f"Current path: {retired_root}\n", encoding="utf-8")
    tools = tmp_path / "tools"
    tools.mkdir()
    (tools / "entry.py").write_text('"""Compatibility wrapper for a retired entry point."""\n', encoding="utf-8")

    violations = scanner.scan(tmp_path)

    assert any("[1-old-project-path]" in violation and "docs/current.md" in violation for violation in violations)
    assert any("[4-shim-wrapper]" in violation and "tools/entry.py" in violation for violation in violations)


def test_architecture_scanner_rejects_unapproved_import_mutation(tmp_path: Path) -> None:
    project_root = Path(__file__).resolve().parents[2]
    sys.path.insert(0, str(project_root / "tools" / "validation"))
    try:
        import check_architecture_cleanliness as scanner
    finally:
        sys.path.remove(str(project_root / "tools" / "validation"))

    source = tmp_path / "packages" / "core" / "src" / "sample.py"
    source.parent.mkdir(parents=True)
    source.write_text("from vendor import Client\nClient.run = lambda self: None\n", encoding="utf-8")

    violations = scanner.scan(tmp_path)

    assert any("[3-monkey-patch]" in violation and "sample.py" in violation for violation in violations)

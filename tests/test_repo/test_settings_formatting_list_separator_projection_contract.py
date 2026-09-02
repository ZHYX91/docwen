"""Fail-closed evidence guards for VIS-2026-07-19-142 list separators."""

from __future__ import annotations

from pathlib import Path

import pytest
from tools.validation.source_family import read_source_text

pytestmark = pytest.mark.unit

PROJECT_ROOT = Path(__file__).resolve().parents[2]
REPORT_NAME = "formatting-yaml-list-separator-consumption-2026-07-19.md"


def _read(relative_path: str) -> str:
    return read_source_text(PROJECT_ROOT / relative_path)


def test_formatting_widget_apply_and_reset_own_the_exact_nested_value() -> None:
    model = _read("packages/apps/gui/src/docwen_gui/models/settings_config.py")
    widget = _read("packages/apps/gui/src/docwen_gui/widgets/settings/formatting_tab.py")
    vm = _read("packages/apps/gui/src/docwen_gui/view_models/settings_vm.py")
    registry = _read("packages/runtime/src/docwen_runtime/config/registry.py")
    shipped = _read("configs/conversion.toml")

    assert 'list_separator: str = "、"' in model
    assert "self._list_separator = QLineEdit(self)" in widget
    assert '"settings.formatting.list_separator_label"' in widget
    assert '"settings.formatting.list_separator_tooltip"' in widget
    assert '"list_separator", text' in widget
    assert "self._list_separator.setText(fmt.list_separator)" in widget
    assert ".strip()" not in "\n".join(line for line in widget.splitlines() if "list_separator" in line)

    assert 'm2d.get("list_separator", "、") is not None' in vm
    assert 'put("conversion.md_to_docx.list_separator", fmt.list_separator)' in vm
    assert '"conversion.md_to_docx.list_separator"' in registry
    assert 'list_separator = "、"' in shipped


def test_docx_xlsx_and_csv_consume_request_config_not_a_route_override() -> None:
    core = _read("packages/core/src/docwen_core/export_semantics/__init__.py")
    docx_converter = _read("packages/plugins/markdown/src/docwen_plugin_markdown/to_docx/converter.py")
    docx_filler = _read("packages/plugins/markdown/src/docwen_plugin_markdown/template_filler.py")
    sheet_converter = _read("packages/plugins/markdown/src/docwen_plugin_markdown/to_spreadsheet/converter.py")
    sheet_template = _read("packages/plugins/markdown/src/docwen_plugin_markdown/to_spreadsheet/template_xlsx.py")
    manifest = _read("packages/plugins/markdown/src/docwen_plugin_markdown/manifest.py")

    assert 'conv.get("md_to_docx", {})' in core
    assert 'yaml_list_separator = md_to_docx.get("list_separator", "、")' in core

    for source in (docx_converter, sheet_converter):
        assert '"conversion.md_to_docx.list_separator"' in source
        assert "get_yaml_list_separator" not in source
    assert "list_separator=_request_yaml_list_separator(context.config)" in docx_converter
    assert sheet_converter.count("list_separator=_request_yaml_list_separator(context)") == 2

    assert "yield from _iter_placeholder_list_items(item)" in docx_filler
    assert "return list_separator.join(items)" in docx_filler
    assert 'item in {"null", "None"}' in docx_filler
    assert "list_separator.join(" in sheet_template

    assert '"list_separator"' not in manifest


def test_direct_exact_value_and_artifact_regressions_remain_discoverable() -> None:
    widget_tests = _read("packages/apps/gui/tests/test_settings_formatting_tab.py")
    vm_tests = _read("packages/apps/gui/tests/test_settings_vm.py")
    reset_tests = _read("packages/apps/gui/tests/test_settings_vm_restore.py")
    core_tests = _read("packages/core/tests/test_export_semantics_runtime.py")
    runtime_tests = _read("packages/runtime/tests/test_runtime_config_wiring_*.py")
    docx_tests = _read("packages/plugins/markdown/tests/test_md_to_docx_*.py")
    sheet_tests = _read("packages/plugins/markdown/tests/test_md_to_spreadsheet_*.py")

    for source, test_name in (
        (widget_tests, "test_formatting_tab_preserves_yaml_list_separator_exactly"),
        (vm_tests, "test_apply_roundtrips_yaml_list_separator_exactly"),
        (reset_tests, "test_reset_formatting_group_preserves_export_and_resets_table_style"),
        (core_tests, "test_from_config_yaml_list_separator"),
        (runtime_tests, "test_snapshot_projection_consumes_nested_yaml_list_separator_override"),
        (docx_tests, "test_md_to_docx_yaml_list_separator_is_request_scoped_and_exact"),
        (sheet_tests, "test_template_name_joins_yaml_lists_with_request_config_separator"),
        (sheet_tests, "test_csv_template_chain_joins_yaml_lists_with_request_config_separator"),
    ):
        assert f"def {test_name}(" in source

    for exact_value in ('", "', '""'):
        assert exact_value in widget_tests
        assert exact_value in vm_tests
        assert exact_value in core_tests
        assert exact_value in runtime_tests
    assert "{{ignored}}" in docx_tests
    assert "['乙', '丙']" in sheet_tests

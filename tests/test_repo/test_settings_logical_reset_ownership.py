"""Guards for VIS-2026-07-18-134 Settings logical reset ownership."""

from __future__ import annotations

from pathlib import Path

import pytest

pytestmark = pytest.mark.unit

PROJECT_ROOT = Path(__file__).resolve().parents[2]
REPORT_NAME = "settings-logical-reset-ownership-2026-07-18.md"


def _read(relative_path: str) -> str:
    return (PROJECT_ROOT / relative_path).read_text(encoding="utf-8")


def _function_block(source: str, name: str, next_name: str) -> str:
    start = source.index(f"    def {name}(")
    end = source.index(f"    def {next_name}(", start)
    return source[start:end]


def test_logical_reset_plan_stays_runtime_owned_and_shared_by_gui_cli() -> None:
    from docwen_cli.commands.config import _TAB_GROUP_MAP
    from docwen_gui.widgets.settings.dialog import RESET_GROUPS
    from docwen_runtime.config.loader import RESET_EXCLUDED
    from docwen_runtime.config.registry import reset_plan_for_group, specs_for_group

    assert RESET_GROUPS == _TAB_GROUP_MAP
    for group in RESET_GROUPS.values():
        plan = reset_plan_for_group(group)
        assert plan.files or plan.dotted_keys, group

    export_plan = reset_plan_for_group("export")
    formatting_plan = reset_plan_for_group("formatting")
    document_plan = reset_plan_for_group("document")

    assert export_plan.files == ("export.toml",)
    assert set(export_plan.dotted_keys) == {
        "conversion.ocr_output.show_blockquote_title",
        "conversion.ocr_output.blockquote_title_override_by_locale",
        "conversion.export.base64_compress_enabled",
        "conversion.export.base64_compress_threshold_kb",
    }
    assert formatting_plan.files == ()
    assert len(formatting_plan.dotted_keys) == 26
    assert "document.style.table.md_to_docx.table_style_mode" in formatting_plan.dotted_keys
    assert "conversion.code_detection.code_font" not in formatting_plan.dotted_keys
    assert "conversion.export.base64_compress_enabled" not in formatting_plan.dotted_keys
    assert document_plan.files == ()
    assert len(document_plan.dotted_keys) == 11
    assert "document.to_md_image_extraction_mode" not in document_plan.dotted_keys
    assert "document.to_md_ocr_placement_mode" not in document_plan.dotted_keys

    precise_counts = {
        "general": 4,
        "text": 6,
        "proofread": 6,
        "formatting": 26,
        "document": 11,
        "spreadsheet": 7,
        "layout": 6,
        "link": 8,
        "other": 2,
        "output": 6,
        "logging": 10,
    }
    for group, count in precise_counts.items():
        plan = reset_plan_for_group(group)
        assert plan.files == (), group
        assert len(plan.dotted_keys) == count, group

    assert reset_plan_for_group("image").files == ("image.toml",)
    assert "link.path_resolution.search_dirs" not in reset_plan_for_group("link").dotted_keys
    assert "output.manifest.save_to_output" not in reset_plan_for_group("output").dotted_keys
    assert "logger.format" not in reset_plan_for_group("logging").dotted_keys
    assert {spec.rel_path for spec in specs_for_group("proofread")} == {
        "proofread/engine.toml",
        "proofread/skip.toml",
        "proofread/pairs.toml",
        "proofread/symbol_map.toml",
        "proofread/typos.toml",
        "proofread/sensitive_words.toml",
    }
    assert {
        "proofread/symbol_map.toml",
        "proofread/typos.toml",
        "proofread/sensitive_words.toml",
    } == RESET_EXCLUDED

    loader = _read("packages/runtime/src/docwen_runtime/config/loader.py")
    port = _read("packages/application/src/docwen_application/ports/runtime.py")
    adapter = _read("packages/bundle/src/docwen_bundle/config_port.py")
    vm = _read("packages/apps/gui/src/docwen_gui/view_models/settings_vm.py")
    cli = _read("packages/apps/cli/src/docwen_cli/commands/config.py")
    cli_main = _read("packages/apps/cli/src/docwen_cli/main.py")

    reset_group_block = _function_block(loader, "reset_group", "reset_all")
    reset_all_block = _function_block(loader, "reset_all", "set_value")
    vm_block = _function_block(vm, "reset_group", "reset_all")

    assert "plan = reset_plan_for_group(group)" in reset_group_block
    assert "rel_path not in RESET_EXCLUDED" in reset_group_block
    assert "dotted_by_file = self._plan_reset_values(plan.dotted_keys)" in reset_group_block
    assert "self._reset_grouped_values_on_disk(dotted_by_file)" in reset_group_block
    assert "self._run_user_file_transaction(" in reset_group_block
    assert 'operation=f"reset_group:{group}"' in reset_group_block
    assert "self.reset_file(" not in reset_group_block
    assert "self.reset_values(" not in reset_group_block

    assert "spec.rel_path not in RESET_EXCLUDED" in reset_all_block
    assert "self._run_user_file_transaction(" in reset_all_block
    assert 'operation="reset_all"' in reset_all_block
    assert "self.reset_file(" not in reset_all_block

    assert "def reset_group(self, group: str) -> bool:" in port
    assert "Reset a group all-or-nothing when handled-failure compensation succeeds." in port
    assert "return self._loader.reset_group(group)" in adapter
    assert 'getattr(controller, "config_port", None)' in vm_block
    assert "reset_group(group)" in vm_block
    assert "get_config_loader" not in vm_block
    assert "specs_for_group" not in vm_block
    assert 'getattr(controller, "config_port", None)' in cli
    assert "config_port.reset_group(group)" in cli
    assert "config_port.reset_all()" in cli
    assert "loader.reset_group(group)" not in cli
    assert "get_config_loader" not in cli
    assert "get_config_loader" not in cli_main
    assert "config_port_factory" in cli_main
    assert "specs_for_group" not in cli

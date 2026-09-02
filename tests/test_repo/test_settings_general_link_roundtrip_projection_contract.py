"""Guards for VIS-2026-07-18-135 General/Link config round-trip evidence."""

from __future__ import annotations

from pathlib import Path

import pytest
from tools.validation.source_family import read_source_text

pytestmark = pytest.mark.unit

PROJECT_ROOT = Path(__file__).resolve().parents[2]
REPORT_NAME = "settings-general-link-config-roundtrip-2026-07-18.md"


def _read(relative_path: str) -> str:
    return read_source_text(PROJECT_ROOT / relative_path)


def _function_block(source: str, name: str, next_name: str) -> str:
    start = source.index(f"    def {name}(")
    end = source.index(f"    def {next_name}(", start)
    return source[start:end]


def test_general_and_link_config_roundtrip_contract_stays_wired_with_current_runtime_boundary() -> None:
    general = _read("packages/apps/gui/src/docwen_gui/widgets/settings/general_tab.py")
    vm = _read("packages/apps/gui/src/docwen_gui/view_models/settings_vm.py")
    regression = _read("packages/apps/gui/tests/test_settings_general_link_roundtrip.py")
    main_window = _read("packages/apps/gui/src/docwen_gui/main_window.py")
    window_policy = _read("packages/apps/gui/src/docwen_gui/window_behavior.py")
    settings_dialog = _read("packages/apps/gui/src/docwen_gui/widgets/settings/dialog.py")
    runtime_regression = _read("packages/apps/gui/tests/test_main_window_window_behavior_*.py")
    production_sources = "\n".join(
        path.read_text(encoding="utf-8")
        for path in (PROJECT_ROOT / "packages").rglob("*.py")
        if "/src/" in path.relative_to(PROJECT_ROOT).as_posix()
    )

    for signal, handler in (
        ("remember.toggled", "_on_remember_state_toggled"),
        ("auto_center.toggled", "_on_auto_center_toggled"),
        ("expand.toggled", "_on_expand_side_panels_toggled"),
    ):
        assert f"{signal}.connect(self.{handler})" in general
        assert f"def {handler}(" in general
    assert '"Expand side panels with window"' in general
    assert '"Expand side panels by default"' not in general

    persist = _function_block(vm, "_persist_to_controller_config", "_collect_conversion_defaults")
    link_keys = {
        "link.format.image_link_style",
        "link.format.md_file_link_style",
        "link.non_embed_links.wiki_mode",
        "link.non_embed_links.markdown_mode",
        "link.embed_links.wiki_image_mode",
        "link.embed_links.markdown_image_mode",
        "link.embed_links.md_file_mode",
        "link.embedding.max_depth",
    }
    for key in link_keys:
        assert f'put("{key}"' in persist

    for token in (
        "fresh_view_model",
        "auto_link_bare_url",
        "path_resolution.search_dirs",
        "error_handling.file_not_found",
        "get_change_summary",
        "QSignalSpy",
    ):
        assert token in regression

    # VIS-137 closed the General runtime boundary.  Link processing is
    # request-scoped: no process-wide fallback may become a second fact source.
    for key in ("remember_gui_state", "auto_center", "expand_side_panels"):
        assert f'"gui.window.{key}"' in window_policy
        assert key in main_window
        assert key in runtime_regression
    assert "WindowBehaviorPolicy" in window_policy
    assert "cfg_port.set_many(" in main_window
    assert "settings_source_changed" in settings_dialog
    assert "settings_source_changed.connect(self._apply_runtime_window_settings)" in main_window
    assert "_get_link_cfg" not in production_sources
    assert "configure_link_runtime_config" not in production_sources
    assert production_sources.count("process_markdown_links(") >= 3
    assert production_sources.count("link_config=") >= 3

"""Saved proofreading defaults, temporary GUI choices and request intent."""

from __future__ import annotations

from types import SimpleNamespace
from typing import TYPE_CHECKING, cast

import pytest

from docwen_application.controller import ApplicationController
from docwen_bundle.config_port import ConfigPortAdapter
from docwen_core.detection import inspect_file
from docwen_core.models import FILE_INSPECTION_METADATA_KEY
from docwen_core.models.file_ref import FileRef
from docwen_core.models.request import POSTPROCESS_PROOFREAD_OPTION
from docwen_core.options.proofread import resolve_proofread_switches
from docwen_gui.execution_requests import ExecutionRequestBuilder
from docwen_gui.path_identity import normalize_path
from docwen_gui.view_models.action_area_vm import PROOFREAD_OPTION_KEYS, ActionAreaViewModel
from docwen_gui.view_models.settings_vm import SettingsViewModel

from ._action_area_vm_support import FakeMainWindowViewModel_for_port, _write_minimal_base_config_tree

pytestmark = pytest.mark.integration

if TYPE_CHECKING:
    from docwen_gui.view_models.batch_list_vm import BatchListViewModel
    from docwen_gui.view_models.main_window_vm import MainWindowViewModel


def test_saved_and_canceled_defaults_refresh_without_overwriting_current_choices(tmp_path) -> None:
    config_dir = tmp_path / "configs"
    _write_minimal_base_config_tree(config_dir)
    port = ConfigPortAdapter(base_dir=config_dir, user_dir=config_dir)
    settings = SettingsViewModel(controller=ApplicationController(config_port=port))
    action = ActionAreaViewModel(cast("MainWindowViewModel", FakeMainWindowViewModel_for_port(port)))
    action.setup_for_md_to_document("/first.md")
    assert all(action.proofread_options.values())
    for key in PROOFREAD_OPTION_KEYS:
        settings.set_field("proofread", key, False)
    action.refresh_proofread_defaults()
    assert all(action.proofread_options.values())
    settings.cancel_changes()
    action.refresh_proofread_defaults()
    assert all(action.proofread_options.values())

    for key in PROOFREAD_OPTION_KEYS:
        settings.set_field("proofread", key, False)
    assert settings.apply_settings()
    action.refresh_proofread_defaults()
    assert not any(action.proofread_options.values())
    action.set_proofread_option("sensitive_word", True)
    action.target_format = "pdf"
    settings.set_field("proofread", "symbol_pairing", True)
    assert settings.apply_settings()
    action.refresh_proofread_defaults()
    action.target_format = "docx"
    assert action.proofread_options == {
        "symbol_pairing": True,
        "symbol_correction": False,
        "typos_rule": False,
        "sensitive_word": True,
    }
    # Reopening a saved profile and selecting a new input start with its defaults.
    fresh_port = ConfigPortAdapter(base_dir=config_dir, user_dir=config_dir)
    fresh = ActionAreaViewModel(cast("MainWindowViewModel", FakeMainWindowViewModel_for_port(fresh_port)))
    fresh.setup_for_md_to_document("/second.md")
    assert fresh.proofread_options == {
        "symbol_pairing": True,
        "symbol_correction": False,
        "typos_rule": False,
        "sensitive_word": False,
    }


@pytest.mark.parametrize("enabled", [False, True])
def test_saved_defaults_reach_request_without_changing_core_option_priority(tmp_path, enabled) -> None:
    config_dir = tmp_path / "configs"
    _write_minimal_base_config_tree(config_dir)
    port = ConfigPortAdapter(base_dir=config_dir, user_dir=config_dir)
    settings = SettingsViewModel(controller=ApplicationController(config_port=port))
    for key in PROOFREAD_OPTION_KEYS:
        settings.set_field("proofread", key, enabled)
    assert settings.apply_settings()
    action = ActionAreaViewModel(cast("MainWindowViewModel", FakeMainWindowViewModel_for_port(port)))
    source = tmp_path / "input.md"
    source.write_text("Visible input.", encoding="utf-8")
    action.setup_for_md_to_document(str(source))
    ref = FileRef(
        path=str(source),
        format="markdown",
        category="markdown",
        metadata={FILE_INSPECTION_METADATA_KEY: inspect_file(source).to_dict()},
    )
    builder = ExecutionRequestBuilder(
        cast("MainWindowViewModel", SimpleNamespace(files=[ref], controller=None)),
        cast("BatchListViewModel", SimpleNamespace(get_file_entry=lambda _path: None)),
        file_contexts=lambda: {normalize_path(str(source)): ("markdown", "markdown")},
        selected_template=lambda: ("docx", "standard"),
    )
    request, _context = builder.single(
        file_path=str(source),
        target_format="docx",
        action_name="",
        options=action.collect_options(),
        route_options=("remove_numbering", "template_name"),
    )
    expected = {f"enable_{key}": enabled for key in PROOFREAD_OPTION_KEYS}
    assert (
        resolve_proofread_switches({"proofread": {"engine": dict.fromkeys(expected, not enabled)}}, expected)
        == expected
    )
    if enabled:
        assert request.options[POSTPROCESS_PROOFREAD_OPTION] == expected
    else:
        assert POSTPROCESS_PROOFREAD_OPTION not in request.options

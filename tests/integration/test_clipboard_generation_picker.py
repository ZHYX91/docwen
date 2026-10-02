"""Generation pickers must use the admitted source's actual Runtime routes."""

import pytest

from docwen_application.controller import ApplicationController
from docwen_bundle.runtime_factory import create_runtime_port
from docwen_gui.view_models._runtime_route_filter import RuntimeRouteSource
from docwen_gui.view_models.action_area_vm import ActionAreaViewModel
from docwen_gui.view_models.main_window_vm import MainWindowViewModel

pytestmark = [pytest.mark.integration, pytest.mark.pr_gate, pytest.mark.release_gate]


@pytest.mark.parametrize("mixed", [False, True])
def test_clipboard_generation_picker_exposes_only_common_routes(tmp_path, monkeypatch, mixed):
    monkeypatch.setattr("docwen_bundle.runtime_factory._runtime_workspace_root", lambda: tmp_path / "runtime")
    controller = ApplicationController(runtime_port=create_runtime_port())
    controller.start()
    try:
        action = ActionAreaViewModel(MainWindowViewModel(controller=controller))
        sources = (RuntimeRouteSource("clipboard_document", "markdown"),)
        if mixed:
            sources += (RuntimeRouteSource("md", "markdown"),)
        action.setup_for_md_to_document(str(tmp_path / "document.dwclip"), source_inputs=sources)
        if controller.describe_gui_runtime_capabilities()["runtime"]["platform"] not in {"windows", "linux"}:
            assert action.available_target_formats == []
            assert not action.generation_ready
            return
        assert ("MD" in action.available_target_formats) is not mixed
        assert "DOCX" in action.available_target_formats
        if not mixed:
            action.target_format = "md"
            action.set_template_ready(False)
            assert action.generation_ready
            assert not action.show_numbering
            assert not action.show_proofread
            assert action.collect_options() == {}
            assert action.target_route_choices_result.get("md").routes[0].source == "clipboard_document"
        action.setup_for_md_to_spreadsheet(str(tmp_path / "document.dwclip"), source_inputs=sources)
        assert {"XLSX", "CSV"} <= set(action.available_target_formats)
        assert "MD" not in action.available_target_formats
        action.setup_for_md_to_document(str(tmp_path / "ordinary.md"))
        assert "MD" not in action.available_target_formats
    finally:
        controller.stop()

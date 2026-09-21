from __future__ import annotations

from io import BytesIO

import pytest
from PIL import Image

from docwen_core.mermaid_runtime import MermaidRuntime
from docwen_gui.models.settings_config import SettingsConfig
from docwen_gui.view_models.settings_vm import SettingsViewModel
from docwen_gui.widgets.settings.formatting_tab import FormattingTab

pytestmark = pytest.mark.gui


def test_mermaid_path_is_draft_state_and_detection_does_not_change_mode(qtbot, monkeypatch) -> None:
    from docwen_gui.widgets.settings import mermaid_controls

    monkeypatch.setattr(
        mermaid_controls,
        "inspect_mermaid_runtime",
        lambda _: MermaidRuntime(False, "unsupported_version", cli_version="10", mermaid_version="10"),
    )
    config = SettingsConfig()
    config.formatting.mermaid_mode = "image"
    vm = SettingsViewModel(config=config)
    tab = FormattingTab(vm)
    qtbot.addWidget(tab)
    controls = tab._mermaid_controls
    controls.path.setText("custom/mmdc")
    assert vm.config.formatting.mermaid_cli_path == "custom/mmdc"
    assert vm.config.formatting.mermaid_mode == "image"
    assert "10" in controls.status.text()
    vm.cancel_changes()
    tab.reload_from_config()
    assert controls.path.text() == ""


def test_test_render_reports_actual_completion_and_preserves_mode(qtbot, monkeypatch) -> None:
    from docwen_gui.widgets.settings import mermaid_controls

    output = BytesIO()
    Image.new("RGB", (200, 60), "white").save(output, "PNG")
    seen = []

    def render(source, **kwargs):
        seen.append((source, kwargs["cli_path"]))
        kwargs["cancellation"].check()
        return output.getvalue()

    monkeypatch.setattr(mermaid_controls, "render_mermaid_png", render)
    vm = SettingsViewModel(config=SettingsConfig())
    tab = FormattingTab(vm)
    qtbot.addWidget(tab)
    controls = tab._mermaid_controls
    controls.path.setText("custom/mmdc")
    controls.test_render()
    qtbot.waitUntil(lambda: not controls.operation.busy)
    assert not controls.preview.pixmap().isNull()
    assert seen == [("flowchart LR\nA[Mermaid] --> B[OK]", "custom/mmdc")]
    assert vm.config.formatting.mermaid_mode == "code"

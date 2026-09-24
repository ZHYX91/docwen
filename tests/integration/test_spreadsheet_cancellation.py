"""Real spreadsheet plugin through runtime cancellation and publication cleanup."""

from pathlib import Path

import openpyxl
import pytest

from docwen_core.models import ConversionRequest, FileRef, OutputPolicy
from docwen_plugin_spreadsheet.plugin import SpreadsheetPlugin
from docwen_runtime.engine.route_resolver import RouteResolver
from docwen_runtime.engine.task_manager import TaskManager
from docwen_runtime.output.finalizer import OutputFinalizer
from docwen_runtime.plugin_registry.registry import PluginRegistry
from docwen_runtime.workspace.manager import WorkspaceManager

pytestmark = [pytest.mark.integration, pytest.mark.pr_gate, pytest.mark.release_gate]


@pytest.mark.parametrize("boundary", ["after-save", "precommit"])
@pytest.mark.parametrize("source_format", ["csv", "tsv"])
def test_spreadsheet_cancel_does_not_publish_or_leave_workspace(
    tmp_path: Path, monkeypatch, boundary, source_format
) -> None:
    source = tmp_path / f"source.{source_format}"
    source.write_text("00123\n=1+1\n", encoding="utf-8")
    original = source.read_bytes()
    staging = tmp_path / "workspaces"
    output = tmp_path / "published"
    registry = PluginRegistry()
    registry.register(SpreadsheetPlugin())
    workspaces = WorkspaceManager(root_dir=str(staging))
    manager = TaskManager(registry, RouteResolver(registry), workspaces, OutputFinalizer())
    request = ConversionRequest(
        request_id="cancel-sheet",
        input_refs=[FileRef(path=str(source), format=source_format, category="spreadsheet")],
        target_format="xlsx",
        output_policy=OutputPolicy(output_dir=str(output)),
    )
    if boundary == "after-save":
        real_save = openpyxl.Workbook.save

        def save_and_cancel(workbook, filename):
            real_save(workbook, filename)
            manager.cancel(request.request_id)

        monkeypatch.setattr(openpyxl.Workbook, "save", save_and_cancel)
    else:
        real_prepare = OutputFinalizer._copy_to_temp

        def prepare_and_cancel(*args, **kwargs):
            result = real_prepare(*args, **kwargs)
            manager.cancel(request.request_id)
            return result

        monkeypatch.setattr(OutputFinalizer, "_copy_to_temp", staticmethod(prepare_and_cancel))
    events = []
    result = manager.execute_single(request, on_event=events.append)
    assert not result.success and result.error is not None
    assert result.error.error_type == "cancelled"
    assert events[-1].event_type == "task_cancelled"
    assert not output.exists() or not list(output.iterdir())
    assert workspaces.get(request.request_id) is None
    assert not list(staging.rglob("*.xlsx"))
    assert source.read_bytes() == original

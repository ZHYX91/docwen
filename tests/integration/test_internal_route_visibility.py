"""Public discovery excludes internal clipboard formats while GUI keeps them."""

from __future__ import annotations

from pathlib import Path

import pytest

from docwen_application.controller import ApplicationController
from docwen_bundle.runtime_factory import create_runtime_port
from docwen_cli.machine.query_service import MachineQueryService
from docwen_gui.view_models._runtime_route_filter import (
    RuntimeRouteSource,
    discover_runtime_route_choices,
)

pytestmark = [pytest.mark.integration, pytest.mark.pr_gate, pytest.mark.release_gate]


def test_runtime_factory_separates_public_machine_formats_from_gui_internal_routes(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(
        "docwen_bundle.runtime_factory._runtime_workspace_root",
        lambda: tmp_path / "runtime-workspaces",
    )
    runtime = create_runtime_port()
    controller = ApplicationController(runtime_port=runtime)
    controller.start()
    try:
        machine = MachineQueryService(controller)
        formats = machine.list_resources("formats")
        public_ids = {item["id"] for item in formats["resources"]}
        assert "clipboard_document" not in public_ids
        assert "markdown" in public_ids

        public_projection = controller.describe_runtime_capabilities()
        gui_projection = controller.describe_gui_runtime_capabilities()
        assert all(source["id"] != "clipboard_document" for source in public_projection["sources"])
        assert any(source["id"] == "clipboard_document" for source in gui_projection["sources"])

        choices = discover_runtime_route_choices(
            controller,
            sources=(RuntimeRouteSource("clipboard_document", "markdown"),),
            operation="conversion",
        )
        assert choices.status == "ready"
        assert {"md", "docx", "xlsx", "csv"} <= set(choices.targets)

        public_optimizations = machine.list_resources("optimizations")
        assert public_optimizations["kind"] == "optimizations"
        assert all(item["id"] != "clipboard_document" for item in public_optimizations["resources"])
    finally:
        controller.stop()

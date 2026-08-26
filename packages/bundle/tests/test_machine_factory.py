"""Machine stdio composition-root event wiring contracts."""

from __future__ import annotations

import io
from typing import Any, cast

import pytest

from docwen_core.models import TaskEvent

pytestmark = pytest.mark.unit


def test_machine_factory_projects_runtime_events_only_after_server_exists(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import docwen_application
    import docwen_application.conversion_service as conversion_service_module
    import docwen_bundle.config_port as config_port_module
    import docwen_bundle.machine_factory as machine_factory
    import docwen_bundle.runtime_factory as runtime_factory_module
    import docwen_cli.machine as machine_module
    import docwen_cli.machine.query_service as query_service_module
    import docwen_runtime.output.artifact_bundle as artifact_bundle_module

    captured: dict[str, Any] = {}
    runtime_port = object()

    def create_runtime_port(*, config_loader: object, event_callback: Any) -> object:
        captured["config_loader"] = config_loader
        captured["event_callback"] = event_callback
        event_callback(TaskEvent("task.before-server", "task_progress", 1, payload={"percent": 5}))
        return runtime_port

    class Controller:
        def __init__(self, *, runtime_port: object, config_port: object) -> None:
            self.runtime_port = runtime_port
            self.config_port = config_port
            self.started = False

        def start(self) -> None:
            self.started = True

        def stop(self) -> None:
            self.started = False

    class Server:
        def __init__(
            self,
            service: object,
            reader: object,
            writer: object,
            *,
            query_service: object,
            close_callback: Any,
        ) -> None:
            self.service = service
            self.reader = reader
            self.writer = writer
            self.query_service = query_service
            self.close_callback = close_callback
            self.events: list[TaskEvent] = []

        def report_runtime_event(self, event: TaskEvent) -> None:
            self.events.append(event)

    monkeypatch.setattr(runtime_factory_module, "create_runtime_port", create_runtime_port)
    monkeypatch.setattr(docwen_application, "ApplicationController", Controller)
    monkeypatch.setattr(config_port_module, "ConfigPortAdapter", lambda loader: ("config", loader))
    monkeypatch.setattr(
        conversion_service_module,
        "ConversionService",
        lambda controller, committer: (controller, committer),
    )
    monkeypatch.setattr(query_service_module, "MachineQueryService", lambda controller, gui: (controller, gui))
    monkeypatch.setattr(artifact_bundle_module, "ArtifactBundleCommitter", lambda: "committer")
    monkeypatch.setattr(machine_module, "MachineProtocolServer", Server)

    config_loader = object()
    server = cast(
        Server,
        machine_factory.create_machine_server(
            config_loader=config_loader,  # type: ignore[arg-type]
            reader=io.BytesIO(),
            writer=io.BytesIO(),
        ),
    )
    runtime_event = TaskEvent(
        "task.accepted",
        "task_progress",
        2,
        payload={"percent": 50, "message": r"C:\private\source.md"},
    )
    captured["event_callback"](runtime_event)

    assert captured["config_loader"] is config_loader
    assert server.events == [runtime_event]

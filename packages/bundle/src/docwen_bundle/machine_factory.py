"""Production composition root for ``docwen serve --stdio``."""

from __future__ import annotations

import sys
from typing import TYPE_CHECKING, Any, BinaryIO

if TYPE_CHECKING:
    from docwen_cli.machine import MachineProtocolServer
    from docwen_core.models import TaskEvent
    from docwen_runtime.config import ConfigLoader


def create_machine_server(
    *,
    config_loader: ConfigLoader | None = None,
    reader: BinaryIO | None = None,
    writer: BinaryIO | None = None,
    gui_control: Any | None = None,
) -> MachineProtocolServer:
    """Create the fully composed Machine Protocol server without opening a network port."""

    from docwen_application import ApplicationController
    from docwen_application.conversion_service import ConversionService
    from docwen_bundle.config_port import ConfigPortAdapter
    from docwen_bundle.runtime_factory import create_runtime_port
    from docwen_cli.machine import MachineProtocolServer
    from docwen_cli.machine.query_service import MachineQueryService
    from docwen_runtime.config import ConfigLoader
    from docwen_runtime.output.artifact_bundle import ArtifactBundleCommitter

    if config_loader is None:
        config_loader = ConfigLoader(runtime_overrides={"logger": {"console_enable": False}})
    server: MachineProtocolServer | None = None

    def project_runtime_event(event: TaskEvent) -> None:
        if server is not None:
            server.report_runtime_event(event)

    runtime = create_runtime_port(
        config_loader=config_loader,
        event_callback=project_runtime_event,
    )
    controller = ApplicationController(
        runtime_port=runtime,
        config_port=ConfigPortAdapter(config_loader),
    )
    controller.start()
    service = ConversionService(controller, ArtifactBundleCommitter())
    query_service = MachineQueryService(controller, gui_control)
    binary_reader: Any = reader if reader is not None else sys.stdin.buffer
    binary_writer: Any = writer if writer is not None else sys.stdout.buffer
    server = MachineProtocolServer(
        service,
        binary_reader,
        binary_writer,
        query_service=query_service,
        close_callback=controller.stop,
    )
    return server


__all__ = ["create_machine_server"]

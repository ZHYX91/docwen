from __future__ import annotations

import sys
from typing import Any

from docwen_runtime.security import NetworkGuardInstallationError, dependency_egress_guard


def _main_with_guard_active(argv: list[str] | None = None) -> int:
    from docwen_bundle.config_port import ConfigPortAdapter
    from docwen_bundle.gui_control_adapter import create_gui_control_adapter
    from docwen_bundle.machine_factory import create_machine_server
    from docwen_bundle.runtime_factory import create_runtime_port
    from docwen_cli.main import main as cli_main
    from docwen_runtime.config import ConfigLoader

    config_loader: Any | None = None

    def shared_loader() -> Any:
        nonlocal config_loader
        if config_loader is None:
            # CLI owns stdout/stderr presentation. Runtime diagnostics remain
            # in the file log and never contaminate machine JSON output.
            config_loader = ConfigLoader(runtime_overrides={"logger": {"console_enable": False}})
        return config_loader

    def runtime_port_factory() -> Any:
        return create_runtime_port(config_loader=shared_loader())

    def config_port_factory() -> Any:
        return ConfigPortAdapter(shared_loader())

    return cli_main(
        argv,
        runtime_port_factory=runtime_port_factory,
        config_port_factory=config_port_factory,
        gui_control_port_factory=create_gui_control_adapter,
        machine_server_factory=lambda: create_machine_server(
            config_loader=shared_loader(),
            gui_control=create_gui_control_adapter(),
        ),
    )


def main(argv: list[str] | None = None) -> int:
    """Run the composed CLI with dependency egress protection enforced."""

    from docwen_runtime.logging import pre_init_logging

    # Keep successful machine-readable CLI runs silent on stderr while still
    # surfacing and buffering warnings raised before ConfigLoader is available.
    pre_init_logging("WARNING")

    try:
        with dependency_egress_guard():
            return _main_with_guard_active(argv)
    except NetworkGuardInstallationError:
        print("错误: 安全检查失败", file=sys.stderr)
        return NetworkGuardInstallationError.exit_code


if __name__ == "__main__":
    raise SystemExit(main())

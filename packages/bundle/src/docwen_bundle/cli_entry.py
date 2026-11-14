from __future__ import annotations

import sys
from typing import Any

from docwen_runtime.security import NetworkGuardInstallationError, dependency_egress_guard


def _main_with_guard_active(argv: list[str] | None = None) -> int:
    from docwen_cli.main import main as cli_main

    config_loader: Any | None = None

    def shared_loader() -> Any:
        nonlocal config_loader
        if config_loader is None:
            from docwen_runtime.config import ConfigLoader

            # CLI owns stdout/stderr presentation. Runtime diagnostics remain
            # in the file log and never contaminate machine JSON output.
            config_loader = ConfigLoader(runtime_overrides={"logger": {"console_enable": False}})
        return config_loader

    def runtime_port_factory() -> Any:
        from docwen_bundle.runtime_factory import create_runtime_port

        return create_runtime_port(config_loader=shared_loader())

    def config_port_factory() -> Any:
        from docwen_bundle.config_port import ConfigPortAdapter

        return ConfigPortAdapter(shared_loader())

    def gui_control_port_factory() -> Any:
        from docwen_bundle.gui_control_adapter import create_gui_control_adapter

        return create_gui_control_adapter()

    def machine_server_factory() -> Any:
        from docwen_bundle.machine_factory import create_machine_server

        return create_machine_server(config_loader=shared_loader(), gui_control=gui_control_port_factory())

    return cli_main(
        argv,
        runtime_port_factory=runtime_port_factory,
        config_port_factory=config_port_factory,
        gui_control_port_factory=gui_control_port_factory,
        machine_server_factory=machine_server_factory,
    )


def main(argv: list[str] | None = None) -> int:
    """Run the composed CLI with dependency egress protection enforced."""

    arguments = list(sys.argv[1:] if argv is None else argv)
    if arguments[:1] == ["serve"] and "--managed" in arguments:
        import argparse

        from docwen_bundle.windows_machine_owner import run_managed_stdio

        parser = argparse.ArgumentParser(prog="DocWenCLI serve", allow_abbrev=False)
        parser.add_argument("--stdio", action="store_true", required=True)
        parser.add_argument("--managed", action="store_true", required=True)
        parser.add_argument("--parent-pid", type=int, required=True)
        parsed = parser.parse_args(arguments[1:])
        if sys.platform != "win32" or not 0 < parsed.parent_pid < 2**32:
            parser.error("managed stdio requires Windows and a live parent process ID")
        return run_managed_stdio(parsed.parent_pid)

    from docwen_runtime.logging import pre_init_logging
    from docwen_runtime.profile_paths import ProfileSelectionError, bind_process_profile

    # Keep successful machine-readable CLI runs silent on stderr while still
    # surfacing and buffering warnings raised before ConfigLoader is available.
    pre_init_logging("WARNING")

    try:
        with bind_process_profile(), dependency_egress_guard():
            return _main_with_guard_active(argv)
    except ProfileSelectionError as exc:
        print(str(exc), file=sys.stderr)
        return 4
    except NetworkGuardInstallationError:
        print("错误: 安全检查失败", file=sys.stderr)
        return NetworkGuardInstallationError.exit_code


if __name__ == "__main__":
    raise SystemExit(main())

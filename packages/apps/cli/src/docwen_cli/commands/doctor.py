"""Doctor command backed by the canonical runtime capability projection.

The command performs local health probes exactly once and reports the runtime
composition without rebuilding a second source/gate model inside the CLI.
Machine output intentionally contains stable identifiers and reason codes, not
host filesystem paths or raw exception strings.
"""

from __future__ import annotations

import argparse
import contextlib
import os
import tempfile
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any

from docwen_cli.capabilities import runtime_capability_projection
from docwen_cli.exit_codes import ExitCode
from docwen_cli.i18n import cli_t
from docwen_cli.parser import get_common_parser


@dataclass(frozen=True, slots=True)
class DoctorCheck:
    """One path-free diagnostic result."""

    id: str
    kind: str
    label: str
    status: str
    reason: str | None = None

    @property
    def ok(self) -> bool:
        return self.status == "ok"

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


def _check_temp_directory() -> DoctorCheck:
    """Verify the process can create and remove a file in its temp directory."""

    probe_path: Path | None = None
    try:
        temp_dir = Path(tempfile.gettempdir())
        if not temp_dir.is_dir():
            return DoctorCheck(
                id="path.temp_directory",
                kind="path",
                label="Temporary directory",
                status="fail",
                reason="not_directory",
            )
        descriptor, raw_path = tempfile.mkstemp(prefix=".docwen-doctor-", dir=temp_dir)
        os.close(descriptor)
        probe_path = Path(raw_path)
        probe_path.write_bytes(b"ok")
        probe_path.unlink()
    except OSError:
        if probe_path is not None:
            with contextlib.suppress(OSError):
                probe_path.unlink(missing_ok=True)
        return DoctorCheck(
            id="path.temp_directory",
            kind="path",
            label="Temporary directory",
            status="fail",
            reason="not_writable",
        )
    return DoctorCheck(
        id="path.temp_directory",
        kind="path",
        label="Temporary directory",
        status="ok",
    )


def _check_config(controller: Any) -> DoctorCheck:
    """Exercise the injected config port without exposing configuration values."""

    config_port = getattr(controller, "config_port", None)
    if config_port is None:
        return DoctorCheck(
            id="config.load",
            kind="config",
            label="Configuration",
            status="fail",
            reason="config_port_unavailable",
        )
    snapshot = getattr(config_port, "snapshot", None)
    if not callable(snapshot):
        return DoctorCheck(
            id="config.load",
            kind="config",
            label="Configuration",
            status="fail",
            reason="config_snapshot_unavailable",
        )
    try:
        result = snapshot()
    except Exception:
        return DoctorCheck(
            id="config.load",
            kind="config",
            label="Configuration",
            status="fail",
            reason="config_load_failed",
        )
    if not isinstance(result, dict):
        return DoctorCheck(
            id="config.load",
            kind="config",
            label="Configuration",
            status="fail",
            reason="config_snapshot_invalid",
        )
    return DoctorCheck(
        id="config.load",
        kind="config",
        label="Configuration",
        status="ok",
    )


def _check_dependency_egress_guard(projection: dict[str, Any]) -> DoctorCheck:
    """Confirm that this command is running inside the enforced bundle guard."""

    security = projection.get("security")
    guard = security.get("dependency_egress_guard") if isinstance(security, dict) else None
    if not isinstance(guard, dict):
        return DoctorCheck(
            id="security.dependency_egress_guard",
            kind="security",
            label="Dependency egress guard",
            status="fail",
            reason="status_unavailable",
        )
    enforced = guard.get("state") == "enforced" and guard.get("installed") is True and guard.get("active") is True
    return DoctorCheck(
        id="security.dependency_egress_guard",
        kind="security",
        label="Dependency egress guard",
        status="ok" if enforced else "fail",
        reason=None if enforced else "not_enforced",
    )


def collect_doctor_checks(controller: Any, projection: dict[str, Any]) -> list[DoctorCheck]:
    """Collect path-free base-health checks from injected product services."""

    return [_check_temp_directory(), _check_config(controller), _check_dependency_egress_guard(projection)]


def _print_capability_text(projection: dict[str, Any]) -> None:
    """Print a concise human view of the canonical runtime projection."""

    runtime = projection["runtime"]
    counts = projection["counts"]
    print("\n── 运行时能力 ──")
    print(f"  平台：{runtime.get('platform', 'unknown')}")
    security = projection.get("security", {})
    guard = security.get("dependency_egress_guard", {}) if isinstance(security, dict) else {}
    if isinstance(guard, dict):
        print(f"  第三方依赖出站保护：{guard.get('state', 'unknown')}")
        print(f"  外部进程：{guard.get('external_processes', 'unknown')}")
    print(
        "  路由："
        f"{counts.get('available_routes', 0)}/{counts.get('routes', 0)} 可用；"
        f"{counts.get('unavailable_routes', 0)} 不可用"
    )

    unavailable_sources: list[str] = []
    for source in projection["sources"]:
        if not isinstance(source, dict):
            continue
        routes = source.get("routes", [])
        if not isinstance(routes, list):
            continue
        if routes and not any(isinstance(route, dict) and route.get("available") is True for route in routes):
            unavailable_sources.append(str(source.get("id", "unknown")))
    if unavailable_sources:
        print(f"  完全不可用的输入：{', '.join(sorted(unavailable_sources))}")


def register_doctor_parser(subparsers: Any) -> argparse.ArgumentParser:
    """Register the ``doctor`` command."""

    return subparsers.add_parser(
        "doctor",
        parents=[get_common_parser()],
        help=cli_t("cli.help.doctor"),
    )


def execute_doctor(args: argparse.Namespace, controller: Any = None) -> int:
    """Execute diagnostics against one initialized runtime composition."""

    projection = runtime_capability_projection(controller)
    checks = collect_doctor_checks(controller, projection)
    all_ok = all(check.ok for check in checks)
    data = {
        "checks": [check.to_dict() for check in checks],
        "all_ok": all_ok,
        "capability_summary": projection,
    }

    if bool(getattr(args, "json", False)):
        from docwen_cli.presenters.json_presenter import JsonPresenter

        JsonPresenter().present_data("doctor", data, success=True)
    else:
        print("DocWen 环境诊断\n")
        print("── 基础检查 ──")
        for check in checks:
            marker = "[OK]" if check.ok else "[FAIL]"
            reason = f": {check.reason}" if check.reason else ""
            print(f"  {marker}  {check.label} ({check.id}){reason}")
        if all_ok:
            print("\n✓ 基础检查通过")
        else:
            print("\n✗ 基础检查失败")
        _print_capability_text(projection)

    return int(ExitCode.OK if all_ok else ExitCode.UNAVAILABLE)

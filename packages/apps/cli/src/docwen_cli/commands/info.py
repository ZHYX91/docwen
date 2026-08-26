"""Lightweight product, protocol, platform, and capability discovery."""

from __future__ import annotations

import argparse
import platform
import sys
from typing import Any

from docwen_cli.exit_codes import ExitCode
from docwen_cli.gui_control_port import GUI_SETTINGS_SECTIONS
from docwen_cli.parser import get_common_parser
from docwen_cli.presenters.json_presenter import JsonPresenter
from docwen_cli.protocol import PROTOCOL_VERSION
from docwen_core.version import PRODUCT_VERSION


def _capability(
    capability_id: str,
    *,
    platforms: list[str],
    runtime_check_required: bool = False,
    available: bool = True,
    reason: str | None = None,
    details: dict[str, Any] | None = None,
) -> dict[str, Any]:
    current_platform = platform.system().lower()
    current_platform_supported = current_platform in platforms
    effective_available = available and current_platform_supported
    effective_runtime_check = runtime_check_required and effective_available
    state = (
        "unavailable"
        if not effective_available
        else ("runtime_check_required" if effective_runtime_check else "available")
    )
    payload: dict[str, Any] = {
        "id": capability_id,
        "contract_version": 1,
        "state": state,
        "available": effective_available,
        "platforms": platforms,
        "current_platform_supported": current_platform_supported,
        "runtime_check_required": effective_runtime_check,
    }
    effective_reason = reason
    if not current_platform_supported:
        effective_reason = f"Current platform {current_platform!r} is not supported."
    if effective_reason is not None:
        payload["reason"] = effective_reason
    if details is not None:
        payload["details"] = dict(details)
    return payload


def build_info_data() -> dict[str, Any]:
    """Return compiled capability facts without initializing the runtime."""

    return {
        "product": {"name": "DocWen", "version": PRODUCT_VERSION},
        "protocol": {"major": PROTOCOL_VERSION, "envelope": "docwen.cli.v3"},
        "platform": {
            "system": platform.system().lower(),
            "machine": platform.machine().lower(),
            "python": platform.python_version(),
        },
        "capabilities": [
            _capability("cli.inspect", platforms=["windows", "linux", "darwin"]),
            _capability("cli.resources", platforms=["windows", "linux", "darwin"]),
            _capability("cli.schema", platforms=["windows", "linux", "darwin"]),
            _capability("cli.convert", platforms=["windows", "linux"], runtime_check_required=True),
            _capability("cli.validate", platforms=["windows", "linux"], runtime_check_required=True),
            _capability("cli.number.markdown", platforms=["windows", "linux"], runtime_check_required=True),
            _capability("cli.merge", platforms=["windows", "linux"], runtime_check_required=True),
            _capability("cli.split.pdf", platforms=["windows", "linux"], runtime_check_required=True),
            _capability(
                "gui.control",
                platforms=["windows"],
                runtime_check_required=True,
            ),
            _capability(
                "gui.settings",
                platforms=["windows"],
                runtime_check_required=True,
                details={"cold_start": True, "sections": list(GUI_SETTINGS_SECTIONS)},
            ),
        ],
    }


def execute_info(args: argparse.Namespace, controller: Any | None = None) -> int:
    """Present lightweight info; *controller* is deliberately unused."""

    del controller
    data = build_info_data()
    if getattr(args, "json", False):
        JsonPresenter(include_timing=getattr(args, "timing", False)).present_data("info", data)
    else:
        print(f"DocWen {PRODUCT_VERSION}")
        print(f"CLI protocol: {PROTOCOL_VERSION}")
        print(f"Platform: {data['platform']['system']} {data['platform']['machine']}")
        unavailable = [item["id"] for item in data["capabilities"] if not item["available"]]
        if unavailable and not getattr(args, "quiet", False):
            print(f"Unavailable: {', '.join(unavailable)}", file=sys.stderr)
    return int(ExitCode.OK)


def register_info_parser(subparsers: Any) -> argparse.ArgumentParser:
    """Register the ``info`` command."""

    return subparsers.add_parser(
        "info",
        parents=[get_common_parser()],
        help="Show product, protocol, platform, and capability information.",
    )

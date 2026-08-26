"""Reusable protocol fixtures for packaged-candidate verifier tests."""

from __future__ import annotations

from pathlib import Path


def fake_numbering_payload(args: tuple[str, ...]) -> dict[str, object]:
    """Write the requested numbering output and return its protocol payload."""

    output_file = Path(args[args.index("--output") + 1])
    operation = args[args.index("--operation") + 1]
    content = (
        "# 一、已有标题\n\n## （一）二级标题\n\n正文。\n"
        if operation == "add"
        else "# 已有标题\n\n## 二级标题\n\n正文。\n"
    )
    output_file.write_text(content, encoding="utf-8")
    return {
        "protocol_version": 3,
        "success": True,
        "command": "number",
        "data": {"output": str(output_file)},
        "error": None,
    }


def fake_dependency_egress_guard() -> dict[str, object]:
    """Return the exact frozen-process dependency-egress status contract."""

    return {
        "state": "enforced",
        "installed": True,
        "active": True,
        "scope": "docwen_python_process",
        "policy": "deny_dns_and_ip",
        "mechanism": "cpython_audit_hook",
        "bootstrap": "pyinstaller_runtime_hook",
        "local_transports": ["windows_named_pipe", "unix_domain_socket"],
        "external_processes": "not_managed",
    }


__all__ = ["fake_dependency_egress_guard", "fake_numbering_payload"]

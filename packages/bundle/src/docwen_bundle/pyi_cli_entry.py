"""PyInstaller entry point for DocWen CLI.

This thin wrapper exists solely as a filesystem entry point for PyInstaller.
It delegates to ``docwen_bundle.cli_entry.main`` — the same composition root
used by the ``docwen`` console_script.

Usage (by PyInstaller only)::

    pyinstaller pyi_cli_entry.py --name=DocWenCLI ...

Do NOT use this file as a user-facing entry point. Use the ``docwen``
console script instead.
"""

import json
import multiprocessing
import os
import sys
from pathlib import Path

import docwen_bundle.cli_entry as cli_entry

_MULTIPROCESS_EGRESS_REPORT_ENV = "DOCWEN_TEST_MULTIPROCESS_EGRESS_REPORT"


def _prepare_multiprocessing() -> None:
    """Let frozen child processes enter their multiprocessing bootstrap."""
    multiprocessing.freeze_support()


def _delegate() -> int:
    """Run the bundle CLI entry and return its exit code.

    Exposed as a function so tests can verify delegation without driving
    ``__main__`` execution. Resolves ``cli_entry.main`` at call time so
    monkeypatching ``cli_entry.main`` in tests takes effect.
    """
    return cli_entry.main()


def _write_unmanaged_child_report(report_path: str) -> None:
    """Frozen child target used only by the packaged boundary verifier."""

    from docwen_runtime.security import dependency_egress_guard_status

    audit_probe_allowed = True
    try:
        sys.audit("socket.getaddrinfo", "docwen-unmanaged-child-probe")
    except RuntimeError:
        audit_probe_allowed = False
    Path(report_path).write_text(
        json.dumps(
            {
                "guard": dependency_egress_guard_status().to_dict(),
                "audit_probe_allowed": audit_probe_allowed,
            },
            ensure_ascii=False,
            sort_keys=True,
        ),
        encoding="utf-8",
    )


def _run_multiprocessing_egress_boundary_probe() -> int | None:
    """Verify parent/child guard ownership when requested by release tooling."""

    raw_report_path = os.environ.get(_MULTIPROCESS_EGRESS_REPORT_ENV, "").strip()
    if not raw_report_path:
        return None

    from docwen_runtime.security import NetworkAccessBlockedError, dependency_egress_guard_status

    report_path = Path(raw_report_path)
    report_path.parent.mkdir(parents=True, exist_ok=True)
    child_report_path = report_path.with_suffix(".child.json")

    parent_probe_blocked = False
    try:
        sys.audit("socket.getaddrinfo", "docwen-managed-parent-probe")
    except NetworkAccessBlockedError:
        parent_probe_blocked = True

    context = multiprocessing.get_context("spawn")
    process = context.Process(
        target=_write_unmanaged_child_report,
        args=(str(child_report_path),),
        name="docwen-egress-boundary-probe",
    )
    process.start()
    process.join(timeout=30.0)
    if process.is_alive():
        process.terminate()
        process.join(timeout=5.0)

    child_payload: object = None
    if child_report_path.is_file():
        child_payload = json.loads(child_report_path.read_text(encoding="utf-8"))
        child_report_path.unlink(missing_ok=True)

    payload = {
        "parent_guard": dependency_egress_guard_status().to_dict(),
        "parent_audit_probe_blocked": parent_probe_blocked,
        "child_exit_code": process.exitcode,
        "child": child_payload,
    }
    report_path.write_text(
        json.dumps(payload, ensure_ascii=False, sort_keys=True),
        encoding="utf-8",
    )
    child_guard = child_payload.get("guard") if isinstance(child_payload, dict) else None
    child_audit_probe_allowed = child_payload.get("audit_probe_allowed") if isinstance(child_payload, dict) else None
    child_ok = (
        process.exitcode == 0
        and isinstance(child_guard, dict)
        and child_guard.get("state") == "not_installed"
        and child_audit_probe_allowed is True
    )
    parent_status = payload["parent_guard"]
    parent_ok = (
        parent_probe_blocked
        and isinstance(parent_status, dict)
        and parent_status.get("state") == "enforced"
        and parent_status.get("bootstrap") == "pyinstaller_runtime_hook"
    )
    return 0 if parent_ok and child_ok else 1


if __name__ == "__main__":
    _prepare_multiprocessing()
    _probe_exit_code = _run_multiprocessing_egress_boundary_probe()
    sys.exit(_delegate() if _probe_exit_code is None else _probe_exit_code)

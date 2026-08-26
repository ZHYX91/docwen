from __future__ import annotations

import json
import os
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

import pytest
from _pytest.stash import StashKey

from tests.support.marker_taxonomy import PRIMARY_TEST_MARKERS
from tests.support.subprocess_runner import clear_subprocess_run_records, load_subprocess_run_records

from .collection import _relative_collection_path
from .dependencies import (
    _COLLECTION_DEPENDENCY_STATUS,
    _NOT_COLLECTED_BY_REASON,
    _REPORT_DIR_ENV,
    _ROOT,
    _RUNTIME_SKIP_DEPENDENCY_STATUS,
)

_MISSING_PRIMARY_MARKER_ITEMS: list[dict[str, Any]] = []
_SLOW_TEST_ITEMS: list[dict[str, Any]] = []
_WORKER_MISSING_MARKER_KEY = "docwen_missing_primary_marker_items"
_WORKER_MISSING_MARKER_STASH_KEY: StashKey[list[dict[str, Any]]] = StashKey()


def _format_dependency_status(statuses: dict[str, bool]) -> str:
    return ", ".join(f"{name}={'ok' if available else 'missing'}" for name, available in statuses.items())


def _extract_skip_reason(report: Any) -> str:
    longrepr = getattr(report, "longrepr", None)
    if isinstance(longrepr, tuple) and len(longrepr) >= 3:
        return str(longrepr[2])

    reprcrash = getattr(longrepr, "reprcrash", None)
    if reprcrash is not None and getattr(reprcrash, "message", None):
        return str(reprcrash.message)

    longreprtext = getattr(report, "longreprtext", "")
    if longreprtext:
        return longreprtext.strip().splitlines()[-1]

    return "skip reason unavailable"


def _skip_location_payload(report: Any) -> dict[str, Any]:
    location = getattr(report, "location", None)
    if isinstance(location, tuple) and len(location) >= 2:
        return {
            "path": _relative_report_path(location[0]),
            "line": int(location[1]) + 1,
        }

    longrepr = getattr(report, "longrepr", None)
    if isinstance(longrepr, tuple) and len(longrepr) >= 2:
        return {
            "path": _relative_report_path(longrepr[0]),
            "line": int(longrepr[1]) + 1,
        }

    return {"path": "", "line": 0}


def _xdist_payload(config: Any) -> dict[str, Any]:
    option = getattr(config, "option", None)
    raw_numprocesses = getattr(option, "numprocesses", None)
    maxprocesses = getattr(option, "maxprocesses", None)
    numprocesses = raw_numprocesses
    if isinstance(raw_numprocesses, int) and isinstance(maxprocesses, int) and maxprocesses > 0:
        numprocesses = min(raw_numprocesses, maxprocesses)
    dist = getattr(option, "dist", "no") or "no"
    enabled = raw_numprocesses not in (None, 0, "0", "no")
    return {
        "enabled": enabled,
        "numprocesses": numprocesses if numprocesses is not None else 0,
        "dist": dist,
    }


def _selection_payload(config: Any) -> dict[str, Any]:
    option = getattr(config, "option", None)
    return {
        "markexpr": getattr(option, "markexpr", "") or "",
        "addopts": str(getattr(config, "getini", lambda _name: "")("addopts") or ""),
    }


def _skip_report_payload(skipped_reports: list[Any], config: Any) -> dict[str, Any]:
    items = []
    for report in skipped_reports:
        items.append(
            {
                "nodeid": getattr(report, "nodeid", ""),
                "phase": getattr(report, "when", "unknown"),
                "reason": _extract_skip_reason(report),
                "location": _skip_location_payload(report),
            }
        )

    return {
        "generated_at": datetime.now(UTC).isoformat(),
        "summary": {"count": len(items), "skipped_count": len(items)},
        "pytest": {
            "selection": _selection_payload(config),
            "xdist": _xdist_payload(config),
        },
        "items": items,
    }


def _not_collected_report_payload(config: Any) -> dict[str, Any]:
    items = []
    by_reason = []
    for reason, paths in sorted(_NOT_COLLECTED_BY_REASON.items()):
        ordered_paths = sorted(paths)
        items.extend({"path": path, "reason": reason} for path in ordered_paths)
        by_reason.append(
            {
                "reason": reason,
                "count": len(ordered_paths),
                "paths": ordered_paths,
            }
        )

    return {
        "generated_at": datetime.now(UTC).isoformat(),
        "summary": {
            "count": len(items),
            "dependency_gated_not_collected_count": len(items),
        },
        "pytest": {
            "selection": _selection_payload(config),
            "xdist": _xdist_payload(config),
        },
        "collection_dependencies": dict(sorted(_COLLECTION_DEPENDENCY_STATUS.items())),
        "runtime_skip_dependencies": dict(sorted(_RUNTIME_SKIP_DEPENDENCY_STATUS.items())),
        "items": items,
        "by_reason": by_reason,
    }


def _slow_report_payload(slow_items: list[dict[str, Any]], config: Any) -> dict[str, Any]:
    ordered_items = sorted(slow_items, key=lambda item: item["duration_seconds"], reverse=True)
    threshold_seconds = float(os.environ.get("DOCWEN_PYTEST_SLOW_THRESHOLD_SECONDS", "1.0"))
    return {
        "generated_at": datetime.now(UTC).isoformat(),
        "summary": {"count": len(ordered_items), "slow_threshold_seconds": threshold_seconds},
        "pytest": {
            "selection": _selection_payload(config),
            "xdist": _xdist_payload(config),
        },
        "items": ordered_items,
    }


def _subprocess_report_payload(subprocess_records: list[dict[str, Any]], config: Any) -> dict[str, Any]:
    ordered_items = sorted(subprocess_records, key=lambda item: item["duration_seconds"], reverse=True)
    timeout_count = sum(1 for item in ordered_items if item["timed_out"])
    return {
        "generated_at": datetime.now(UTC).isoformat(),
        "summary": {
            "count": len(ordered_items),
            "timeout_count": timeout_count,
        },
        "pytest": {
            "selection": _selection_payload(config),
            "xdist": _xdist_payload(config),
        },
        "items": ordered_items,
    }


def _missing_marker_report_payload(missing_marker_items: list[dict[str, Any]], config: Any) -> dict[str, Any]:
    ordered_items = sorted(missing_marker_items, key=lambda item: item["nodeid"])
    return {
        "generated_at": datetime.now(UTC).isoformat(),
        "summary": {"count": len(ordered_items), "missing_primary_marker_count": len(ordered_items)},
        "pytest": {
            "selection": _selection_payload(config),
            "xdist": _xdist_payload(config),
        },
        "required_primary_markers": sorted(PRIMARY_TEST_MARKERS),
        "items": ordered_items,
    }


def _report_output_dir() -> Path:
    configured = Path(os.environ.get(_REPORT_DIR_ENV, _ROOT / ".pytest_cache" / "docwen_reports"))
    if not configured.is_absolute():
        configured = (_ROOT / configured).resolve()
    return configured


def _write_json_report(file_path: Path, payload: dict[str, Any]) -> None:
    file_path.parent.mkdir(parents=True, exist_ok=True)
    file_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def _write_visibility_reports(
    skipped_reports: list[Any],
    config: Any,
    *,
    slow_items: list[dict[str, Any]] | None = None,
    subprocess_records: list[dict[str, Any]] | None = None,
    missing_marker_items: list[dict[str, Any]] | None = None,
) -> dict[str, Path]:
    output_dir = _report_output_dir()
    skip_report = output_dir / "skip_report.json"
    not_collected_report = output_dir / "not_collected_report.json"
    slow_report = output_dir / "slow_report.json"
    subprocess_report = output_dir / "subprocess_report.json"
    missing_marker_report = output_dir / "missing_marker_report.json"
    _write_json_report(skip_report, _skip_report_payload(skipped_reports, config))
    _write_json_report(not_collected_report, _not_collected_report_payload(config))
    _write_json_report(slow_report, _slow_report_payload(slow_items or _SLOW_TEST_ITEMS, config))
    _write_json_report(
        subprocess_report,
        _subprocess_report_payload(
            subprocess_records or load_subprocess_run_records(output_dir),
            config,
        ),
    )
    _write_json_report(
        missing_marker_report,
        _missing_marker_report_payload(missing_marker_items or _MISSING_PRIMARY_MARKER_ITEMS, config),
    )
    return {
        "skip_report": skip_report,
        "not_collected_report": not_collected_report,
        "slow_report": slow_report,
        "subprocess_report": subprocess_report,
        "missing_marker_report": missing_marker_report,
    }


def _relative_report_path(path_value: str | Path) -> str:
    path = Path(path_value)
    try:
        return path.resolve().relative_to(_ROOT).as_posix()
    except ValueError:
        return path.as_posix()


def pytest_sessionstart(session: Any) -> None:
    del session
    _MISSING_PRIMARY_MARKER_ITEMS.clear()
    _SLOW_TEST_ITEMS.clear()


def pytest_collection_modifyitems(config: Any, items: list[Any]) -> None:
    _MISSING_PRIMARY_MARKER_ITEMS.clear()
    _SLOW_TEST_ITEMS.clear()

    if not os.environ.get("PYTEST_XDIST_WORKER"):
        clear_subprocess_run_records(_report_output_dir())

    for item in items:
        marker_names = {marker.name for marker in item.iter_markers()}
        primary_markers = sorted(marker_names & PRIMARY_TEST_MARKERS)
        if primary_markers:
            continue
        _MISSING_PRIMARY_MARKER_ITEMS.append(
            {
                "nodeid": item.nodeid,
                "path": _relative_report_path(getattr(item, "fspath", "")),
                "primary_markers": primary_markers,
            }
        )
    config.stash[_WORKER_MISSING_MARKER_STASH_KEY] = list(_MISSING_PRIMARY_MARKER_ITEMS)


def pytest_collection_finish(session: Any) -> None:
    worker_output = getattr(session.config, "workeroutput", None)
    if isinstance(worker_output, dict):
        worker_output[_WORKER_MISSING_MARKER_KEY] = list(session.config.stash.get(_WORKER_MISSING_MARKER_STASH_KEY, []))


@pytest.hookimpl(optionalhook=True)
def pytest_testnodedown(node: Any, error: object | None) -> None:
    del error
    worker_output = getattr(node, "workeroutput", None)
    if not isinstance(worker_output, dict):
        return
    worker_items = worker_output.get(_WORKER_MISSING_MARKER_KEY, [])
    if not isinstance(worker_items, list):
        return

    existing_nodeids = {item.get("nodeid") for item in _MISSING_PRIMARY_MARKER_ITEMS}
    for item in worker_items:
        if not isinstance(item, dict):
            continue
        nodeid = item.get("nodeid")
        if not isinstance(nodeid, str) or not nodeid or nodeid in existing_nodeids:
            continue
        _MISSING_PRIMARY_MARKER_ITEMS.append(dict(item))
        existing_nodeids.add(nodeid)


def pytest_runtest_logreport(report: Any) -> None:
    if report.when != "call":
        return

    threshold_seconds = float(os.environ.get("DOCWEN_PYTEST_SLOW_THRESHOLD_SECONDS", "1.0"))
    duration_seconds = round(float(getattr(report, "duration", 0.0)), 6)
    if duration_seconds < threshold_seconds:
        return

    _SLOW_TEST_ITEMS.append(
        {
            "nodeid": getattr(report, "nodeid", ""),
            "outcome": getattr(report, "outcome", "unknown"),
            "duration_seconds": duration_seconds,
        }
    )


def pytest_terminal_summary(terminalreporter: Any, exitstatus: int, config: Any) -> None:
    del exitstatus

    skipped_reports = list(terminalreporter.stats.get("skipped", []))
    skipped_count = len(skipped_reports)
    not_collected_count = sum(len(paths) for paths in _NOT_COLLECTED_BY_REASON.values())
    missing_collection_dependencies = [
        name for name, available in _COLLECTION_DEPENDENCY_STATUS.items() if not available
    ]
    subprocess_records = load_subprocess_run_records(_report_output_dir())

    terminalreporter.section("skip/not-collected visibility", sep="=")
    terminalreporter.write_line(f"skipped tests: {skipped_count} (details shown by pytest -ra)")
    terminalreporter.write_line(f"dependency-gated not collected paths: {not_collected_count}")
    if not_collected_count == 0:
        terminalreporter.write_line("not collected detail: none in current environment")
    else:
        for reason, paths in sorted(_NOT_COLLECTED_BY_REASON.items()):
            ordered_paths = sorted(paths)
            terminalreporter.write_line(f"{reason}: {len(ordered_paths)} path(s)")
            for rel_path in ordered_paths[:10]:
                terminalreporter.write_line(f"  - {rel_path}")
            remaining = len(ordered_paths) - 10
            if remaining > 0:
                terminalreporter.write_line(f"  ... {remaining} more path(s)")

    terminalreporter.write_line(
        "collection-critical dependencies: " + _format_dependency_status(_COLLECTION_DEPENDENCY_STATUS)
    )
    terminalreporter.write_line(
        "runtime-skip dependencies: " + _format_dependency_status(_RUNTIME_SKIP_DEPENDENCY_STATUS)
    )
    selection = _selection_payload(config)
    xdist = _xdist_payload(config)
    terminalreporter.write_line(
        "selection mark expression: " + (selection["markexpr"] if selection["markexpr"] else "<none>")
    )
    terminalreporter.write_line(
        f"xdist: enabled={xdist['enabled']}, numprocesses={xdist['numprocesses']}, dist={xdist['dist']}"
    )
    if missing_collection_dependencies:
        terminalreporter.write_line(
            "missing collection-critical dependencies shrink the test set: "
            + ", ".join(missing_collection_dependencies)
        )
    else:
        terminalreporter.write_line(
            "all collection-critical dependencies are available; no dependency-driven collection shrink occurred"
        )

    report_paths = _write_visibility_reports(skipped_reports, config)
    terminalreporter.write_line(
        "structured visibility reports: "
        f"skip_report={_relative_collection_path(report_paths['skip_report'])}, "
        f"not_collected_report={_relative_collection_path(report_paths['not_collected_report'])}, "
        f"slow_report={_relative_collection_path(report_paths['slow_report'])}, "
        f"subprocess_report={_relative_collection_path(report_paths['subprocess_report'])}, "
        f"missing_marker_report={_relative_collection_path(report_paths['missing_marker_report'])}"
    )
    terminalreporter.write_line(
        "slow tests / subprocess / marker gaps: "
        f"{len(_SLOW_TEST_ITEMS)} / {len(subprocess_records)} / {len(_MISSING_PRIMARY_MARKER_ITEMS)}"
    )


def pytest_sessionfinish(session: Any, exitstatus: int) -> None:
    """Fail a green governed run when primary-marker debt exceeds its ratchet."""

    raw_limit = os.environ.get("DOCWEN_PYTEST_MAX_MISSING_PRIMARY_MARKERS", "").strip()
    if not raw_limit:
        return
    terminalreporter = session.config.pluginmanager.get_plugin("terminalreporter")
    try:
        limit = int(raw_limit)
    except ValueError:
        if terminalreporter is not None:
            terminalreporter.write_sep("=", f"invalid primary marker debt limit: {raw_limit!r}")
        session.exitstatus = pytest.ExitCode.USAGE_ERROR
        return
    if limit < 0:
        if terminalreporter is not None:
            terminalreporter.write_sep("=", f"invalid primary marker debt limit: {limit}")
        session.exitstatus = pytest.ExitCode.USAGE_ERROR
        return
    if len(_MISSING_PRIMARY_MARKER_ITEMS) > limit and exitstatus == pytest.ExitCode.OK:
        if terminalreporter is not None:
            terminalreporter.write_sep(
                "=",
                f"primary marker debt exceeded: actual={len(_MISSING_PRIMARY_MARKER_ITEMS)}, limit={limit}",
            )
        session.exitstatus = pytest.ExitCode.TESTS_FAILED


__all__ = [
    "pytest_collection_finish",
    "pytest_collection_modifyitems",
    "pytest_runtest_logreport",
    "pytest_sessionfinish",
    "pytest_sessionstart",
    "pytest_terminal_summary",
    "pytest_testnodedown",
]

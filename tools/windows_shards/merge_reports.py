"""Build the standard Windows QA artifact after both coverage shards finish."""

from __future__ import annotations

import argparse
import json
import shutil
import sys
import tempfile
from pathlib import Path
from typing import Any

from tools import qa_reports
from tools.qa import PYTEST_PRIMARY_MARKER_DEBT_LIMIT
from tools.windows_shards import compare


def _visibility(name: str, left: dict[str, Any], right: dict[str, Any]) -> dict[str, Any]:
    result = {**left, "generated_at": max(left["generated_at"], right["generated_at"])}
    if "pytest" in left or "pytest" in right:
        if left["pytest"]["selection"] != right["pytest"]["selection"]:
            raise ValueError(f"shards used different pytest selections for {name}")
        # Each runner resolves its own auto worker count. Retain both sources.
        workers = {"apps-plugins": left["pytest"]["xdist"], "remaining": right["pytest"]["xdist"]}
        xdist = {**workers["apps-plugins"], "numprocesses": sum(item["numprocesses"] for item in workers.values())}
        limits = [item.get("maxprocesses") for item in workers.values()]
        xdist["maxprocesses"] = sum(limits) if all(isinstance(limit, int) for limit in limits) else None
        result["pytest"] = {**left["pytest"], "xdist": xdist, "shards": workers}

    # Full collection precedes partitioning on both runners. Count its
    # exclusions and primary-marker audit once, while checking they agree.
    if name in {"not_collected_report.json", "missing_marker_report.json"}:
        fields = left.keys() | right.keys()
        if any(left[field] != right[field] for field in fields - {"generated_at", "pytest"}):
            raise ValueError(f"the two full collections disagree on {name}")
        count = left["summary"]["count"]
        if count != len(left["items"]):
            raise ValueError(f"inconsistent collection count in {name}")
        if name == "missing_marker_report.json" and count > PYTEST_PRIMARY_MARKER_DEBT_LIMIT:
            raise ValueError("required primary markers missing in the Windows collection")
        if name == "not_collected_report.json" and (count or not all(left["collection_dependencies"].values())):
            raise ValueError("the complete Windows suite requires all collection-critical dependencies")
        return result

    items = left["items"] + right["items"]
    summary = dict(left["summary"], count=len(items))
    if name == "slow_report.json":
        if summary["slow_threshold_seconds"] != right["summary"]["slow_threshold_seconds"]:
            raise ValueError("shards used different slow-test thresholds")
        items.sort(key=lambda item: item["duration_seconds"], reverse=True)
    elif name == "subprocess_report.json":
        items.sort(key=lambda item: item["duration_seconds"], reverse=True)
        summary["timeout_count"] = sum(bool(item["timed_out"]) for item in items)
    elif name == "skip_report.json":
        items.sort(key=lambda item: (item["nodeid"], item["phase"]))
        summary["skipped_count"] = len(items)
    else:
        raise ValueError(f"unsupported visibility report {name}")
    return {**result, "summary": summary, "items": items}


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("apps_plugins", type=Path)
    parser.add_argument("remaining", type=Path)
    parser.add_argument("output", type=Path)
    args = parser.parse_args(argv)
    output = args.output.resolve()
    with tempfile.TemporaryDirectory(prefix="docwen-windows-coverage-") as temporary:
        runtime = Path(temporary)
        code = compare.main(["merge", str(args.apps_plugins), str(args.remaining), str(runtime)])
        if code == 0:
            reports = runtime / "reports"
            reports.mkdir()
            try:
                for name in qa_reports.VISIBILITY_REPORTS:
                    sources = [
                        json.loads((directory / name).read_text(encoding="utf-8"))
                        for directory in (args.apps_plugins, args.remaining)
                    ]
                    merged = _visibility(name, *sources)
                    (reports / name).write_text(
                        json.dumps(merged, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
                    )
            except (OSError, KeyError, ValueError) as error:
                sys.stderr.write(f"could not combine Windows visibility reports: {error}\n")
                code = 1
        else:
            sys.stderr.write(
                "Both Windows shards must succeed from the same run and attempt. Re-run all jobs if needed.\n"
            )
        # Use the normal export to retain the historical artifact layout,
        # per-report digests and the coverage XML, including failure evidence.
        qa_reports.export_reports(runtime, output, coverage=True, exit_code=code)
        for name in ("merge.json", "coverage-gates.log"):
            if (runtime / name).is_file():
                shutil.copyfile(runtime / name, output / name)
    return code


if __name__ == "__main__":
    raise SystemExit(main())

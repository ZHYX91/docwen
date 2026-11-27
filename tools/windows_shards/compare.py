"""Verify a full-suite control against two Windows shards; merge gates can run independently."""

from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
import tempfile
from functools import cache
from pathlib import Path
from typing import Any

from coverage import CoverageData
from coverage.exceptions import CoverageException
from tools import qa_reports
from tools.windows_shards import GROUPS, selected

RAW_DATABASE = "coverage-data.sqlite"


def _read(directory: Path, name: str) -> dict[str, Any]:
    return json.loads((directory / name).read_text(encoding="utf-8"))


def _require(condition: bool, message: str) -> None:
    if not condition:
        raise ValueError(message)


def _unique(items: list[str], name: str) -> set[str]:
    values = set(items)
    _require(len(items) == len(values), f"duplicate {name}")
    return values


def _difference(expected: set, actual: set) -> dict:
    return {"missing": sorted(expected - actual), "extra": sorted(actual - expected)}


def _same(expected: set, actual: set, name: str) -> None:
    if expected != actual:
        difference = _difference(expected, actual)
        sample = {side: values[:8] for side, values in difference.items()}
        raise ValueError(f"{name} differs: {sample}")


def _exclusions(directory: Path) -> dict[str, Any]:
    report = _read(directory, "not_collected_report.json")
    return {field: report[field] for field in ("items", "collection_dependencies", "runtime_skip_dependencies")}


def _skips(directory: Path) -> list[tuple[str, str, str]]:
    return sorted(
        (item["nodeid"], item["phase"], item["reason"]) for item in _read(directory, "skip_report.json")["items"]
    )


def check_inventories(directories: dict[str, Path]) -> dict[str, int]:
    inventories = {}
    for group, directory in directories.items():
        evidence = _read(directory, "node-evidence.json")
        _require(evidence["group"] == group, f"{group}: unexpected shard identity")
        _require(
            evidence["exitstatus"] == 0 and not evidence["errors"], f"{group}: incomplete run: {evidence['errors']}"
        )
        inventory = evidence["inventory"]
        deselected = inventory["deselected"]
        _require(
            deselected == [],
            f"{group}: tests deselected before the full-suite partition: {deselected[:8]}",
        )
        full = _unique(inventory["full"], f"{group} full node ID")
        chosen = _unique(inventory["selected"], f"{group} selected node ID")
        _require(bool(chosen), f"{group}: empty test selection")
        _same({nodeid for nodeid in full if selected(group, nodeid)}, chosen, f"{group} selection")
        outcomes = evidence["outcomes"]
        _same(chosen, set(outcomes), f"{group} completed nodes")
        for nodeid, phases in outcomes.items():
            setup = phases.get("setup")
            expected = {"setup", "call", "teardown"} if setup == "passed" else {"setup", "teardown"}
            _require(
                set(phases) == expected
                and all(value in {"passed", "skipped", "xskipped", "xpassed"} for value in phases.values()),
                f"{group}: incomplete or failed phases for {nodeid}: {phases}",
            )
        inventories[group] = {
            "full": full,
            "chosen": chosen,
            "outcomes": outcomes,
            "identity": evidence["identity"],
            "exclusions": _exclusions(directory),
            "skips": _skips(directory),
        }

    first = inventories["apps-plugins"]
    for group, inventory in inventories.items():
        _same(first["full"], inventory["full"], f"{group} full inventory")
        _require(first["exclusions"] == inventory["exclusions"], f"{group}: collection exclusions differ")
        _require(first["identity"] == inventory["identity"], f"{group}: checkout or run attempt differs")
    other = inventories["remaining"]
    _require(first["chosen"].isdisjoint(other["chosen"]), "shards overlap")
    _same(first["full"], first["chosen"] | other["chosen"], "union of completed shards")
    if "control" in inventories:
        control = inventories["control"]
        split = first["outcomes"] | other["outcomes"]
        changed = [nodeid for nodeid, value in split.items() if value != control["outcomes"][nodeid]]
        _require(not changed, f"outcomes differ for {len(changed)} nodes: {changed[:8]}")
        _require(sorted(first["skips"] + other["skips"]) == control["skips"], "skip phases or reasons differ")
    return {group: len(inventory["chosen"]) for group, inventory in inventories.items()}


def _raw_database(directory: Path) -> CoverageData:
    path = directory / RAW_DATABASE
    _require(path.is_file(), f"coverage data missing: {path}")
    data = CoverageData(basename=path)
    data.read()
    _require(data.has_arcs() and bool(data.measured_files()), f"branch coverage data missing: {path}")
    return data


def combined_data(directories: list[Path], destination: Path) -> CoverageData:
    repo = Path.cwd().resolve()

    @cache
    def source_path(name: str) -> str:
        path = (repo / name.replace("\\", "/")).resolve(strict=True)
        return str(path.relative_to(repo))

    _require(not destination.exists(), f"refusing to reuse existing coverage database: {destination}")
    data = CoverageData(basename=destination)
    for directory in directories:
        data.update(_raw_database(directory), map_path=source_path)
    data.write()
    return data


def coverage_differences(control: CoverageData, combined: CoverageData) -> dict:
    expected = set(control.measured_files())
    actual = set(combined.measured_files())
    result = {"files": _difference(expected, actual)} if expected != actual else {}
    for path in sorted(expected & actual):
        details = {}
        for label, before, after in (
            ("lines", control.lines(path), combined.lines(path)),
            ("arcs", control.arcs(path), combined.arcs(path)),
        ):
            left, right = set(before or []), set(after or [])
            if left != right:
                details[label] = _difference(left, right)
        if details:
            result[path] = details
    return result


def _coverage_gates(output: Path) -> dict[str, int]:
    env = {**os.environ, "COVERAGE_FILE": str(output / RAW_DATABASE)}
    commands = [
        ["-m", "coverage", "xml", "--rcfile=pyproject.toml", "-o", str(output / "coverage.xml")],
        *qa_reports.coverage_checks(output),
    ]
    results = {}
    with (output / "coverage-gates.log").open("w", encoding="utf-8") as log:
        for command in commands:
            result = subprocess.run([sys.executable, *command], env=env, stdout=log, stderr=subprocess.STDOUT)
            results[" ".join(command)] = result.returncode
    return results


def main(argv: list[str] | None = None) -> int:
    values = list(sys.argv[1:] if argv is None else argv)
    merge = bool(values and values[0] == "merge")
    parser = argparse.ArgumentParser(description=__doc__)
    groups = GROUPS[1:] if merge else GROUPS
    parser.add_argument("inputs", nargs=len(groups), type=Path, metavar="INPUT")
    parser.add_argument("output", type=Path)
    args = parser.parse_args(values[1:] if merge else values)
    directories = dict(zip(groups, args.inputs, strict=True))
    output = args.output.resolve()
    output.mkdir(parents=True, exist_ok=True)
    result: dict[str, Any] = {}
    status = 1
    try:
        result["test_counts"] = check_inventories(directories)
        shards = [directories[group] for group in GROUPS[1:]]
        if merge:
            combined_data(shards, output / RAW_DATABASE)
            result["coverage_gates"] = _coverage_gates(output)
            _require(not any(result["coverage_gates"].values()), "coverage gates failed; see coverage-gates.log")
        else:
            with tempfile.TemporaryDirectory() as temporary:
                control = combined_data([directories["control"]], Path(temporary) / "control.sqlite")
                combined = combined_data(shards, Path(temporary) / "shards.sqlite")
                result["coverage_differences"] = coverage_differences(control, combined)
            _require(not result["coverage_differences"], "executed coverage files, lines or arcs differ")
        status = 0
    except (CoverageException, KeyError, TypeError, ValueError, OSError) as error:
        result["error"] = str(error)
        sys.stderr.write(f"{error}\n")
    report = output / ("merge.json" if merge else "comparison.json")
    report.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return status


if __name__ == "__main__":
    raise SystemExit(main())

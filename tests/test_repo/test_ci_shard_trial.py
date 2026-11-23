from __future__ import annotations

import json
from pathlib import Path

import pytest
from coverage import CoverageData
from tools.windows_shards import GROUPS, selected
from tools.windows_shards.compare import check_inventories, coverage_differences

pytestmark = pytest.mark.unit

APP = "packages/apps/gui/tests/test_dialog.py::test_dialog"
NEW = "packages/future/tests/test_new.py::test_new"


def _inputs(tmp_path: Path) -> dict[str, Path]:
    inputs = {}
    for group in GROUPS:
        directory = tmp_path / group
        directory.mkdir()
        inputs[group] = directory
        chosen = [node for node in (APP, NEW) if selected(group, node)]
        receipt = {
            "group": group,
            "errors": [],
            "identity": {"checkout": "commit", "run": "1", "attempt": "1"},
            "exitstatus": 0,
            "inventory": {"full": [APP, NEW], "selected": chosen, "deselected": []},
            "outcomes": {node: {"setup": "passed", "call": "passed", "teardown": "passed"} for node in chosen},
        }
        (directory / "node-evidence.json").write_text(json.dumps(receipt), encoding="utf-8")
        (directory / "not_collected_report.json").write_text(
            json.dumps({"items": [], "collection_dependencies": {}, "runtime_skip_dependencies": {}}), encoding="utf-8"
        )
        (directory / "skip_report.json").write_text('{"items":[]}', encoding="utf-8")
    return inputs


def test_full_union_preserves_new_test_paths(tmp_path: Path) -> None:
    assert check_inventories(_inputs(tmp_path)) == {"control": 2, "apps-plugins": 1, "remaining": 1}


@pytest.mark.parametrize(
    "mutation",
    ["duplicate", "wrong-shard", "unfinished", "skip", "skip-reason", "excluded", "wrong-checkout", "worker-error"],
)
def test_incomplete_or_changed_evidence_cannot_pass(tmp_path: Path, mutation: str) -> None:
    directories = _inputs(tmp_path)
    directory = directories["remaining"]
    path = directory / "node-evidence.json"
    receipt = json.loads(path.read_text(encoding="utf-8"))
    if mutation == "duplicate":
        receipt["inventory"]["selected"].append(NEW)
    elif mutation == "wrong-shard":
        receipt["inventory"]["selected"].append(APP)
    elif mutation == "unfinished":
        del receipt["outcomes"][NEW]["teardown"]
    elif mutation == "skip":
        receipt["outcomes"][NEW] = {"setup": "skipped", "teardown": "passed"}
    elif mutation == "wrong-checkout":
        receipt["identity"]["checkout"] = "another"
    elif mutation == "worker-error":
        receipt["errors"].append("duplicate call")
    elif mutation == "excluded":
        receipt["inventory"]["deselected"].append("different")
    else:
        (directory / "skip_report.json").write_text(
            json.dumps({"items": [{"nodeid": NEW, "phase": "setup", "reason": "changed"}]}), encoding="utf-8"
        )
    path.write_text(json.dumps(receipt), encoding="utf-8")
    with pytest.raises(ValueError):
        check_inventories(directories)


def test_matching_external_deselections_do_not_make_a_full_windows_run(tmp_path: Path) -> None:
    directories = _inputs(tmp_path)
    directories.pop("control")  # Production compares the two independent shard collections.
    for directory in directories.values():
        path = directory / "node-evidence.json"
        receipt = json.loads(path.read_text(encoding="utf-8"))
        receipt["inventory"]["deselected"].append("tests/test_lost.py::test_lost")
        path.write_text(json.dumps(receipt), encoding="utf-8")
    with pytest.raises(ValueError, match=r"deselected before the full-suite partition: .*test_lost"):
        check_inventories(directories)


def test_platform_skip_is_a_selected_and_completed_node(tmp_path: Path) -> None:
    directories = _inputs(tmp_path)
    for group in ("control", "remaining"):
        directory = directories[group]
        path = directory / "node-evidence.json"
        receipt = json.loads(path.read_text(encoding="utf-8"))
        receipt["outcomes"][NEW] = {"setup": "skipped", "teardown": "passed"}
        path.write_text(json.dumps(receipt), encoding="utf-8")
        (directory / "skip_report.json").write_text(
            json.dumps({"items": [{"nodeid": NEW, "phase": "setup", "reason": "requires macOS"}]}), encoding="utf-8"
        )
    assert check_inventories(directories) == {"control": 2, "apps-plugins": 1, "remaining": 1}


def test_identical_hit_lines_can_have_different_arcs() -> None:
    control, split = CoverageData(no_disk=True), CoverageData(no_disk=True)
    control.add_arcs({"x.py": [(1, 2), (2, 3)]})
    split.add_arcs({"x.py": [(1, 3), (2, 3)]})
    assert coverage_differences(control, split) == {"x.py": {"arcs": {"missing": [(1, 2)], "extra": [(1, 3)]}}}

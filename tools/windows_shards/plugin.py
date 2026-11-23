"""Collect the unfiltered inventory on each worker and actual test phases on the controller."""

from __future__ import annotations

import json
import os
import subprocess
from collections.abc import Generator, Sequence
from pathlib import Path
from typing import Any

import pytest
from tools.windows_shards import GROUPS, selected


def pytest_configure(config: pytest.Config) -> None:
    group = os.environ.get("DOCWEN_PYTEST_SHARD", "")
    output = Path(os.environ.get("DOCWEN_PYTEST_SHARD_OUTPUT", ""))
    if group not in GROUPS or not output.is_absolute():
        raise pytest.UsageError("shard evidence requires a known group and an absolute output directory")
    config.pluginmanager.register(ShardEvidence(group, output), "docwen-shard-evidence")


class ShardEvidence:
    def __init__(self, group: str, output: Path) -> None:
        self.group = group
        self.output = output
        self.inventory: dict[str, list[str]] | None = None
        self.external_deselections: list[str] = []
        self.partitioning = False
        self.outcomes: dict[str, dict[str, str]] = {}
        self.errors: list[str] = []

    def pytest_deselected(self, items: Sequence[pytest.Item]) -> None:
        if not self.partitioning:
            self.external_deselections.extend(item.nodeid for item in items)

    @pytest.hookimpl(wrapper=True, tryfirst=True)
    def pytest_collection_modifyitems(self, config: pytest.Config, items: list[pytest.Item]) -> Generator[None]:
        # Wait for the ordinary suite's selectors before partitioning the final list.
        yield
        self.inventory = {"full": [item.nodeid for item in items], "deselected": self.external_deselections}
        excluded = [item for item in items if not selected(self.group, item.nodeid)]
        if excluded:
            self.partitioning = True
            try:
                config.hook.pytest_deselected(items=excluded)
            finally:
                self.partitioning = False
            items[:] = [item for item in items if selected(self.group, item.nodeid)]

    def pytest_collection_finish(self, session: pytest.Session) -> None:
        if self.inventory is None:
            return  # The xdist controller does not collect tests.
        self.inventory["selected"] = [item.nodeid for item in session.items]
        workeroutput = getattr(session.config, "workeroutput", None)
        if workeroutput is not None:
            workeroutput["shard_inventory"] = self.inventory

    @pytest.hookimpl(optionalhook=True)
    def pytest_testnodedown(self, node: Any, error: object | None) -> None:
        inventory = getattr(node, "workeroutput", {}).get("shard_inventory")
        if error is not None or inventory is None:
            self.errors.append(f"{node.gateway.id} exited without a clean collection receipt")
        elif self.inventory is None:
            self.inventory = inventory
        elif self.inventory != inventory:
            self.errors.append(f"{node.gateway.id} collected a different inventory")

    def pytest_runtest_logreport(self, report: pytest.TestReport) -> None:
        phases = self.outcomes.setdefault(report.nodeid, {})
        if report.when in phases:
            self.errors.append(f"duplicate {report.when} phase for {report.nodeid}")
        outcome = report.outcome
        if hasattr(report, "wasxfail"):
            outcome = f"x{outcome}"
        phases[report.when] = outcome

    def pytest_sessionfinish(self, session: pytest.Session, exitstatus: int) -> None:
        if hasattr(session.config, "workerinput"):
            return
        self.output.mkdir(parents=True, exist_ok=True)
        revision = subprocess.run(
            ["git", "rev-parse", "HEAD"], cwd=session.config.rootpath, capture_output=True, text=True, timeout=10
        )
        if revision.returncode and os.environ.get("GITHUB_ACTIONS") == "true":
            self.errors.append("could not identify the checked-out revision")
        receipt = {
            "identity": {
                "checkout": revision.stdout.strip() if revision.returncode == 0 else None,
                "run": os.environ.get("GITHUB_RUN_ID"),
                "attempt": os.environ.get("GITHUB_RUN_ATTEMPT"),
            },
            "group": self.group,
            "exitstatus": exitstatus,
            "errors": self.errors,
            "inventory": self.inventory,
            "outcomes": self.outcomes,
        }
        with (self.output / "node-evidence.json").open("x", encoding="utf-8") as stream:
            json.dump(receipt, stream, ensure_ascii=False)

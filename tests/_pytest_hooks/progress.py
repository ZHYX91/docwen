"""Controller-side progress for QA runs that export reports."""

from collections.abc import Generator
from time import monotonic
from typing import Any

import pytest

_ACTIVE_WORKERS = pytest.StashKey[set[str]]()
_TEST_LOOP_END = pytest.StashKey[float]()


def _progress(config: pytest.Config, message: str) -> None:
    reporter = config.pluginmanager.get_plugin("terminalreporter")
    if reporter is not None:
        reporter.write(f"\n[qa] {message}\n", flush=True)


@pytest.hookimpl(optionalhook=True)
def pytest_configure_node(node: Any) -> None:
    node.config.stash.setdefault(_ACTIVE_WORKERS, set()).add(node.gateway.id)


@pytest.hookimpl(optionalhook=True)
def pytest_testnodedown(node: Any, error: object | None) -> None:
    workers = node.config.stash[_ACTIVE_WORKERS]
    workers.discard(node.gateway.id)
    state = "failed" if error else "exited"
    _progress(node.config, f"worker {node.gateway.id} {state}; active workers: {', '.join(sorted(workers)) or 'none'}")


@pytest.hookimpl(wrapper=True, trylast=True)
def pytest_runtestloop(session: pytest.Session) -> Generator[None, Any, Any]:
    # The inner wrapper resumes before pytest-cov combines data and writes its reports.
    try:
        return (yield)
    finally:
        if not hasattr(session.config, "workerinput"):
            session.config.stash[_TEST_LOOP_END] = monotonic()
            coverage = session.config.getoption("cov_source", default=[]) and not session.config.getoption(
                "no_cov", default=False
            )
            detail = " (including coverage)" if coverage else ""
            _progress(session.config, f"test execution ended; pytest finalizing{detail}")


@pytest.hookimpl(tryfirst=True)
def pytest_sessionfinish(session: pytest.Session) -> None:
    started = session.config.stash.get(_TEST_LOOP_END, None)
    if started is not None:
        _progress(session.config, f"post-test hooks before session finish: {monotonic() - started:.2f}s")

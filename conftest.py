"""Root pytest wiring."""

from __future__ import annotations

import hashlib
from pathlib import Path

import pytest
from tests._pytest_hooks.collection import pytest_ignore_collect  # noqa: F401
from tests._pytest_hooks.reporting import (  # noqa: F401
    pytest_collection_finish,
    pytest_collection_modifyitems,
    pytest_runtest_logreport,
    pytest_sessionfinish,
    pytest_sessionstart,
    pytest_terminal_summary,
    pytest_testnodedown,
)
from tests._pytest_hooks.unraisable import pytest_configure  # noqa: F401

pytest_plugins = ("tests.support.runtime", "tests.support.qt")


@pytest.fixture
def tmp_path(request: pytest.FixtureRequest, tmp_path_factory: pytest.TempPathFactory) -> Path:
    """Keep Windows test paths short without changing their physical identity."""
    digest = hashlib.sha256(request.node.nodeid.encode("utf-8")).hexdigest()[:12]
    return tmp_path_factory.mktemp(f"t{digest}")

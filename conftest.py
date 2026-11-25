"""Root pytest wiring."""

from __future__ import annotations

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

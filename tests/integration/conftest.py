"""Conftest for repo-level integration tests.

Exposes the session-scoped ``round_trip_runtime`` fixture — a real,
fully-wired runtime built via the production ``create_runtime_port``
factory — so cross-plugin round-trip tests can drive the full chain
without deep-importing any plugin's internals.

The round-trip leg functions live in ``_round_trip_helper`` (imported
by tests); this file holds only the pytest fixture, because pytest
collects fixtures from conftest modules, not arbitrary helpers.
"""

from __future__ import annotations

from typing import Any

import pytest


@pytest.fixture(scope="session")
def round_trip_runtime() -> Any:
    """A real, fully-wired runtime for cross-plugin round-trip tests.

    Built via the production ``create_runtime_port`` factory so that
    every default plugin (markdown, document, and the rest) is
    registered exactly as in CLI/GUI. Session-scoped because plugin
    registration is expensive and stateless across conversions.
    """
    from docwen_bundle.runtime_factory import create_runtime_port

    return create_runtime_port()

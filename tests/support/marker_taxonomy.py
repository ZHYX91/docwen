from __future__ import annotations

PRIMARY_TEST_MARKERS = frozenset(
    {
        "unit",
        "contract",
        "integration",
        "gui",
        "gui_smoke",
        "e2e",
        "packaged",
        "host",
    }
)

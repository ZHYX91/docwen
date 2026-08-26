from __future__ import annotations

from pathlib import Path

import pytest

pytestmark = pytest.mark.contract

ROOT = Path(__file__).resolve().parents[2]


def test_reset_drains_published_slots_under_the_global_cache_boundary() -> None:
    source = (ROOT / "packages/core/src/docwen_core/text/ocr.py").read_text(encoding="utf-8")
    reset = source[source.index("def reset_ocr()") :]

    assert "with _ocr_lock:" in reset
    assert "slots = tuple(_ocr_instances.values())" in reset
    assert "slot.invocation_lock.acquire()" in reset
    assert reset.index("slot.invocation_lock.acquire()") < reset.index("_ocr_instances.clear()")
    assert "for invocation_lock in reversed(acquired_locks):" in reset


def test_active_call_and_lookup_boundary_regressions_remain_owned() -> None:
    tests = (ROOT / "packages/core/tests/test_ocr_request_isolation.py").read_text(encoding="utf-8")

    assert "test_reset_ocr_waits_for_active_invocation_before_clearing_cache" in tests
    assert "test_reset_ocr_drains_all_active_slots_and_blocks_new_lookup" in tests
    assert "test_reset_ocr_waits_for_in_flight_initialization_before_clearing_cache" in tests
    assert "test_run_ocr_allows_different_cached_engines_to_run_in_parallel" in tests

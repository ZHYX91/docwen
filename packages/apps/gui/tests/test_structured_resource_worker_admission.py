"""Corrupt frozen bundles after completed admission and preserve valid siblings."""

from pathlib import Path

import pytest
from packages.apps.gui.tests import test_structured_resource_batch_admission as cases

pytestmark = [pytest.mark.gui, pytest.mark.pr_gate, pytest.mark.release_gate]


@pytest.mark.parametrize(
    "fault",
    [
        "intact",
        "after-pending-marker-missing",
        "after-pending-marker-tampered",
        "after-pending-partial",
        "after-pending-marker-missing-empty",
        "after-pending-partial-empty",
        "after-pending-main-missing",
        "after-pending-main-tampered",
        "after-pending-resource-missing",
        "after-pending-resource-tampered",
        "missing-frozen-facts",
        "duplicate-frozen-facts",
    ],
)
def test_worker_isolates_damage_after_completed_admission(qapp, qtbot, tmp_path, monkeypatch, fault):
    original_create = cases.create_main_window
    barrier_observed = []

    def instrumented_create(**kwargs):
        window = original_create(**kwargs)
        if fault.endswith("frozen-facts"):
            original_batch = window._requests.batch

            def corrupt_frozen_facts(**build_kwargs):
                request, context = original_batch(**build_kwargs)
                bundles = context["_clipboard_bundles"]
                assert len(bundles) == 2
                context["_clipboard_bundles"] = bundles[1:] if fault.startswith("missing") else (bundles[0], *bundles)
                return request, context

            monkeypatch.setattr(window._requests, "batch", corrupt_frozen_facts)
        original_admit = window._workflow._admit_many

        def after_completed_admission(parent_request, requests, context, **admit_kwargs):
            assert "_clipboard_bundles" not in context
            admitted = original_admit(parent_request, requests, context, **admit_kwargs)
            assert admitted is True
            assert admit_kwargs["invalid_indices"] == set()
            assert len(requests) == 2
            store = window._clipboard_store_for_paste()
            first = store.bundle(requests[0].input_refs[0].path)
            assert first is not None and store.snapshot_available(first.main.path)
            if fault.startswith("after-pending"):
                case = fault.removesuffix("-empty")
                if case.endswith("partial"):
                    Path(first.root_path, ".bundle.partial").write_bytes(b"controlled-incomplete")
                else:
                    if "marker" in case:
                        member = Path(first.marker_path)
                    elif "main" in case:
                        member = Path(first.main.path)
                    else:
                        member = Path(first.resources[0].path)
                    if case.endswith("missing"):
                        member.unlink()
                    else:
                        original = member.read_bytes()
                        member.write_bytes(bytes([original[0] ^ 1]) + original[1:])
                assert not store.snapshot_available(first.main.path)
            barrier_observed.append(True)
            return admitted

        monkeypatch.setattr(window._workflow, "_admit_many", after_completed_admission)
        return window

    monkeypatch.setattr(cases, "create_main_window", instrumented_create)
    cases.test_real_main_window_batch_keeps_valid_sibling_before_runtime(qapp, qtbot, tmp_path, monkeypatch, fault)
    assert barrier_observed == [True]

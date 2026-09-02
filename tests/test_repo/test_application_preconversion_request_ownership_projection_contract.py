from __future__ import annotations

from pathlib import Path

import pytest
from tools.validation.source_family import read_source_text

pytestmark = pytest.mark.contract

ROOT = Path(__file__).resolve().parents[2]
CONTROLLER = ROOT / "packages/application/src/docwen_application/controller.py"
POLICY = ROOT / "packages/application/src/docwen_application/preconversion/intermediate_policy.py"
CONTROLLER_TESTS = ROOT / "packages/application/tests/test_controller_*.py"


def test_application_preconversion_admits_one_snapshot_before_external_work() -> None:
    controller = CONTROLLER.read_text(encoding="utf-8")

    assert 'config_snapshot = deepcopy(getattr(request, "config_snapshot", {}))' in controller
    assert "captured_snapshot = self._config_port.snapshot()" in controller
    assert "config_snapshot = deepcopy(captured_snapshot)" in controller
    assert "backend_priority = self._configured_priority(config_snapshot" in controller
    assert "policy_snapshot" not in controller
    assert "config_snapshot=config_snapshot" in controller
    assert "config_snapshot=deepcopy(config_snapshot)" in controller
    assert "self._config_port.get(key, list(default))" not in controller
    assert controller.index("captured_snapshot = self._config_port.snapshot()") < controller.index(
        "pre_result = pre_convert("
    )


def test_intermediate_save_policy_uses_only_the_request_snapshot() -> None:
    policy = POLICY.read_text(encoding="utf-8")

    assert "def should_save_intermediates(*, config_snapshot:" in policy
    assert 'output = config_snapshot.get("output", {})' in policy
    assert 'intermediate_files.get("save_to_output", False)' in policy
    assert "should_save_intermediates(config_snapshot=config_snapshot)" in policy
    assert "should_save_intermediate_files" not in policy


def test_controller_regressions_cover_conflicts_capture_and_untrusted_state() -> None:
    tests = read_source_text(CONTROLLER_TESTS)

    for token in (
        "test_preconversion_intermediate_save_prefers_request_snapshot",
        "test_preconversion_intermediate_save_captures_config_port_snapshot",
        "test_preconversion_empty_config_port_snapshot_is_authoritative",
        "test_preconversion_rejects_untrusted_config_before_backend_execution",
        "mock_config.snapshot.assert_called_once_with()",
        "mock_config.get.assert_not_called()",
        "assert request.config_snapshot == {}",
        'pytest.raises(RuntimeError, match="configuration state is untrusted")',
    ):
        assert token in tests

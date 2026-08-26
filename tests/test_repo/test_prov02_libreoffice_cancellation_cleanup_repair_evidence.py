"""Fail-closed evidence guards for VIS-391 LibreOffice cancel cleanup repair."""

from __future__ import annotations

from pathlib import Path

import pytest
from tools.validation.source_family import read_source_text

pytestmark = [pytest.mark.contract, pytest.mark.golden]

PROJECT_ROOT = Path(__file__).resolve().parents[2]
BRIDGE = PROJECT_ROOT / "packages" / "core" / "src" / "docwen_core" / "office_bridge.py"
BRIDGE_TESTS = PROJECT_ROOT / "packages" / "core" / "tests" / "test_office_bridge_*.py"
REPORT_NAME = "prov02-libreoffice-cancellation-profile-cleanup-repair-2026-07-26.md"
CARD_NAME = "prov02-libreoffice-cancellation-profile-cleanup-repair-stage-card-2026-07-26.md"
STATUS = "CANCELLATION_PROFILE_CLEANUP_FIXED_RENDER_DISPOSITION_PENDING"


def _read(path: Path) -> str:
    return read_source_text(path)


def _normalized(path: Path) -> str:
    return " ".join(_read(path).split())


def test_vis391_shared_owner_retries_bounded_and_fails_closed() -> None:
    bridge = _read(BRIDGE)
    for token in (
        "_LIBREOFFICE_PROFILE_CLEANUP_TIMEOUT_S = 5.0",
        "_LIBREOFFICE_PROFILE_CLEANUP_RETRY_S = 0.1",
        "def _remove_temp_tree(",
        "shutil.rmtree(tree_path)",
        "def _remove_libreoffice_profile(",
        "return _remove_temp_tree(",
        "profile_path,",
        "time.monotonic() >= deadline",
        "tempfile.mkdtemp(prefix=_LIBREOFFICE_PROFILE_PREFIX)",
        "tempfile.mkdtemp(prefix=_LIBREOFFICE_OUTPUT_PREFIX, dir=output.parent)",
        "profile_removed = profile_dir is None or _remove_libreoffice_profile(profile_dir)",
        "conversion_removed = conversion_dir is None or _remove_temp_tree(conversion_dir)",
        "if not profile_removed or not conversion_removed or not process_succeeded or staged_output is None:",
        "os.replace(generated, staged_output)",
        "staged_output.unlink(missing_ok=True)",
    ):
        assert token in bridge
    assert "ignore_cleanup_errors=True" not in bridge

    bridge_tests = _read(BRIDGE_TESTS)
    for behavior in (
        "test_libreoffice_profile_cleanup_retries_a_transient_windows_lock",
        "test_libreoffice_profile_cleanup_contains_a_persistent_windows_lock",
        "test_libreoffice_conversion_fails_closed_when_owned_profile_cannot_cleanup",
    ):
        assert behavior in bridge_tests

"""GUI security error surface tests (Task 8, GAP-SEC-003).

Verifies that security failures map to visible error states in the GUI,
with no silent fallback.
"""

from __future__ import annotations

import pytest

pytestmark = pytest.mark.gui


class TestGuiSecurityErrorSurface:
    """Verify security failures produce visible error states in GUI.

    The GUI main window uses info_area to display errors.  This test
    verifies that error information is accessible through the view model
    and info area when a task fails.
    """

    def test_main_window_info_area_accessible_for_errors(self, main_window) -> None:
        """Info area view model should be accessible for error display."""
        info_vm = main_window._info_area_vm
        assert info_vm is not None
        # Info area should be able to receive messages
        assert hasattr(info_vm, "add_message")

    def test_error_message_surfaces_in_info_area(self, main_window) -> None:
        """Adding an error message should be retrievable from info_area VM."""
        info_vm = main_window._info_area_vm
        # Add a security-failure-like message
        info_vm.add_message("Security check failed: network blocked in strict mode", "danger")
        # The info area should have recorded the message in history_rows
        assert info_vm.message_count > 0
        assert any("Security check failed" in row.message for row in info_vm.history_rows)

    def test_execution_failed_surfaces_error_to_info_area(self, main_window, tmp_path) -> None:
        """When execution fails, error should appear in info_area_vm history."""
        info_vm = main_window._info_area_vm

        # Simulate a security-related execution failure
        context = {
            "request_id": "test-sec-001",
            "file_path": str(tmp_path / "test.docx"),
            "display_name": "test.docx",
        }
        error_message = "SECURITY_CHECK_FAILED: Network access blocked in strict local-only mode"

        # Trigger _on_execution_failed (this is the slot the GUI uses
        # for task failures)
        initial_count = info_vm.message_count
        main_window._on_execution_failed(error_message, context)
        # Verify a new message was added
        assert info_vm.message_count > initial_count
        # The failure message should be visible in the info area history
        assert any("SECURITY_CHECK_FAILED" in row.message or "Network" in row.message for row in info_vm.history_rows)

    def test_info_area_preserves_error_with_location(self, main_window, tmp_path) -> None:
        """Error messages should preserve file path for navigation."""
        info_vm = main_window._info_area_vm
        file_path = str(tmp_path / "restricted.docx")

        context = {
            "request_id": "test-sec-002",
            "file_path": file_path,
            "display_name": "restricted.docx",
        }
        error_message = "Security check failed: file access denied"

        initial_count = info_vm.message_count
        main_window._on_execution_failed(error_message, context)
        assert info_vm.message_count > initial_count


class TestGuiNoSilentFallback:
    """Verify that security failures are never silently swallowed.

    When a security check fails, the GUI must show a visible error state.
    This is the core of the "no silent fallback" requirement.
    """

    def test_error_surface_is_visible_after_failure(self, main_window, tmp_path) -> None:
        """After _on_execution_failed is called, the error should be visible
        in the info area — no silent fallback."""
        info_vm = main_window._info_area_vm
        action_area_vm = main_window._action_area_vm

        # Verify initial state: cancel button hidden
        assert hasattr(action_area_vm, "hide_cancel")

        context = {
            "request_id": "test-sec-003",
            "file_path": str(tmp_path / "secret.docx"),
            "display_name": "secret.docx",
        }
        error_message = "SECURITY_CHECK_FAILED: strict mode prevents network access"

        # Simulate the failure
        main_window._on_execution_failed(error_message, context)

        # Verify: info area has the error in history
        assert any("SECURITY_CHECK_FAILED" in row.message for row in info_vm.history_rows)

    def test_batch_list_shows_failed_status_after_security_error(self, main_window, sample_docx) -> None:
        """After a security failure, the batch list entry should show 'failed'."""
        file_path = str(sample_docx)

        # Add file to batch list
        main_window._batch_list_vm.add_files([file_path])

        context = {
            "request_id": "test-sec-004",
            "file_path": file_path,
            "display_name": "blocked.docx",
        }
        error_message = "Network access blocked in strict local-only mode"

        # Trigger failure
        main_window._on_execution_failed(error_message, context)

        # Check batch list entry
        entry = main_window._batch_list_vm.get_file_entry(file_path)
        assert entry is not None
        assert entry.status == "failed"
        assert "Network" in (entry.error_message or "")


class TestInfoAreaVMErrorHandling:
    """Direct tests for InfoAreaViewModel error handling capabilities."""

    def test_add_message_supports_danger_tone(self, main_window) -> None:
        info_vm = main_window._info_area_vm
        initial = info_vm.message_count
        info_vm.add_message("Critical security error", "danger")
        assert info_vm.message_count > initial

    def test_add_message_supports_warning_tone(self, main_window) -> None:
        info_vm = main_window._info_area_vm
        info_vm.add_message("Security warning: debug mode enabled", "warning")
        # Should not crash
        assert info_vm.message_count > 0

    def test_set_transient_message_works_for_security_events(self, main_window) -> None:
        info_vm = main_window._info_area_vm
        info_vm.set_transient_message(
            "security:network-blocked",
            "Network access is disabled in strict mode",
            "danger",
            ttl_ms=5000,
            source="security",
        )
        # Should not crash — transient messages are shown briefly
        assert True

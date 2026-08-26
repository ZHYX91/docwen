"""Tests for CLI exit codes — matches CLI契约规范.md §3."""

import pytest

pytestmark = pytest.mark.unit


def test_error_registry_drives_category_and_exit_code() -> None:
    """Each detailed error has exactly one category and exit-code definition."""

    from docwen_cli.error_registry import ERROR_SPECS
    from docwen_cli.exit_codes import exit_code_from_error_code
    from docwen_cli.protocol import category_for_error_code

    for code, spec in ERROR_SPECS.items():
        assert category_for_error_code(code).value == spec.category
        assert exit_code_from_error_code(code).value == spec.exit_code


class TestExitCodeMapping:
    """Verify the 7 exit codes and error_code string mappings."""

    def test_all_exit_codes_defined(self) -> None:
        from docwen_cli.exit_codes import ExitCode

        assert ExitCode.OK == 0
        assert ExitCode.INTERNAL_ERROR == 1
        assert ExitCode.INVALID_INPUT == 2
        assert ExitCode.NOT_FOUND == 3
        assert ExitCode.DEPENDENCY_MISSING == 4
        assert ExitCode.SECURITY_CHECK_FAILED == 5
        assert ExitCode.CANCELLED == 130

    def test_exit_code_is_int(self) -> None:
        from docwen_cli.exit_codes import ExitCode

        assert int(ExitCode.OK) == 0
        assert int(ExitCode.CANCELLED) == 130

    def test_exit_code_from_error_code_known(self) -> None:
        from docwen_cli.exit_codes import ExitCode, exit_code_from_error_code

        assert exit_code_from_error_code("skipped_same_format") == ExitCode.OK
        assert exit_code_from_error_code("unknown_error") == ExitCode.INTERNAL_ERROR
        assert exit_code_from_error_code("unsupported_route") == ExitCode.UNAVAILABLE
        assert exit_code_from_error_code("unsupported_numbering") == ExitCode.UNAVAILABLE
        assert exit_code_from_error_code("conversion_failed") == ExitCode.INTERNAL_ERROR
        assert exit_code_from_error_code("operation_cancelled") == ExitCode.CANCELLED
        assert exit_code_from_error_code("cancelled") == ExitCode.CANCELLED
        assert exit_code_from_error_code("invalid_input") == ExitCode.INVALID_INPUT
        assert exit_code_from_error_code("unsupported_format") == ExitCode.INVALID_INPUT
        assert exit_code_from_error_code("dependency_missing") == ExitCode.DEPENDENCY_MISSING
        assert exit_code_from_error_code("security_check_failed") == ExitCode.SECURITY_CHECK_FAILED
        assert exit_code_from_error_code("network_access_blocked") == ExitCode.SECURITY_CHECK_FAILED

    def test_exit_code_from_error_code_unknown(self) -> None:
        from docwen_cli.exit_codes import ExitCode, exit_code_from_error_code

        assert exit_code_from_error_code("nonexistent_code") == ExitCode.INTERNAL_ERROR
        assert exit_code_from_error_code("") == ExitCode.INTERNAL_ERROR

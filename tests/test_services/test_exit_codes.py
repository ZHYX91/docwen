from docwen.errors import ExitCode, exit_code_from_error_code
from docwen.services.error_codes import ERROR_CODE_SECURITY_CHECK_FAILED


def test_security_check_failed_has_dedicated_exit_code() -> None:
    assert exit_code_from_error_code(ERROR_CODE_SECURITY_CHECK_FAILED) == ExitCode.SECURITY_CHECK_FAILED


"""Contract tests for bounded public Windows path handling."""

from __future__ import annotations

import argparse
import json

import pytest

pytestmark = pytest.mark.contract


@pytest.mark.parametrize(
    "namespace_path",
    [
        r"\\?\C:\very-long\document.md",
        r"\\.\PhysicalDrive0",
        "//?/C:/very-long/document.md",
        "//./PhysicalDrive0",
    ],
)
def test_windows_namespaces_are_rejected_after_normalization(namespace_path: str) -> None:
    from docwen_cli.path_policy import check_public_path

    issue = check_public_path(namespace_path, platform_name="nt")

    assert issue is not None
    assert "extended-length" in issue.message


def test_windows_exact_utf16_path_limit_is_rejected_before_backends() -> None:
    from docwen_cli.path_policy import WINDOWS_PUBLIC_PATH_LIMIT, check_public_path

    at_limit = "C:\\" + "a" * (WINDOWS_PUBLIC_PATH_LIMIT - 3)
    over_limit = f"{at_limit}a"

    assert check_public_path(at_limit, platform_name="nt") is None
    issue = check_public_path(over_limit, platform_name="nt")
    assert issue is not None
    assert str(WINDOWS_PUBLIC_PATH_LIMIT) in issue.message
    assert check_public_path("/tmp/" + "a" * 400, platform_name="posix") is None


def test_windows_public_path_limit_counts_non_bmp_utf16_units() -> None:
    from docwen_cli.path_policy import check_public_path

    at_limit = "C:\\" + "\U0001f4c4" * 128
    over_limit = f"{at_limit}x"

    assert len(over_limit) < 260
    assert check_public_path(at_limit, platform_name="nt") is None
    issue = check_public_path(over_limit, platform_name="nt")
    assert issue is not None
    assert "UTF-16 code units" in issue.message


def test_canonical_template_resource_id_bypasses_path_policy_but_legacy_path_does_not(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_cli import path_policy

    native_check = path_policy.check_public_path
    monkeypatch.setattr(
        path_policy,
        "check_public_path",
        lambda raw_path: native_check(raw_path, platform_name="nt"),
    )
    monkeypatch.setattr(path_policy.ntpath, "abspath", lambda value: "C:\\" + "x" * 300 + value)
    canonical_id = f"template.docx.{'a' * 64}"

    assert path_policy.first_namespace_path_issue(argparse.Namespace(template=canonical_id)) is None
    issue = path_policy.first_namespace_path_issue(argparse.Namespace(template="legacy-template.docx"))
    assert issue is not None
    assert issue.path == "legacy-template.docx"


def test_nested_path_error_keeps_leaf_command_in_json_envelope(
    capsys: pytest.CaptureFixture[str],
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_cli import path_policy
    from docwen_cli.exit_codes import ExitCode
    from docwen_cli.main import main

    native_check = path_policy.check_public_path
    monkeypatch.setattr(
        path_policy,
        "check_public_path",
        lambda raw_path: native_check(raw_path, platform_name="nt"),
    )

    exit_code = main(
        [
            "number",
            "markdown",
            r"\\?\C:\source.md",
            "--operation",
            "add",
            "--output",
            r"\\?\C:\output.md",
            "--json",
            "--quiet",
        ]
    )

    payload = json.loads(capsys.readouterr().out)
    assert exit_code == int(ExitCode.INVALID_INPUT)
    assert payload["command"] == "number markdown"
    assert payload["error"]["category"] == "invalid_input"
    assert payload["error"]["code"] == "invalid_path"
    assert payload["error"]["details"]["path"].startswith("\\\\?\\")

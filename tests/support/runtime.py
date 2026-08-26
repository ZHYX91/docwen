from __future__ import annotations

import os
from collections.abc import Iterator
from pathlib import Path
from typing import Literal

import pytest


@pytest.fixture(scope="session", autouse=True)
def isolate_user_runtime_roots(tmp_path_factory: pytest.TempPathFactory) -> Iterator[None]:
    """Keep the test session away from the real user's DocWen data.

    Tests that need a specific configuration root override this environment
    variable explicitly.  Default user-mode logging is redirected without
    defining ``DOCWEN_LOG_DIR`` so tests still exercise the real precedence
    rules for explicit environment and custom-directory overrides.

    Both session defaults must be disposable: a test run must neither acquire
    the desktop application's lock nor mutate a developer's settings or logs
    merely because their home directory is writable.
    """

    import platformdirs

    config_variable = "DOCWEN_CONFIG_DIR"
    previous_config_root = os.environ.get(config_variable)
    os.environ[config_variable] = str(tmp_path_factory.mktemp("docwen-user-config"))

    isolated_log_root = tmp_path_factory.mktemp("docwen-user-log")
    original_user_log_dir = platformdirs.user_log_dir

    def isolated_user_log_dir(
        appname: str | None = None,
        appauthor: str | Literal[False] | None = None,
        version: str | None = None,
        opinion: bool = True,
        ensure_exists: bool = False,
        use_site_for_root: bool = False,
    ) -> str:
        if appname == "docwen" and appauthor is False:
            path = isolated_log_root / version if version else isolated_log_root
            if ensure_exists:
                Path(path).mkdir(parents=True, exist_ok=True)
            return str(path)
        return original_user_log_dir(
            appname=appname,
            appauthor=appauthor,
            version=version,
            opinion=opinion,
            ensure_exists=ensure_exists,
            use_site_for_root=use_site_for_root,
        )

    platformdirs.user_log_dir = isolated_user_log_dir
    try:
        yield
    finally:
        platformdirs.user_log_dir = original_user_log_dir
        if previous_config_root is None:
            os.environ.pop(config_variable, None)
        else:
            os.environ[config_variable] = previous_config_root

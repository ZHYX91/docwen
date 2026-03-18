"""converter 单元测试。"""

from __future__ import annotations

import pytest

from docwen.converter.formats.common import fallback as common_fallback

pytestmark = pytest.mark.unit


def test_find_soffice_path_returns_which(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(common_fallback.shutil, "which", lambda _: "C:/bin/soffice")
    monkeypatch.setattr(common_fallback.sys, "platform", "linux")
    assert common_fallback.find_soffice_path() == "C:/bin/soffice"


def test_find_soffice_path_darwin_checks_default_paths(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(common_fallback.shutil, "which", lambda _: None)
    monkeypatch.setattr(common_fallback.sys, "platform", "darwin")
    monkeypatch.setattr(
        common_fallback.os.path,
        "expanduser",
        lambda _: "/Users/u/Applications/LibreOffice.app/Contents/MacOS/soffice",
    )
    monkeypatch.setattr(
        common_fallback.os.path,
        "isfile",
        lambda p: p == "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    )
    assert common_fallback.find_soffice_path() == "/Applications/LibreOffice.app/Contents/MacOS/soffice"


def test_find_soffice_path_falls_back_to_name(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(common_fallback.shutil, "which", lambda _: None)
    monkeypatch.setattr(common_fallback.sys, "platform", "linux")
    assert common_fallback.find_soffice_path() == "soffice"

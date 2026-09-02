"""Fail-closed evidence guards for VIS-389 Windows discovery repair."""

from __future__ import annotations

from pathlib import Path

import pytest

pytestmark = [pytest.mark.contract, pytest.mark.golden]

PROJECT_ROOT = Path(__file__).resolve().parents[2]
BRIDGE = PROJECT_ROOT / "packages" / "core" / "src" / "docwen_core" / "office_bridge.py"
REPORT_NAME = "prov02-windows-registered-discovery-repair-2026-07-26.md"
CARD_NAME = "prov02-windows-registered-discovery-repair-stage-card-2026-07-26.md"
STATUS = "REGISTERED_INSTALL_DISCOVERY_FIXED_ROUTE_ACCEPTANCE_PENDING"


def _read(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def _normalized(path: Path) -> str:
    return " ".join(_read(path).split())


def test_vis389_shared_owner_is_bounded_read_only_and_fail_closed() -> None:
    bridge = _read(BRIDGE)
    discovery = bridge[bridge.index("def _existing_file_path") : bridge.index("def _import_win32")]
    for token in (
        'shutil.which("soffice")',
        "_WINDOWS_SOFFICE_APP_PATH",
        "winreg.HKEY_CURRENT_USER",
        "winreg.HKEY_LOCAL_MACHINE",
        "winreg.QueryValueEx",
        '"ProgramW6432"',
        '"ProgramFiles"',
        '"ProgramFiles(x86)"',
        "path.is_file()",
    ):
        assert token in discovery
    for forbidden in ("SetValue", "setx", "os.environ[", "rglob(", "walk("):
        assert forbidden not in discovery

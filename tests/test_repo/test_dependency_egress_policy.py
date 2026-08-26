"""Static ownership gate for production network-client imports."""

from __future__ import annotations

import ast
from pathlib import Path

import pytest

pytestmark = pytest.mark.unit

ROOT = Path(__file__).resolve().parents[2]
PACKAGES = ROOT / "packages"
SOCKET_OWNER = Path("runtime/src/docwen_runtime/security/network.py")
FORBIDDEN_MODULES = {
    "aiohttp",
    "dns",
    "ftplib",
    "grpc",
    "http.client",
    "httpx",
    "requests",
    "smtplib",
    "telnetlib",
    "urllib3",
    "websocket",
    "websockets",
}
FORBIDDEN_QT_MODULES = {"QtNetwork", "QtNetworkAuth", "QtWebSockets"}


def _is_forbidden_module(module: str) -> bool:
    return any(module == forbidden or module.startswith(f"{forbidden}.") for forbidden in FORBIDDEN_MODULES)


def _production_sources() -> list[Path]:
    return sorted(path for path in PACKAGES.rglob("*.py") if "tests" not in path.parts)


def test_production_packages_do_not_import_network_clients() -> None:
    violations: list[str] = []
    for path in _production_sources():
        tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
        package_relative = path.relative_to(PACKAGES)
        for node in ast.walk(tree):
            if isinstance(node, ast.Import):
                modules = {alias.name for alias in node.names}
            elif isinstance(node, ast.ImportFrom):
                modules = {node.module or ""}
                if node.module == "PySide6":
                    modules.update(
                        f"PySide6.{alias.name}" for alias in node.names if alias.name in FORBIDDEN_QT_MODULES
                    )
            else:
                continue
            if any(_is_forbidden_module(module) for module in modules):
                violations.append(f"{package_relative}:{node.lineno}: network client import")
            if any(
                module == f"PySide6.{qt_module}" or module.startswith(f"PySide6.{qt_module}.")
                for module in modules
                for qt_module in FORBIDDEN_QT_MODULES
            ):
                violations.append(f"{package_relative}:{node.lineno}: forbidden Qt network import")
            if "socket" in modules and package_relative.as_posix() != SOCKET_OWNER.as_posix():
                violations.append(f"{package_relative}:{node.lineno}: socket import outside security owner")
    assert not violations, "\n".join(violations)


def test_urllib_request_is_limited_to_local_path_decoding() -> None:
    violations: list[str] = []
    for path in _production_sources():
        tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
        for node in ast.walk(tree):
            if isinstance(node, ast.Import) and any(alias.name == "urllib.request" for alias in node.names):
                violations.append(f"{path.relative_to(ROOT)}:{node.lineno}: module import")
            if isinstance(node, ast.ImportFrom) and node.module == "urllib.request":
                imported = {alias.name for alias in node.names}
                if imported != {"url2pathname"}:
                    violations.append(f"{path.relative_to(ROOT)}:{node.lineno}: {sorted(imported)}")
    assert not violations, "\n".join(violations)

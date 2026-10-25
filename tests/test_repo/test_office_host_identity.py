"""Installed-host identity reads must not launch Windows LibreOffice."""

from __future__ import annotations

import hashlib
import subprocess
import sys
from pathlib import Path
from types import SimpleNamespace

import pytest
from scripts.release import office_host_identity as identity

pytestmark = pytest.mark.contract


def _selected_executable(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> Path:
    from docwen_core import office_bridge

    executable = tmp_path / "selected-soffice.exe"
    executable.write_bytes(b"selected executable identity fixture")
    monkeypatch.setattr(office_bridge, "find_soffice_path", lambda: str(executable))
    monkeypatch.setattr(identity, "os", SimpleNamespace(name="nt"))

    def forbidden_launch(*_args: object, **_kwargs: object) -> None:
        pytest.fail("Windows host identity must not launch a process")

    monkeypatch.setattr(identity.subprocess, "run", forbidden_launch)
    monkeypatch.setattr(identity.subprocess, "Popen", forbidden_launch)
    return executable


def test_windows_libreoffice_reads_selected_pe_version_without_launch(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    executable = _selected_executable(tmp_path, monkeypatch)
    calls: list[tuple[str, str]] = []

    def version_info(path: str, resource: str) -> dict[str, int]:
        calls.append((path, resource))
        return {"FileVersionMS": (26 << 16) | 2, "FileVersionLS": (6 << 16) | 3}

    monkeypatch.setitem(sys.modules, "win32api", SimpleNamespace(GetFileVersionInfo=version_info))
    result = identity.office_host_identity("LibreOffice")
    assert calls == [(str(executable.resolve()), "\\")]
    assert result == {
        "backend": "LibreOffice",
        "version": "26.2.6.3",
        "versionSource": "selected LibreOffice executable FileVersion",
        "executable": str(executable.resolve()),
        "bytes": executable.stat().st_size,
        "sha256": hashlib.sha256(executable.read_bytes()).hexdigest(),
    }


@pytest.mark.parametrize(
    "metadata",
    [
        {},
        {"FileVersionMS": 0, "FileVersionLS": 0},
        {"FileVersionMS": "26", "FileVersionLS": 1},
        {"FileVersionMS": -1, "FileVersionLS": 1},
        {"FileVersionMS": 1, "FileVersionLS": 0x100000000},
    ],
)
def test_windows_libreoffice_rejects_missing_or_invalid_version(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, metadata: dict[str, object]
) -> None:
    _selected_executable(tmp_path, monkeypatch)
    monkeypatch.setitem(sys.modules, "win32api", SimpleNamespace(GetFileVersionInfo=lambda *_args: metadata))
    with pytest.raises(RuntimeError, match="office_version_empty_or_invalid"):
        identity.office_host_identity("libreoffice")


def test_windows_libreoffice_preserves_version_read_error(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    _selected_executable(tmp_path, monkeypatch)

    def unavailable(*_args: object) -> None:
        raise OSError("PE version resource missing")

    monkeypatch.setitem(sys.modules, "win32api", SimpleNamespace(GetFileVersionInfo=unavailable))
    with pytest.raises(OSError, match="PE version resource missing"):
        identity.office_host_identity("libreoffice")


@pytest.mark.parametrize("backend_id", ["msoffice_powerpoint", "wps_presentation"])
def test_presentation_producer_ids_and_labels_resolve_registered_host_identity(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, backend_id: str
) -> None:
    from docwen_plugin_presentation.pptx_md.ppt_converter import _PRESENTATION_CANDIDATES

    candidate = _PRESENTATION_CANDIDATES[backend_id]
    assert candidate.app_type == "powerpoint" and candidate.save_format == 24
    executable = _selected_executable(tmp_path, monkeypatch)
    registered: list[str] = []

    def resolve(prog_id: str) -> Path:
        registered.append(prog_id)
        return executable

    monkeypatch.setattr(identity, "_com_executable", resolve)
    monkeypatch.setitem(
        sys.modules,
        "win32api",
        SimpleNamespace(GetFileVersionInfo=lambda *_args: {"FileVersionMS": 16 << 16, "FileVersionLS": 42}),
    )
    for backend in (backend_id, candidate.name):
        result = identity.office_host_identity(backend)
        assert result["backend"] == backend
        assert result["version"] == "16.0.0.42"
        assert result["versionSource"] == "registered COM server executable FileVersion"
        assert result["executable"] == str(executable)
        assert result["sha256"] == hashlib.sha256(executable.read_bytes()).hexdigest()
    assert registered == [candidate.prog_id, candidate.prog_id]


@pytest.mark.parametrize("stdout", ["LibreOffice 26.2.6.3\n", "\n"])
def test_nonwindows_libreoffice_keeps_command_version_probe(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, stdout: str
) -> None:
    executable = _selected_executable(tmp_path, monkeypatch)
    monkeypatch.setattr(identity, "os", SimpleNamespace(name="posix"))
    calls: list[list[str]] = []

    def probe(command: list[str], **_kwargs: object) -> subprocess.CompletedProcess[str]:
        calls.append(command)
        return subprocess.CompletedProcess(command, 0, stdout=stdout)

    monkeypatch.setattr(identity.subprocess, "run", probe)
    if stdout.strip():
        result = identity.office_host_identity("libreoffice")
        assert result["version"] == stdout.strip()
        assert result["versionSource"] == "selected LibreOffice executable --version"
    else:
        with pytest.raises(RuntimeError, match="office_version_empty"):
            identity.office_host_identity("libreoffice")
    assert calls == [[str(executable.resolve()), "--version"]]

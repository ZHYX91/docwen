from __future__ import annotations

import hashlib
import stat
import subprocess
import sys
import zipfile
from pathlib import Path

import pytest
from scripts.build import payload_normalization

pytestmark = pytest.mark.contract

REPO_ROOT = Path(__file__).resolve().parents[2]


def _write_base_library(path: Path, names: list[str], *, year: int) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name in names:
            info = zipfile.ZipInfo(name, date_time=(year, 1, 2, 3, 4, 6))
            info.create_system = 3
            info.external_attr = (stat.S_IFREG | 0o600) << 16
            archive.writestr(info, f"data:{name}".encode(), compress_type=zipfile.ZIP_DEFLATED)


def test_base_library_normalization_removes_order_and_metadata_variance(tmp_path: Path) -> None:
    first = tmp_path / "first"
    second = tmp_path / "second"
    names = ["encodings/utf_8.pyc", "abc.pyc", "re/__init__.pyc"]
    _write_base_library(first / "_internal" / "base_library.zip", names, year=2025)
    _write_base_library(second / "_internal" / "base_library.zip", list(reversed(names)), year=2026)

    assert payload_normalization.normalize_base_library_zip(first) == {"changedFiles": 1, "entries": 3}
    assert payload_normalization.normalize_base_library_zip(second) == {"changedFiles": 1, "entries": 3}

    first_bytes = (first / "_internal" / "base_library.zip").read_bytes()
    second_bytes = (second / "_internal" / "base_library.zip").read_bytes()
    assert hashlib.sha256(first_bytes).digest() == hashlib.sha256(second_bytes).digest()
    assert payload_normalization.normalize_base_library_zip(first) == {"changedFiles": 0, "entries": 3}
    with zipfile.ZipFile(first / "_internal" / "base_library.zip") as archive:
        assert archive.namelist() == sorted(names, key=lambda value: value.encode("utf-8"))
        assert archive.comment == b""
        assert all(info.date_time == (1980, 1, 1, 0, 0, 0) for info in archive.infolist())
        assert all(info.extra == b"" for info in archive.infolist())


def test_record_normalization_removes_only_external_rows(tmp_path: Path) -> None:
    payload = tmp_path / "payload"
    record = payload / "_internal" / "example-1.0.dist-info" / "RECORD"
    record.parent.mkdir(parents=True)
    record.write_text(
        "package.py,sha256=stable,1\n../../../bin/example,sha256=volatile,2\nexample-1.0.dist-info/RECORD,,\n",
        encoding="utf-8",
    )

    assert payload_normalization.normalize_packaged_record_files(payload) == {
        "changedFiles": 1,
        "removedRows": 1,
    }
    assert record.read_text(encoding="utf-8") == ("package.py,sha256=stable,1\nexample-1.0.dist-info/RECORD,,\n")


def test_production_builder_direct_entrypoint_resolves_shared_normalizer() -> None:
    result = subprocess.run(
        [sys.executable, "scripts/release/build_production_candidate.py", "--help"],
        cwd=REPO_ROOT,
        capture_output=True,
        text=True,
        check=False,
    )

    assert result.returncode == 0, result.stderr
    assert "--calibrate-allowlist" in result.stdout

"""Canonicalize PyInstaller metadata that otherwise depends on build-local state."""

from __future__ import annotations

import csv
import io
import os
import stat
import tempfile
import zipfile
from pathlib import Path, PurePosixPath


class PayloadNormalizationError(RuntimeError):
    """Raised when a PyInstaller payload cannot be normalized safely."""


def _require(condition: bool, message: str) -> None:
    if not condition:
        raise PayloadNormalizationError(message)


def normalize_packaged_record_files(payload: Path) -> dict[str, int]:
    """Remove wheel ``RECORD`` rows for launchers absent from the payload.

    Console launchers live outside ``site-packages`` in the build venv. Their
    shebangs contain the venv path, so their hashes vary even though PyInstaller
    does not ship the launchers. Keeping those rows would bind the frozen
    payload to files that are both absent and build-location-dependent.
    """

    internal = payload / "_internal"
    changed_files = 0
    removed_rows = 0
    for record in sorted(internal.glob("*.dist-info/RECORD"), key=lambda item: item.name.encode("utf-8")):
        source = record.read_text(encoding="utf-8")
        reader = csv.reader(io.StringIO(source))
        kept: list[list[str]] = []
        removed = 0
        for row in reader:
            _require(bool(row), f"packaged_record_empty_row:{record.name}")
            path = PurePosixPath(row[0].replace("\\", "/"))
            if path.is_absolute() or ".." in path.parts:
                removed += 1
                continue
            kept.append(row)
        if removed:
            buffer = io.StringIO(newline="")
            writer = csv.writer(buffer, lineterminator="\n")
            writer.writerows(kept)
            record.write_text(buffer.getvalue(), encoding="utf-8", newline="\n")
            changed_files += 1
            removed_rows += removed
    return {"changedFiles": changed_files, "removedRows": removed_rows}


def _zip_path_key(value: str) -> bytes:
    return value.encode("utf-8")


def normalize_base_library_zip(payload: Path) -> dict[str, int]:
    """Repack PyInstaller's standard-library ZIP with canonical metadata/order."""

    archive_path = payload / "_internal" / "base_library.zip"
    _require(archive_path.is_file() and not archive_path.is_symlink(), "base_library_zip_missing")
    original_mode = stat.S_IMODE(archive_path.stat().st_mode)
    original_bytes = archive_path.read_bytes()
    entries: list[tuple[str, bool, bytes]] = []
    seen: set[str] = set()
    folded: dict[str, str] = {}
    try:
        with zipfile.ZipFile(io.BytesIO(original_bytes), "r") as source:
            _require(source.testzip() is None, "base_library_zip_crc_invalid")
            for info in source.infolist():
                name = info.filename
                normalized_name = name[:-1] if info.is_dir() and name.endswith("/") else name
                path = PurePosixPath(normalized_name)
                _require(
                    bool(normalized_name) and "\\" not in name and not path.is_absolute() and ".." not in path.parts,
                    f"base_library_zip_path_invalid:{name}",
                )
                _require(name not in seen, f"base_library_zip_duplicate:{name}")
                seen.add(name)
                folded_name = name.casefold()
                _require(
                    folded_name not in folded or folded[folded_name] == name,
                    f"base_library_zip_casefold_collision:{folded.get(folded_name)}:{name}",
                )
                folded[folded_name] = name
                entry_type = stat.S_IFMT(info.external_attr >> 16)
                _require(entry_type != stat.S_IFLNK, f"base_library_zip_symlink_forbidden:{name}")
                entries.append((name, info.is_dir(), b"" if info.is_dir() else source.read(info)))
    except (OSError, zipfile.BadZipFile, RuntimeError) as exc:
        if isinstance(exc, PayloadNormalizationError):
            raise
        raise PayloadNormalizationError("base_library_zip_unreadable") from exc

    entries.sort(key=lambda item: _zip_path_key(item[0]))
    temporary_path: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="wb",
            prefix=f".{archive_path.name}.",
            suffix=".tmp",
            dir=archive_path.parent,
            delete=False,
        ) as temporary:
            temporary_path = Path(temporary.name)
        with zipfile.ZipFile(
            temporary_path,
            "w",
            compression=zipfile.ZIP_DEFLATED,
            compresslevel=9,
            strict_timestamps=True,
        ) as destination:
            destination.comment = b""
            for name, is_directory, data in entries:
                info = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
                info.create_system = 3
                info.compress_type = zipfile.ZIP_DEFLATED
                info.flag_bits |= 0x800
                mode = 0o755 if is_directory else 0o644
                file_type = stat.S_IFDIR if is_directory else stat.S_IFREG
                info.external_attr = (file_type | mode) << 16
                destination.writestr(info, data, compress_type=zipfile.ZIP_DEFLATED, compresslevel=9)
        normalized_bytes = temporary_path.read_bytes()
        changed = normalized_bytes != original_bytes
        if changed:
            os.replace(temporary_path, archive_path)
            temporary_path = None
            archive_path.chmod(original_mode)
        return {"changedFiles": int(changed), "entries": len(entries)}
    finally:
        if temporary_path is not None:
            temporary_path.unlink(missing_ok=True)


__all__ = [
    "PayloadNormalizationError",
    "normalize_base_library_zip",
    "normalize_packaged_record_files",
]

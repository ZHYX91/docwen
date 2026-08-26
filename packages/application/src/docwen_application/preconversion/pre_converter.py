"""Pre-convert document-family non-hub files to the hub format.

Uses the config-free ordered bridge executor from ``docwen_core.office_bridge``.
The controller supplies the configured backend order while this module owns
the legal candidates required by its document-to-DOCX hub step.
"""

from __future__ import annotations

import contextlib
import os
import stat
import sys
import tempfile
from collections.abc import Iterator, Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import TYPE_CHECKING, BinaryIO

from docwen_application.preconversion.chain_resolver import _CATEGORY_RULES
from docwen_core.office_bridge import BridgeCandidate, BridgeResult, convert_with_backend_priority

if TYPE_CHECKING:
    from docwen_core.protocols import CancellationTokenView

# Legal COM candidates for application-owned document-to-DOCX pre-conversion.
_DOCUMENT_CANDIDATES: dict[str, BridgeCandidate] = {
    "wps_writer": BridgeCandidate("WPS Writer", "Kwps.Application", 12, "word"),
    "msoffice_word": BridgeCandidate("Microsoft Word", "Word.Application", 12, "word"),
}
_DEFAULT_WORD_PRIORITY = ("wps_writer", "msoffice_word", "libreoffice")
_DEFAULT_ODT_PRIORITY = ("msoffice_word", "libreoffice")
_PROTECTIVE_SOURCE_SUFFIXES = {
    "doc": ".doc",
    "wps": ".doc",
    "rtf": ".rtf",
    "odt": ".odt",
}
_SNAPSHOT_CHUNK_SIZE = 1024 * 1024


class _SnapshotCancelled(Exception):
    """Preparation stopped before a snapshot was published."""


class _SnapshotSourceChanged(OSError):
    """The opened source changed while its snapshot was being prepared."""


class _SnapshotUnsafePath(OSError):
    """A source or staging path is a link/reparse/non-regular object."""


def backend_priority_spec(source_format: str) -> tuple[str, list[str]]:
    """Return the config key and defensive default for a hub conversion."""
    if source_format == "odt":
        return "software.special_conversions.odt", list(_DEFAULT_ODT_PRIORITY)
    return "software.default_priority.word_processors", list(_DEFAULT_WORD_PRIORITY)


@dataclass(frozen=True)
class PreConversionResult:
    """Metadata about a successful pre-conversion."""

    pre_converted_path: str
    """Path to the pre-converted hub-format file."""
    original_source_format: str
    """The original source format (e.g. ``"doc"``, ``"wps"``)."""
    backend: str
    """The backend that performed the conversion (e.g. ``"WPS Writer"``)."""


@dataclass(frozen=True)
class PreConversionFailure:
    """Structured bridge failure retained for the Application owner."""

    message: str
    """The bridge's original failure message."""

    cancelled: bool = False
    """Whether the bridge reported cooperative cancellation."""

    error_type: str = "dependency_missing"
    """Stable error category projected by the Application controller."""

    diagnostic_code: str = ""
    """Optional machine-readable diagnostic code."""

    cleanup_message: str = ""
    """Office workspace/publication cleanup status, when available."""

    cleanup_failed: bool = False
    """Whether the Office bridge could not fully clean private state."""


def _bridge_failure_error_type(result: BridgeResult) -> str:
    """Project one bridge failure without confusing execution with availability."""
    if result.cancelled:
        return "cancelled"
    if (
        result.error_code == "OFFICE_BACKEND_FAILED"
        and result.attempted_backend_ids
        and not result.available_backend_ids
    ):
        return "dependency_missing"
    return "conversion_failed"


def _detect_category(source_format: str) -> str | None:
    """Return the category name for *source_format*, or ``None``."""
    for category, (hub, sources) in _CATEGORY_RULES.items():
        if source_format in sources:
            return category
        if source_format == hub:
            return category
    return None


def _is_cancel_requested(cancel: CancellationTokenView | None) -> bool:
    """Read the Application-owned cancellation view."""
    return cancel is not None and cancel.is_cancelled


def _is_link_or_reparse(path: Path) -> bool:
    """Inspect *path* itself without following its terminal component."""
    try:
        path_stat = path.lstat()
    except FileNotFoundError:
        return False
    file_attributes = getattr(path_stat, "st_file_attributes", 0)
    reparse_flag = getattr(stat, "FILE_ATTRIBUTE_REPARSE_POINT", 0x400)
    return stat.S_ISLNK(path_stat.st_mode) or bool(file_attributes & reparse_flag)


def _is_name_surrogate(path: Path) -> bool:
    """Return whether *path* redirects name resolution to another target."""
    try:
        path_stat = path.lstat()
    except FileNotFoundError:
        return False
    if stat.S_ISLNK(path_stat.st_mode):
        return True
    file_attributes = getattr(path_stat, "st_file_attributes", 0)
    reparse_flag = getattr(stat, "FILE_ATTRIBUTE_REPARSE_POINT", 0x400)
    if not file_attributes & reparse_flag:
        return False
    reparse_tag = getattr(path_stat, "st_reparse_tag", 0)
    # Windows name-surrogate tags (junctions and symlinks) set bit 29. Cloud
    # placeholders are reparse points but not path redirections and remain
    # legal user sources once their data can be opened.
    return not reparse_tag or bool(reparse_tag & 0x20000000)


def _validate_snapshot_paths(source_path: Path, stage_path: Path, protected_input: Path) -> None:
    """Reject terminal link/reparse objects before any staging write."""
    if _is_name_surrogate(source_path):
        raise _SnapshotUnsafePath(f"Source is a symlink or name-surrogate reparse point: {source_path}")
    if _is_link_or_reparse(stage_path):
        raise _SnapshotUnsafePath(f"Staging directory is a symlink or reparse point: {stage_path}")
    try:
        stage_stat = stage_path.lstat()
    except FileNotFoundError as exc:
        raise _SnapshotUnsafePath(f"Staging directory does not exist: {stage_path}") from exc
    if not stat.S_ISDIR(stage_stat.st_mode):
        raise _SnapshotUnsafePath(f"Staging path is not a directory: {stage_path}")

    try:
        protected_stat = protected_input.lstat()
    except FileNotFoundError:
        return
    if _is_link_or_reparse(protected_input) or not stat.S_ISREG(protected_stat.st_mode):
        raise _SnapshotUnsafePath(f"Protected input destination is unsafe: {protected_input}")


def _snapshot_identity(source_stat: os.stat_result) -> tuple[int, int, int, int, int]:
    """Return mutation-sensitive identity for one already-open source handle."""
    return (
        source_stat.st_dev,
        source_stat.st_ino,
        source_stat.st_size,
        source_stat.st_mtime_ns,
        source_stat.st_ctime_ns,
    )


def _copy_snapshot_stream(
    source_stream: BinaryIO,
    destination_stream: BinaryIO,
    cancel: CancellationTokenView | None,
) -> None:
    """Copy one source handle with cancellation checks between bounded chunks."""
    while True:
        if _is_cancel_requested(cancel):
            raise _SnapshotCancelled
        chunk = source_stream.read(_SNAPSHOT_CHUNK_SIZE)
        if not chunk:
            break
        destination_stream.write(chunk)
        if _is_cancel_requested(cancel):
            raise _SnapshotCancelled


@contextlib.contextmanager
def _open_source_snapshot(source_path: Path) -> Iterator[BinaryIO]:
    """Open a no-follow source handle and hold the strongest host read lock."""
    if os.name == "nt":
        with _open_windows_source_snapshot(source_path) as source_stream:
            yield source_stream
        return

    flags = os.O_RDONLY | getattr(os, "O_CLOEXEC", 0) | getattr(os, "O_NOFOLLOW", 0)
    descriptor = os.open(source_path, flags)
    try:
        source_stream = os.fdopen(descriptor, "rb", buffering=0)
    except Exception:
        os.close(descriptor)
        raise
    try:
        if sys.platform != "win32":
            import fcntl

            fcntl.flock(descriptor, fcntl.LOCK_SH | fcntl.LOCK_NB)
        yield source_stream
    finally:
        if not source_stream.closed:
            source_stream.close()


def _open_windows_source_snapshot(source_path: Path) -> BinaryIO:
    """Open a Windows source while denying concurrent write/delete handles."""
    import ctypes
    import msvcrt  # type: ignore[import-untyped]
    from ctypes import wintypes

    ctypes_api = vars(ctypes)
    win_dll = ctypes_api["WinDLL"]
    win_error = ctypes_api["WinError"]
    get_last_error = ctypes_api["get_last_error"]
    open_osfhandle = vars(msvcrt)["open_osfhandle"]

    class _ByHandleFileInformation(ctypes.Structure):
        _fields_ = [
            ("dwFileAttributes", wintypes.DWORD),
            ("ftCreationTime", wintypes.FILETIME),
            ("ftLastAccessTime", wintypes.FILETIME),
            ("ftLastWriteTime", wintypes.FILETIME),
            ("dwVolumeSerialNumber", wintypes.DWORD),
            ("nFileSizeHigh", wintypes.DWORD),
            ("nFileSizeLow", wintypes.DWORD),
            ("nNumberOfLinks", wintypes.DWORD),
            ("nFileIndexHigh", wintypes.DWORD),
            ("nFileIndexLow", wintypes.DWORD),
        ]

    kernel32 = win_dll("kernel32", use_last_error=True)
    create_file = kernel32.CreateFileW
    create_file.argtypes = [
        wintypes.LPCWSTR,
        wintypes.DWORD,
        wintypes.DWORD,
        wintypes.LPVOID,
        wintypes.DWORD,
        wintypes.DWORD,
        wintypes.HANDLE,
    ]
    create_file.restype = wintypes.HANDLE
    get_information = kernel32.GetFileInformationByHandle
    get_information.argtypes = [wintypes.HANDLE, ctypes.POINTER(_ByHandleFileInformation)]
    get_information.restype = wintypes.BOOL
    close_handle = kernel32.CloseHandle
    close_handle.argtypes = [wintypes.HANDLE]
    close_handle.restype = wintypes.BOOL

    generic_read = 0x80000000
    file_share_read = 0x00000001
    open_existing = 3
    file_flag_sequential_scan = 0x08000000
    handle = create_file(
        str(source_path),
        generic_read,
        file_share_read,
        None,
        open_existing,
        file_flag_sequential_scan,
        None,
    )
    invalid_handle = ctypes.c_void_p(-1).value
    if handle == invalid_handle:
        raise win_error(get_last_error())

    information = _ByHandleFileInformation()
    if not get_information(handle, ctypes.byref(information)):
        error = get_last_error()
        close_handle(handle)
        raise win_error(error)
    directory_flag = getattr(stat, "FILE_ATTRIBUTE_DIRECTORY", 0x10)
    if information.dwFileAttributes & directory_flag:
        close_handle(handle)
        raise _SnapshotUnsafePath(f"Source is not a regular file: {source_path}")

    try:
        descriptor = open_osfhandle(
            int(handle),
            os.O_RDONLY | getattr(os, "O_BINARY", 0),
        )
    except Exception:
        close_handle(handle)
        raise
    try:
        return os.fdopen(descriptor, "rb", buffering=0)
    except Exception:
        os.close(descriptor)
        raise


def _apply_snapshot_metadata(destination: Path, source_stat: os.stat_result) -> None:
    """Copy portable mode/timestamps from the locked source observation."""
    # ``destination`` is a private file created by ``mkstemp`` and was already
    # validated by identity, so following a terminal link is impossible here.
    # Windows does not implement the ``follow_symlinks`` chmod keyword.
    os.chmod(destination, stat.S_IMODE(source_stat.st_mode))
    os.utime(
        destination,
        ns=(source_stat.st_atime_ns, source_stat.st_mtime_ns),
    )


def _prepare_protective_snapshot(
    source_path: Path,
    protected_input: Path,
    cancel: CancellationTokenView | None,
) -> None:
    """Prepare privately and atomically publish a stable protective snapshot."""
    stage_path = protected_input.parent
    _validate_snapshot_paths(source_path, stage_path, protected_input)
    descriptor, temporary_name = tempfile.mkstemp(
        prefix=f".{protected_input.name}.docwen-snapshot-",
        suffix=".tmp",
        dir=stage_path,
    )
    temporary_path = Path(temporary_name)
    try:
        with _open_source_snapshot(source_path) as source_stream:
            source_stat = os.fstat(source_stream.fileno())
            if not stat.S_ISREG(source_stat.st_mode):
                raise _SnapshotUnsafePath(f"Source is not a regular file: {source_path}")
            initial_identity = _snapshot_identity(source_stat)
            with os.fdopen(descriptor, "wb", buffering=0) as destination_stream:
                descriptor = -1
                _copy_snapshot_stream(source_stream, destination_stream, cancel)
                destination_stream.flush()
                os.fsync(destination_stream.fileno())
            if _snapshot_identity(os.fstat(source_stream.fileno())) != initial_identity:
                raise _SnapshotSourceChanged(f"Source changed while preparing snapshot: {source_path}")
            _apply_snapshot_metadata(temporary_path, source_stat)
            if _is_cancel_requested(cancel):
                raise _SnapshotCancelled
            os.replace(temporary_path, protected_input)
    finally:
        if descriptor >= 0:
            os.close(descriptor)
        temporary_path.unlink(missing_ok=True)


def pre_convert(
    input_path: str,
    source_format: str,
    *,
    staging_dir: str,
    cancel: CancellationTokenView | None = None,
    backend_priority: Sequence[str] | None = None,
) -> PreConversionResult | PreConversionFailure:
    """Pre-convert a document-family non-hub file to the hub format.

    Uses the caller-supplied configured priority. If no order is supplied,
    the authoritative config defaults are used defensively. ODT excludes WPS
    because that backend is not a legal candidate for this route.

    Before invoking an external backend, the source is copied to
    ``input.{canonical_source_extension}`` inside ``staging_dir``.  Only that
    application-owned copy crosses the external-process boundary.  The
    converted file is still written as ``{original_stem}.{hub_extension}``.
    The controller owns a unique staging directory per input, so neither name
    can collide within a batch and downstream names remain tied to the user
    source rather than staging.

    Args:
        input_path: Path to the user-owned source file.
        source_format: Detected source format (e.g. ``"doc"``).
        staging_dir: Directory where the pre-converted file will be written.
        cancel: Optional read-only cancellation token.
        backend_priority: Parsed software backend ids in user-selected order.

    Returns:
        A ``PreConversionResult`` on success, otherwise a structured
        ``PreConversionFailure`` retaining cancellation and bridge detail.
    """
    category = _detect_category(source_format)
    source_suffix = _PROTECTIVE_SOURCE_SUFFIXES.get(source_format)
    if category is None or source_suffix is None:
        return PreConversionFailure(message=f"Unsupported pre-conversion source format: {source_format}")

    hub_format = _CATEGORY_RULES[category][0]
    _priority_key, default_priority = backend_priority_spec(source_format)
    selected_priority = list(backend_priority) if backend_priority is not None else default_priority
    com_candidates = dict(_DOCUMENT_CANDIDATES)
    if source_format == "odt":
        com_candidates.pop("wps_writer")
    source_path = Path(input_path)
    stem = source_path.stem
    stage_path = Path(staging_dir)
    output_path = str(stage_path / f"{stem}.{hub_format}")

    if _is_cancel_requested(cancel):
        return PreConversionFailure(message="cancelled", cancelled=True, error_type="cancelled")

    protected_input = stage_path / f"input{source_suffix}"
    try:
        _prepare_protective_snapshot(source_path, protected_input, cancel)
    except _SnapshotCancelled:
        return PreConversionFailure(message="cancelled", cancelled=True, error_type="cancelled")
    except _SnapshotSourceChanged as exc:
        if _is_cancel_requested(cancel):
            return PreConversionFailure(message="cancelled", cancelled=True, error_type="cancelled")
        return PreConversionFailure(
            message=f"Source changed during {source_format.upper()} protective snapshot: {exc}",
            error_type="conversion_failed",
            diagnostic_code="PRECONVERSION_SOURCE_CHANGED",
        )
    except _SnapshotUnsafePath as exc:
        if _is_cancel_requested(cancel):
            return PreConversionFailure(message="cancelled", cancelled=True, error_type="cancelled")
        return PreConversionFailure(
            message=f"Unsafe path for {source_format.upper()} protective snapshot: {exc}",
            error_type="conversion_failed",
            diagnostic_code="PRECONVERSION_UNSAFE_PATH",
        )
    except OSError as exc:
        if _is_cancel_requested(cancel):
            return PreConversionFailure(message="cancelled", cancelled=True, error_type="cancelled")
        return PreConversionFailure(
            message=f"Cannot prepare protective input copy for {source_format.upper()} pre-conversion: {exc}",
            error_type="conversion_failed",
            diagnostic_code="PRECONVERSION_INPUT_COPY_FAILED",
        )

    # Re-check after the atomic publication and before admitting the snapshot
    # to any external Office process.
    if _is_cancel_requested(cancel):
        return PreConversionFailure(message="cancelled", cancelled=True, error_type="cancelled")

    result = convert_with_backend_priority(
        str(protected_input),
        output_path,
        source_format=source_format,
        backend_priority=selected_priority,
        com_candidates=com_candidates,
        libreoffice_format=hub_format,
        cancel=cancel,
        failure_subject=f"Configured {source_format.upper()}→{hub_format.upper()} pre-conversion backends",
    )

    if not result.success or not result.output_path:
        message = result.message or "External office pre-conversion did not produce an output file."
        cancelled = result.cancelled
        return PreConversionFailure(
            message=message,
            cancelled=cancelled,
            error_type=_bridge_failure_error_type(result),
            diagnostic_code=result.error_code,
            cleanup_message=result.cleanup_message,
            cleanup_failed=result.cleanup_failed,
        )

    return PreConversionResult(
        pre_converted_path=result.output_path,
        original_source_format=source_format,
        backend=result.backend or "unknown",
    )

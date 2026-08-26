"""Generic external-office bridge used by plugins.

This module keeps the low-level COM / LibreOffice integration in one place.
Callers provide route-specific SaveAs format codes and application types.
"""

from __future__ import annotations

import contextlib
import csv
import io
import multiprocessing
import os
import shutil
import signal
import stat
import subprocess
import sys
import tempfile
import time
import zipfile
from collections.abc import Iterable, Mapping
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Literal, Protocol

from docwen_core.detection import SUPPORTED_EXTENSION_FORMATS

AppType = Literal["word", "excel", "powerpoint"]

COM_CONVERSION_TIMEOUT_S = 90.0
LIBREOFFICE_CONVERSION_TIMEOUT_S = 90.0
_OFFICE_SNAPSHOT_CHUNK_SIZE = 1024 * 1024
_COM_POLL_INTERVAL_S = 0.2
_COM_WORKER_STOP_TIMEOUT_S = 2.0
_LIBREOFFICE_PROFILE_CLEANUP_TIMEOUT_S = 5.0
_LIBREOFFICE_PROFILE_CLEANUP_RETRY_S = 0.1
_OFFICE_WORKSPACE_PREFIX = ".dw-office-"
_OFFICE_PUBLICATION_PREFIX = ".dw-publish-"
_LIBREOFFICE_PROFILE_PREFIX = ".dw-lo-profile-"
_LIBREOFFICE_OUTPUT_PREFIX = ".dw-lo-output-"
_LIBREOFFICE_RESULT_PREFIX = ".dw-lo-result-"
_LIBREOFFICE_TIMES_FORMULA_PROFILE = """<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry"
           xmlns:xs="http://www.w3.org/2001/XMLSchema"
           xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <item oor:path="/org.openoffice.Office.Common/Font/Substitution">
    <prop oor:name="Replacement" oor:op="fuse"><value>true</value></prop>
  </item>
  <item oor:path="/org.openoffice.Office.Common/Font/Substitution/FontPairs">
    <node oor:name="docwen-times-formula-fallback" oor:op="replace">
      <prop oor:name="ReplaceFont" oor:op="fuse"><value>Liberation Serif</value></prop>
      <prop oor:name="SubstituteFont" oor:op="fuse"><value>Times New Roman</value></prop>
      <prop oor:name="Always" oor:op="fuse"><value>true</value></prop>
      <prop oor:name="OnScreenOnly" oor:op="fuse"><value>false</value></prop>
    </node>
  </item>
</oor:items>
"""
_WINDOWS_SOFFICE_APP_PATH = r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\soffice.exe"


class _ComWorkerProcess(Protocol):
    def is_alive(self) -> bool: ...

    def join(self, timeout: float | None = None) -> None: ...

    def terminate(self) -> None: ...

    def kill(self) -> None: ...


@dataclass(slots=True)
class BridgeCandidate:
    name: str
    prog_id: str
    save_format: int
    app_type: AppType
    suppress_new_revisions: bool = False


@dataclass(slots=True)
class BridgeResult:
    success: bool
    output_path: str | None = None
    backend: str = ""
    message: str = ""
    attempted_backend_ids: tuple[str, ...] = ()
    available_backend_ids: tuple[str, ...] = ()
    cancelled: bool = False
    error_code: str = ""
    cleanup_message: str = ""
    cleanup_failed: bool = False


@dataclass(frozen=True, slots=True)
class _SourceIdentity:
    device: int
    inode: int
    size: int
    modified_ns: int


class _OfficeSnapshotCancelled(Exception):
    """The caller cancelled while the private Office input was copied."""


class _OfficeSnapshotSourceChanged(OSError):
    """The user-owned source changed while its private copy was made."""


class _OfficeSourceFormatUnsupported(ValueError):
    """No canonical external-Office filename exists for the admitted format."""


class _OfficeOutputExtensionRequired(ValueError):
    """External Office requires a concrete target filename extension."""


def _is_com_candidate_available(candidate: BridgeCandidate) -> bool:
    if sys.platform != "win32":
        return False
    try:
        import winreg

        with winreg.OpenKey(
            winreg.HKEY_CLASSES_ROOT,
            rf"{candidate.prog_id}\CLSID",
        ):
            return True
    except (ImportError, OSError):
        return False


def _existing_file_path(candidate: str | os.PathLike[str] | None) -> str | None:
    if not candidate:
        return None
    try:
        path = Path(os.path.expandvars(os.fspath(candidate).strip().strip('"'))).expanduser()
        return str(path) if path.is_file() else None
    except (OSError, TypeError, ValueError):
        return None


def _read_windows_soffice_app_path() -> str | None:
    if sys.platform != "win32":
        return None
    try:
        import winreg
    except ImportError:
        return None

    view_flags = [0]
    for name in ("KEY_WOW64_64KEY", "KEY_WOW64_32KEY"):
        flag = int(getattr(winreg, name, 0))
        if flag and flag not in view_flags:
            view_flags.append(flag)

    for root in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
        for view_flag in view_flags:
            try:
                with winreg.OpenKey(
                    root,
                    _WINDOWS_SOFFICE_APP_PATH,
                    0,
                    winreg.KEY_READ | view_flag,
                ) as key:
                    value, _value_type = winreg.QueryValueEx(key, "")
            except OSError:
                continue
            if isinstance(value, str) and value.strip():
                return value
    return None


def _find_windows_registered_soffice() -> str | None:
    registered = _existing_file_path(_read_windows_soffice_app_path())
    if registered:
        return registered

    roots: list[str] = []
    for name in ("ProgramW6432", "ProgramFiles", "ProgramFiles(x86)"):
        value = os.environ.get(name)
        if value and value not in roots:
            roots.append(value)
    for root in roots:
        candidate = _existing_file_path(Path(root) / "LibreOffice" / "program" / "soffice.exe")
        if candidate:
            return candidate
    return None


def find_soffice_path() -> str | None:
    soffice = _existing_file_path(shutil.which("soffice"))
    if soffice:
        return soffice

    if sys.platform == "win32":
        return _find_windows_registered_soffice()

    if sys.platform == "darwin":
        mac_paths = [
            "/Applications/LibreOffice.app/Contents/MacOS/soffice",
            str(Path.home() / "Applications/LibreOffice.app/Contents/MacOS/soffice"),
        ]
        for path in mac_paths:
            soffice = _existing_file_path(path)
            if soffice:
                return soffice

    return None


def _import_win32() -> tuple[Any | None, Any | None]:
    try:
        import pythoncom  # type: ignore
        import win32com.client  # type: ignore
    except ImportError:
        return None, None
    return pythoncom, win32com.client


def _is_cancel_requested(cancel: object | None) -> bool:
    if cancel is None:
        return False

    is_set = getattr(cancel, "is_set", None)
    if callable(is_set):
        return bool(is_set())

    is_cancelled = getattr(cancel, "is_cancelled", None)
    if callable(is_cancelled):
        return bool(is_cancelled())
    return bool(is_cancelled)


def _get_com_app_pid(app: object) -> int | None:
    try:
        import win32process  # type: ignore
    except ImportError:
        return None

    try:
        hwnd = getattr(app, "Hwnd", None)
        if not hwnd:
            return None
        _thread_id, process_id = win32process.GetWindowThreadProcessId(hwnd)
        return int(process_id) if process_id else None
    except Exception:
        return None


def _process_exists(process_id: int) -> bool:
    if sys.platform == "win32":
        try:
            proc = subprocess.run(
                ["tasklist", "/FI", f"PID eq {process_id}", "/FO", "CSV", "/NH"],
                check=False,
                text=True,
                capture_output=True,
                timeout=5,
            )
        except Exception:
            return False
        for row in csv.reader(io.StringIO(proc.stdout)):
            if len(row) >= 2:
                with contextlib.suppress(ValueError):
                    if int(row[1]) == process_id:
                        return True
        return False

    try:
        os.kill(process_id, 0)
    except OSError:
        return False
    return True


def _wait_for_process_exit(process_id: int, *, timeout_s: float = 5.0) -> bool:
    deadline = time.monotonic() + timeout_s
    while time.monotonic() < deadline:
        if not _process_exists(process_id):
            return True
        time.sleep(0.1)
    return not _process_exists(process_id)


def _terminate_process(process_id: int) -> None:
    with contextlib.suppress(Exception):
        os.kill(process_id, signal.SIGTERM)


def _office_process_names(*, prog_id: str, app_type: AppType) -> tuple[str, ...]:
    prog_id_lower = prog_id.lower()
    if prog_id_lower.startswith("kwps"):
        return ("WPS.EXE",)
    if prog_id_lower.startswith("ket"):
        return ("KET.EXE",)
    if prog_id_lower.startswith("kwpp"):
        return ("KWPP.EXE",)
    if app_type == "word":
        return ("WINWORD.EXE",)
    if app_type == "excel":
        return ("EXCEL.EXE",)
    return ("POWERPNT.EXE",)


def _snapshot_process_ids(process_names: tuple[str, ...]) -> set[int]:
    if sys.platform != "win32":
        return set()

    process_ids: set[int] = set()
    for process_name in process_names:
        try:
            proc = subprocess.run(
                ["tasklist", "/FI", f"IMAGENAME eq {process_name}", "/FO", "CSV", "/NH"],
                check=False,
                text=True,
                capture_output=True,
                timeout=5,
            )
        except Exception:
            continue
        for row in csv.reader(io.StringIO(proc.stdout)):
            if len(row) < 2 or row[0].upper() != process_name.upper():
                continue
            with contextlib.suppress(ValueError):
                process_ids.add(int(row[1]))
    return process_ids


def _try_com_conversion(
    input_path: str,
    output_path: str,
    *,
    prog_id: str,
    save_format: int,
    app_type: AppType,
    suppress_new_revisions: bool = False,
) -> str | None:
    pythoncom, win32_client = _import_win32()
    if pythoncom is None or win32_client is None:
        return None

    app: Any = None
    doc_or_wb: Any = None
    app_pid: int | None = None
    office_process_names: tuple[str, ...] = ()
    office_before_pids: set[int] = set()
    try:
        pythoncom.CoInitialize()  # pyright: ignore[reportAttributeAccessIssue]
        office_process_names = _office_process_names(prog_id=prog_id, app_type=app_type)
        office_before_pids = _snapshot_process_ids(office_process_names)
        dispatch_ex = getattr(win32_client, "DispatchEx", None)
        if callable(dispatch_ex):
            try:
                app = dispatch_ex(prog_id)
                app_pid = _get_com_app_pid(app)
            except Exception:
                app = win32_client.Dispatch(prog_id)  # pyright: ignore[reportAttributeAccessIssue]
        else:
            app = win32_client.Dispatch(prog_id)  # pyright: ignore[reportAttributeAccessIssue]
        with contextlib.suppress(Exception):
            app.Visible = False
        with contextlib.suppress(Exception):
            app.DisplayAlerts = False

        if app_type == "excel":
            doc_or_wb = app.Workbooks.Open(
                str(Path(input_path).resolve()),
                UpdateLinks=0,
                ReadOnly=True,
                IgnoreReadOnlyRecommended=True,
                AddToMru=False,
            )
        elif app_type == "word":
            doc_or_wb = app.Documents.Open(
                str(Path(input_path).resolve()),
                ReadOnly=True,
                ConfirmConversions=False,
                AddToRecentFiles=False,
            )
        else:
            try:
                doc_or_wb = app.Presentations.Open(str(Path(input_path).resolve()), True, False, False)
            except Exception:
                # Some PowerPoint-compatible COM servers expose fewer optional
                # arguments.  Retrying with the explicit ReadOnly argument is
                # safe; falling back to ``Open(path)`` would silently grant the
                # external process write access to the private input.
                doc_or_wb = app.Presentations.Open(str(Path(input_path).resolve()), True)

        if app_type == "word" and suppress_new_revisions:
            with contextlib.suppress(Exception):
                doc_or_wb.TrackRevisions = False

        resolved_output = str(Path(output_path).resolve())
        if app_type == "powerpoint":
            doc_or_wb.SaveAs(resolved_output, save_format)
        elif app_type == "excel" and save_format == 57:
            try:
                doc_or_wb.ExportAsFixedFormat(0, resolved_output)
            except Exception:
                doc_or_wb.SaveAs(resolved_output, FileFormat=save_format)
        else:
            doc_or_wb.SaveAs(resolved_output, FileFormat=save_format)

        if Path(resolved_output).exists():
            return resolved_output
        return None
    except Exception:
        return None
    finally:
        if doc_or_wb is not None:
            with contextlib.suppress(Exception):
                if app_type == "powerpoint":
                    doc_or_wb.Close()
                else:
                    doc_or_wb.Close(SaveChanges=False)
        if app is not None:
            if app_pid is None:
                app_pid = _get_com_app_pid(app)
            with contextlib.suppress(Exception):
                app.Quit()
            process_ids: set[int] = set()
            if app_pid is not None:
                process_ids.add(app_pid)
            if office_process_names:
                process_ids.update(_snapshot_process_ids(office_process_names) - office_before_pids)
            for _ in range(2):
                for process_id in process_ids:
                    if not _wait_for_process_exit(process_id):
                        _terminate_process(process_id)
                if not office_process_names:
                    break
                time.sleep(0.2)
                process_ids = _snapshot_process_ids(office_process_names) - office_before_pids
                if not process_ids:
                    break
        if pythoncom is not None:
            with contextlib.suppress(Exception):
                pythoncom.CoUninitialize()  # pyright: ignore[reportAttributeAccessIssue]


def _com_conversion_worker(
    result_connection: Any,
    input_path: str,
    output_path: str,
    prog_id: str,
    save_format: int,
    app_type: AppType,
    suppress_new_revisions: bool,
) -> None:
    """Run one synchronous COM attempt behind a killable process boundary."""
    result: str | None = None
    try:
        result = _try_com_conversion(
            input_path,
            output_path,
            prog_id=prog_id,
            save_format=save_format,
            app_type=app_type,
            suppress_new_revisions=suppress_new_revisions,
        )
    finally:
        with contextlib.suppress(Exception):
            result_connection.send(result)
        result_connection.close()


def _start_com_conversion_process(
    input_path: str,
    output_path: str,
    *,
    prog_id: str,
    save_format: int,
    app_type: AppType,
    suppress_new_revisions: bool,
) -> tuple[_ComWorkerProcess, Any]:
    context = multiprocessing.get_context("spawn")
    result_connection, worker_connection = context.Pipe(duplex=False)
    process = context.Process(
        target=_com_conversion_worker,
        args=(
            worker_connection,
            input_path,
            output_path,
            prog_id,
            save_format,
            app_type,
            suppress_new_revisions,
        ),
        daemon=True,
    )
    try:
        process.start()
    except BaseException:
        result_connection.close()
        worker_connection.close()
        raise
    worker_connection.close()
    return process, result_connection


def _stop_com_worker(process: _ComWorkerProcess) -> None:
    if process.is_alive():
        with contextlib.suppress(Exception):
            process.terminate()
    with contextlib.suppress(Exception):
        process.join(timeout=_COM_WORKER_STOP_TIMEOUT_S)
    if process.is_alive():
        with contextlib.suppress(Exception):
            process.kill()
        with contextlib.suppress(Exception):
            process.join(timeout=_COM_WORKER_STOP_TIMEOUT_S)


def _terminate_new_office_processes(
    process_names: tuple[str, ...],
    before_process_ids: set[int],
) -> None:
    """Clean Office processes created by a worker that could not run ``finally``."""
    for _ in range(2):
        new_process_ids = _snapshot_process_ids(process_names) - before_process_ids
        for process_id in new_process_ids:
            _terminate_process(process_id)
        if not new_process_ids:
            break
        time.sleep(0.2)


def _try_com_conversion_bounded(
    input_path: str,
    output_path: str,
    *,
    prog_id: str,
    save_format: int,
    app_type: AppType,
    suppress_new_revisions: bool = False,
    cancel: object | None = None,
    timeout_s: float = COM_CONVERSION_TIMEOUT_S,
) -> str | None:
    """Run COM conversion with a hard timeout and cooperative parent cancellation.

    The cancellation view deliberately stays in the parent because core's token is
    thread-safe, not process-safe.  Only primitive conversion arguments cross the
    spawned-process boundary.
    """
    if _is_cancel_requested(cancel) or timeout_s <= 0:
        return None
    if sys.platform != "win32":
        return _try_com_conversion(
            input_path,
            output_path,
            prog_id=prog_id,
            save_format=save_format,
            app_type=app_type,
            suppress_new_revisions=suppress_new_revisions,
        )

    process_names = _office_process_names(prog_id=prog_id, app_type=app_type)
    before_process_ids = _snapshot_process_ids(process_names)
    try:
        process, result_connection = _start_com_conversion_process(
            input_path,
            output_path,
            prog_id=prog_id,
            save_format=save_format,
            app_type=app_type,
            suppress_new_revisions=suppress_new_revisions,
        )
    except Exception:
        return None

    deadline = time.monotonic() + timeout_s
    forced_stop = False
    try:
        while process.is_alive():
            if _is_cancel_requested(cancel) or time.monotonic() >= deadline:
                forced_stop = True
                _stop_com_worker(process)
                return None
            if result_connection.poll(_COM_POLL_INTERVAL_S):
                with contextlib.suppress(EOFError, OSError):
                    result = result_connection.recv()
                    process.join(timeout=_COM_WORKER_STOP_TIMEOUT_S)
                    if process.is_alive():
                        _stop_com_worker(process)
                    return result if isinstance(result, str) else None

        process.join(timeout=_COM_WORKER_STOP_TIMEOUT_S)
        if result_connection.poll():
            with contextlib.suppress(EOFError, OSError):
                result = result_connection.recv()
                return result if isinstance(result, str) else None
        return None
    finally:
        result_connection.close()
        if process.is_alive():
            forced_stop = True
            _stop_com_worker(process)
        if forced_stop:
            with contextlib.suppress(OSError):
                Path(output_path).unlink(missing_ok=True)
            _terminate_new_office_processes(process_names, before_process_ids)


def _remove_temp_tree(
    tree_path: Path,
    *,
    timeout_s: float = _LIBREOFFICE_PROFILE_CLEANUP_TIMEOUT_S,
    retry_interval_s: float = _LIBREOFFICE_PROFILE_CLEANUP_RETRY_S,
) -> bool:
    deadline = time.monotonic() + max(timeout_s, 0.0)
    while True:
        try:
            shutil.rmtree(tree_path)
        except FileNotFoundError:
            return True
        except OSError:
            if time.monotonic() >= deadline:
                return False
            time.sleep(max(retry_interval_s, 0.0))
        else:
            return True


def _remove_libreoffice_profile(
    profile_path: Path,
    *,
    timeout_s: float = _LIBREOFFICE_PROFILE_CLEANUP_TIMEOUT_S,
    retry_interval_s: float = _LIBREOFFICE_PROFILE_CLEANUP_RETRY_S,
) -> bool:
    return _remove_temp_tree(
        profile_path,
        timeout_s=timeout_s,
        retry_interval_s=retry_interval_s,
    )


def _preferred_extension_for_format(source_format: str) -> str | None:
    """Return the canonical filename extension from the shared ingress registry."""

    normalized = str(source_format or "").strip().casefold()
    return next(
        (
            extension
            for extension, declared_format in SUPPORTED_EXTENSION_FORMATS.items()
            if declared_format == normalized
        ),
        None,
    )


def _source_identity(file_stat: os.stat_result) -> _SourceIdentity:
    return _SourceIdentity(
        device=int(file_stat.st_dev),
        inode=int(file_stat.st_ino),
        size=int(file_stat.st_size),
        modified_ns=int(file_stat.st_mtime_ns),
    )


def _paths_resolve_same(input_path: str, output_path: str) -> bool:
    source = Path(input_path).resolve()
    output = Path(output_path).resolve()
    if os.path.normcase(str(source)) == os.path.normcase(str(output)):
        return True
    if source.exists() and output.exists():
        with contextlib.suppress(OSError):
            return source.samefile(output)
    return False


def _copy_private_office_input(
    source: Path,
    destination: Path,
    *,
    cancel: object | None,
) -> None:
    """Copy one stable source observation without ever sharing its inode."""

    if _is_cancel_requested(cancel):
        raise _OfficeSnapshotCancelled
    path_identity = _source_identity(source.stat())
    with source.open("rb", buffering=0) as source_stream:
        opened_identity = _source_identity(os.fstat(source_stream.fileno()))
        if opened_identity != path_identity or not stat.S_ISREG(os.fstat(source_stream.fileno()).st_mode):
            raise _OfficeSnapshotSourceChanged(f"Source identity changed before copy: {source}")

        with destination.open("xb", buffering=0) as destination_stream:
            while True:
                if _is_cancel_requested(cancel):
                    raise _OfficeSnapshotCancelled
                chunk = source_stream.read(_OFFICE_SNAPSHOT_CHUNK_SIZE)
                if not chunk:
                    break
                destination_stream.write(chunk)
            destination_stream.flush()
            os.fsync(destination_stream.fileno())

        if _is_cancel_requested(cancel):
            raise _OfficeSnapshotCancelled
        if _source_identity(os.fstat(source_stream.fileno())) != path_identity:
            raise _OfficeSnapshotSourceChanged(f"Source changed while copying: {source}")

    if _source_identity(source.stat()) != path_identity:
        raise _OfficeSnapshotSourceChanged(f"Source path changed while copying: {source}")
    if destination.stat().st_size != path_identity.size:
        raise _OfficeSnapshotSourceChanged(f"Private copy size does not match source: {source}")


@dataclass(slots=True)
class _OfficeConversionOwner:
    """Own the complete private source/backend-output lifecycle."""

    source: Path
    output: Path
    source_format: str
    cancel: object | None
    workspace: Path | None = None
    input_snapshot: Path | None = None
    backend_output: Path | None = None
    publication_output: Path | None = None

    def prepare(self) -> tuple[Path, Path]:
        preferred_extension = _preferred_extension_for_format(self.source_format)
        if preferred_extension is None:
            raise _OfficeSourceFormatUnsupported(
                f"Unsupported admitted source format for external Office: {self.source_format or 'empty'}."
            )
        source_stat = self.source.stat()
        if not stat.S_ISREG(source_stat.st_mode):
            raise OSError(f"External Office input is not a regular file: {self.source}")
        if not self.output.suffix:
            raise _OfficeOutputExtensionRequired("External Office output requires a concrete filename extension.")

        self.output.parent.mkdir(parents=True, exist_ok=True)
        self.workspace = Path(tempfile.mkdtemp(prefix=_OFFICE_WORKSPACE_PREFIX, dir=self.output.parent))
        self.input_snapshot = self.workspace / f"input{preferred_extension}"
        self.backend_output = self.workspace / f"output{self.output.suffix}"
        _copy_private_office_input(
            self.source,
            self.input_snapshot,
            cancel=self.cancel,
        )
        return self.input_snapshot, self.backend_output

    def stage_backend_output(self) -> Path:
        backend_output = self.backend_output
        if backend_output is None or not backend_output.is_file():
            raise FileNotFoundError("External Office reported success without its private output file.")
        descriptor, publication_name = tempfile.mkstemp(
            prefix=_OFFICE_PUBLICATION_PREFIX,
            suffix=self.output.suffix,
            dir=self.output.parent,
        )
        self.publication_output = Path(publication_name)
        os.close(descriptor)
        os.replace(backend_output, self.publication_output)
        return self.publication_output

    def cleanup_workspace(self) -> tuple[bool, str]:
        if self.workspace is None:
            return True, "No private Office workspace was created."
        try:
            removed = _remove_temp_tree(self.workspace)
        except Exception as exc:
            return False, f"Private Office workspace cleanup raised an exception: {exc}"
        if removed:
            return True, "Private Office workspace cleaned up."
        return False, f"Private Office workspace cleanup failed: {self.workspace}"

    def discard_publication(self) -> tuple[bool, str]:
        if self.publication_output is None:
            return True, "No private Office publication file was created."
        try:
            self.publication_output.unlink(missing_ok=True)
        except Exception as exc:
            return False, f"Private Office publication cleanup failed: {exc}"
        return True, "Private Office publication file cleaned up."

    def publish(self) -> None:
        if self.publication_output is None:
            raise FileNotFoundError("Private Office publication file is missing.")
        os.replace(self.publication_output, self.output)


def _writer_conversion_needs_times_formula_fallback(input_path: Path, convert_to: str) -> bool:
    """Keep Writer formula exports stable without overriding explicit Liberation text."""
    output_extension = convert_to.partition(":")[0].strip().lstrip(".").casefold()
    if output_extension not in {"odt", "pdf"}:
        return False

    suffix = input_path.suffix.casefold()
    if suffix not in {".docx", ".odt"}:
        return False

    try:
        with zipfile.ZipFile(input_path) as archive:
            names = set(archive.namelist())
            if suffix == ".docx":
                document = archive.read("word/document.xml")
                has_formula = (
                    b"schemas.openxmlformats.org/officeDocument/2006/math" in document and b"oMath" in document
                )
                font_members = (
                    "word/document.xml",
                    "word/fontTable.xml",
                    "word/settings.xml",
                    "word/styles.xml",
                    "word/theme/theme1.xml",
                )
            else:
                manifest = archive.read("META-INF/manifest.xml")
                has_formula = b"application/vnd.oasis.opendocument.formula" in manifest
                font_members = ("content.xml", "styles.xml")

            font_xml = b"\n".join(archive.read(name) for name in font_members if name in names).lower()
    except (KeyError, OSError, zipfile.BadZipFile):
        return False

    return has_formula and b"times new roman" in font_xml and b"liberation serif" not in font_xml


def _prepare_libreoffice_profile(profile_path: Path, *, times_formula_fallback: bool) -> None:
    if not times_formula_fallback:
        return
    profile_user = profile_path / "user"
    profile_user.mkdir(parents=True, exist_ok=True)
    (profile_user / "registrymodifications.xcu").write_text(
        _LIBREOFFICE_TIMES_FORMULA_PROFILE,
        encoding="utf-8",
        newline="\n",
    )


def _run_libreoffice_process(
    cmd: list[str],
    *,
    cancel: object | None = None,
    timeout_s: float = LIBREOFFICE_CONVERSION_TIMEOUT_S,
) -> bool:
    proc: subprocess.Popen[str] | None = None
    try:
        proc = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
        )
        deadline = time.monotonic() + max(timeout_s, 0.0)
        while True:
            return_code = proc.poll()
            if return_code is not None:
                break
            if _is_cancel_requested(cancel):
                proc.terminate()
                try:
                    proc.wait(timeout=5)
                except subprocess.TimeoutExpired:
                    proc.kill()
                    proc.wait(timeout=5)
                return False
            if time.monotonic() >= deadline:
                proc.kill()
                proc.wait(timeout=5)
                return False
            time.sleep(0.2)
        proc.communicate()
    except OSError:
        if proc is not None:
            with contextlib.suppress(Exception):
                proc.kill()
        return False
    return proc.returncode == 0


def _try_libreoffice_conversion(
    input_path: str,
    output_path: str,
    *,
    convert_to: str,
    cancel: object | None = None,
    timeout_s: float = LIBREOFFICE_CONVERSION_TIMEOUT_S,
) -> str | None:
    if _is_cancel_requested(cancel):
        return None

    soffice = find_soffice_path()
    if not soffice:
        return None

    output = Path(output_path).resolve()
    profile_dir: Path | None = None
    conversion_dir: Path | None = None
    staged_output: Path | None = None
    process_succeeded = False
    try:
        output.parent.mkdir(parents=True, exist_ok=True)
        # LibreOffice on Windows can terminate with 0xC0000409 when the
        # UserInstallation URI is nested inside a deep request workspace.
        # Its disposable profile is process state, not an output artifact, so
        # own it directly beneath the system temp root.  Conversion output
        # stays beside the requested output to preserve same-volume staging.
        profile_dir = Path(tempfile.mkdtemp(prefix=_LIBREOFFICE_PROFILE_PREFIX))
        conversion_dir = Path(tempfile.mkdtemp(prefix=_LIBREOFFICE_OUTPUT_PREFIX, dir=output.parent))
        _prepare_libreoffice_profile(
            profile_dir,
            times_formula_fallback=_writer_conversion_needs_times_formula_fallback(Path(input_path), convert_to),
        )
        process_succeeded = _run_libreoffice_process(
            [
                soffice,
                f"-env:UserInstallation={profile_dir.resolve().as_uri()}",
                "--headless",
                "--convert-to",
                convert_to,
                "--outdir",
                str(conversion_dir),
                str(Path(input_path).resolve()),
            ],
            cancel=cancel,
            timeout_s=timeout_s,
        )
        output_extension = convert_to.partition(":")[0].strip().lstrip(".")
        generated = conversion_dir / f"{Path(input_path).stem}.{output_extension}"
        if process_succeeded and output_extension and generated.is_file() and not _is_cancel_requested(cancel):
            descriptor, staged_name = tempfile.mkstemp(
                prefix=_LIBREOFFICE_RESULT_PREFIX,
                suffix=output.suffix,
                dir=output.parent,
            )
            staged_output = Path(staged_name)
            os.close(descriptor)
            os.replace(generated, staged_output)
        else:
            process_succeeded = False
    except OSError:
        process_succeeded = False
    finally:
        profile_removed = profile_dir is None or _remove_libreoffice_profile(profile_dir)
        conversion_removed = conversion_dir is None or _remove_temp_tree(conversion_dir)

    if not profile_removed or not conversion_removed or not process_succeeded or staged_output is None:
        if staged_output is not None:
            with contextlib.suppress(OSError):
                staged_output.unlink(missing_ok=True)
        return None
    if _is_cancel_requested(cancel):
        with contextlib.suppress(OSError):
            staged_output.unlink(missing_ok=True)
        return None
    try:
        os.replace(staged_output, output)
    except OSError:
        with contextlib.suppress(OSError):
            staged_output.unlink(missing_ok=True)
        return None
    return str(output)


def _convert_with_fallback(
    input_path: str,
    output_path: str,
    *,
    com_candidates: list[BridgeCandidate],
    libreoffice_format: str | None = None,
    cancel: object | None = None,
    com_timeout_s: float = COM_CONVERSION_TIMEOUT_S,
    libreoffice_timeout_s: float = LIBREOFFICE_CONVERSION_TIMEOUT_S,
) -> BridgeResult:
    com_deadline = time.monotonic() + max(com_timeout_s, 0.0)
    for candidate in com_candidates:
        if _is_cancel_requested(cancel):
            return BridgeResult(
                False,
                message="cancelled",
                cancelled=True,
                error_code="OFFICE_CONVERSION_CANCELLED",
            )
        remaining_timeout_s = com_deadline - time.monotonic()
        if remaining_timeout_s <= 0:
            break
        result = _try_com_conversion_bounded(
            input_path,
            output_path,
            prog_id=candidate.prog_id,
            save_format=candidate.save_format,
            app_type=candidate.app_type,
            suppress_new_revisions=candidate.suppress_new_revisions,
            cancel=cancel,
            timeout_s=remaining_timeout_s,
        )
        if result:
            return BridgeResult(True, output_path=result, backend=candidate.name)
        if _is_cancel_requested(cancel):
            return BridgeResult(
                False,
                message="cancelled",
                cancelled=True,
                error_code="OFFICE_CONVERSION_CANCELLED",
            )

    if libreoffice_format:
        result = _try_libreoffice_conversion(
            input_path,
            output_path,
            convert_to=libreoffice_format,
            cancel=cancel,
            timeout_s=libreoffice_timeout_s,
        )
        if result:
            return BridgeResult(True, output_path=result, backend="LibreOffice")
        if _is_cancel_requested(cancel):
            return BridgeResult(
                False,
                message="cancelled",
                cancelled=True,
                error_code="OFFICE_CONVERSION_CANCELLED",
            )

    return BridgeResult(
        False,
        message=("No external office backend succeeded. Install Microsoft Office/WPS (Windows COM) or LibreOffice."),
        error_code="OFFICE_BACKEND_FAILED",
    )


def _convert_with_backend_priority_canonical_input(
    input_path: str,
    output_path: str,
    *,
    backend_priority: Iterable[str],
    com_candidates: Mapping[str, BridgeCandidate],
    libreoffice_format: str | None = None,
    cancel: object | None = None,
    com_timeout_s: float = COM_CONVERSION_TIMEOUT_S,
    libreoffice_timeout_s: float = LIBREOFFICE_CONVERSION_TIMEOUT_S,
    failure_subject: str = "Configured external Office backends",
) -> BridgeResult:
    """Try named COM/LibreOffice backends in the caller-provided order.

    Route owners resolve configuration keys and provide their legal COM
    candidates; core remains the only owner of external process execution.
    Unknown backend ids are skipped with an auditable failure message, and
    LibreOffice is attempted only where it appears in ``backend_priority``.
    """
    if _is_cancel_requested(cancel):
        return BridgeResult(
            False,
            message="cancelled",
            cancelled=True,
            error_code="OFFICE_CONVERSION_CANCELLED",
        )

    messages: list[str] = []
    attempted_backend_ids: list[str] = []
    available_backend_ids: list[str] = []
    remaining_com_timeout_s = max(com_timeout_s, 0.0)
    for raw_backend_id in backend_priority:
        if _is_cancel_requested(cancel):
            return BridgeResult(
                False,
                message="cancelled",
                attempted_backend_ids=tuple(attempted_backend_ids),
                available_backend_ids=tuple(available_backend_ids),
                cancelled=True,
                error_code="OFFICE_CONVERSION_CANCELLED",
            )
        backend_id = str(raw_backend_id)
        if backend_id == "libreoffice":
            if libreoffice_format is None:
                messages.append("libreoffice: unavailable for this route")
                continue
            attempted_backend_ids.append(backend_id)
            if find_soffice_path() is not None:
                available_backend_ids.append(backend_id)
            result = _convert_with_fallback(
                input_path,
                output_path,
                com_candidates=[],
                libreoffice_format=libreoffice_format,
                cancel=cancel,
                libreoffice_timeout_s=libreoffice_timeout_s,
            )
        else:
            candidate = com_candidates.get(backend_id)
            if candidate is None:
                messages.append(f"unsupported backend id: {backend_id}")
                continue
            attempted_backend_ids.append(backend_id)
            if _is_com_candidate_available(candidate):
                available_backend_ids.append(backend_id)
            if remaining_com_timeout_s <= 0:
                messages.append(f"{backend_id}: COM timeout budget exhausted")
                continue
            started_at = time.monotonic()
            result = _convert_with_fallback(
                input_path,
                output_path,
                com_candidates=[candidate],
                libreoffice_format=None,
                cancel=cancel,
                com_timeout_s=remaining_com_timeout_s,
                libreoffice_timeout_s=libreoffice_timeout_s,
            )
            remaining_com_timeout_s = max(
                remaining_com_timeout_s - max(time.monotonic() - started_at, 0.0),
                0.0,
            )

        if result.success and result.output_path is not None:
            if backend_id not in available_backend_ids:
                available_backend_ids.append(backend_id)
            result.attempted_backend_ids = tuple(attempted_backend_ids)
            result.available_backend_ids = tuple(available_backend_ids)
            return result
        if result.cancelled:
            result.attempted_backend_ids = tuple(attempted_backend_ids)
            result.available_backend_ids = tuple(available_backend_ids)
            return result
        if result.message:
            messages.append(f"{backend_id}: {result.message}")

    message = f"{failure_subject} did not succeed."
    if messages:
        message = f"{message} {'; '.join(messages)}"
    return BridgeResult(
        False,
        message=message,
        attempted_backend_ids=tuple(attempted_backend_ids),
        available_backend_ids=tuple(available_backend_ids),
        error_code="OFFICE_BACKEND_FAILED",
    )


def convert_with_backend_priority(
    input_path: str,
    output_path: str,
    *,
    source_format: str,
    backend_priority: Iterable[str],
    com_candidates: Mapping[str, BridgeCandidate],
    libreoffice_format: str | None = None,
    cancel: object | None = None,
    com_timeout_s: float = COM_CONVERSION_TIMEOUT_S,
    libreoffice_timeout_s: float = LIBREOFFICE_CONVERSION_TIMEOUT_S,
    failure_subject: str = "Configured external Office backends",
) -> BridgeResult:
    """Convert through one private, content-named snapshot and publish atomically.

    The external backend never receives ``input_path`` itself, including when
    its suffix already agrees with ``source_format``.  Both the input snapshot
    and backend output live in an owner-controlled workspace next to the
    requested staging output.  The requested output is published only after
    the workspace is proven clean.
    """

    normalized_format = str(source_format or "").strip().casefold()
    no_cleanup = "No private Office workspace was created."
    try:
        source = Path(input_path).resolve()
        output = Path(output_path).resolve()
        same_path = _paths_resolve_same(input_path, output_path)
    except (OSError, RuntimeError, ValueError) as exc:
        return BridgeResult(
            False,
            message=f"Cannot resolve external Office paths: {exc}",
            error_code="OFFICE_PATH_RESOLUTION_FAILED",
            cleanup_message=no_cleanup,
        )
    if same_path:
        return BridgeResult(
            False,
            message="External Office input and output must resolve to different paths.",
            error_code="OFFICE_INPUT_OUTPUT_CONFLICT",
            cleanup_message=no_cleanup,
        )
    if _is_cancel_requested(cancel):
        return BridgeResult(
            False,
            message="cancelled",
            cancelled=True,
            error_code="OFFICE_CONVERSION_CANCELLED",
            cleanup_message=no_cleanup,
        )

    owner = _OfficeConversionOwner(
        source=source,
        output=output,
        source_format=normalized_format,
        cancel=cancel,
    )
    result: BridgeResult
    try:
        private_input, private_output = owner.prepare()
    except _OfficeSnapshotCancelled:
        result = BridgeResult(
            False,
            message="cancelled",
            cancelled=True,
            error_code="OFFICE_CONVERSION_CANCELLED",
        )
    except _OfficeSnapshotSourceChanged as exc:
        result = BridgeResult(
            False,
            message=str(exc),
            error_code="OFFICE_SOURCE_CHANGED",
        )
    except _OfficeSourceFormatUnsupported as exc:
        result = BridgeResult(
            False,
            message=str(exc),
            error_code="OFFICE_SOURCE_FORMAT_UNSUPPORTED",
        )
    except _OfficeOutputExtensionRequired as exc:
        result = BridgeResult(
            False,
            message=str(exc),
            error_code="OFFICE_OUTPUT_EXTENSION_REQUIRED",
        )
    except (OSError, RuntimeError) as exc:
        result = BridgeResult(
            False,
            message=f"Cannot prepare private external Office input: {exc}",
            error_code="OFFICE_SNAPSHOT_PREPARATION_FAILED",
        )
    else:
        try:
            result = _convert_with_backend_priority_canonical_input(
                str(private_input),
                str(private_output),
                backend_priority=backend_priority,
                com_candidates=com_candidates,
                libreoffice_format=libreoffice_format,
                cancel=cancel,
                com_timeout_s=com_timeout_s,
                libreoffice_timeout_s=libreoffice_timeout_s,
                failure_subject=failure_subject,
            )
        except Exception as exc:
            result = BridgeResult(
                False,
                message=f"External Office backend raised an exception: {exc}",
                error_code="OFFICE_BACKEND_EXCEPTION",
            )
        if result.success and _is_cancel_requested(cancel):
            result = BridgeResult(
                False,
                backend=result.backend,
                message="cancelled",
                attempted_backend_ids=result.attempted_backend_ids,
                available_backend_ids=result.available_backend_ids,
                cancelled=True,
                error_code="OFFICE_CONVERSION_CANCELLED",
            )
        if result.success:
            try:
                owner.stage_backend_output()
            except (OSError, RuntimeError) as exc:
                result = BridgeResult(
                    False,
                    backend=result.backend,
                    message=f"Cannot stage external Office output: {exc}",
                    attempted_backend_ids=result.attempted_backend_ids,
                    available_backend_ids=result.available_backend_ids,
                    error_code="OFFICE_OUTPUT_STAGE_FAILED",
                )

    workspace_cleaned, workspace_message = owner.cleanup_workspace()
    if not workspace_cleaned:
        _publication_cleaned, publication_message = owner.discard_publication()
        cleanup_message = f"{workspace_message} {publication_message}"
        if result.cancelled:
            result.cleanup_message = cleanup_message
            result.cleanup_failed = True
            return result
        if not result.success:
            result.cleanup_message = cleanup_message
            result.cleanup_failed = True
            if not result.error_code:
                result.error_code = "OFFICE_SNAPSHOT_CLEANUP_FAILED"
            return result
        return BridgeResult(
            False,
            backend=result.backend,
            message="Private external Office workspace could not be cleaned up.",
            attempted_backend_ids=result.attempted_backend_ids,
            available_backend_ids=result.available_backend_ids,
            error_code="OFFICE_SNAPSHOT_CLEANUP_FAILED",
            cleanup_message=cleanup_message,
            cleanup_failed=True,
        )

    result.cleanup_message = workspace_message
    if not result.success:
        publication_cleaned, publication_message = owner.discard_publication()
        if not publication_cleaned:
            result.cleanup_message = f"{result.cleanup_message} {publication_message}"
            result.cleanup_failed = True
        return result

    if _is_cancel_requested(cancel):
        publication_cleaned, publication_message = owner.discard_publication()
        return BridgeResult(
            False,
            backend=result.backend,
            message="cancelled",
            attempted_backend_ids=result.attempted_backend_ids,
            available_backend_ids=result.available_backend_ids,
            cancelled=True,
            error_code="OFFICE_CONVERSION_CANCELLED",
            cleanup_message=f"{workspace_message} {publication_message}",
            cleanup_failed=not publication_cleaned,
        )

    try:
        owner.publish()
    except (OSError, RuntimeError) as exc:
        publication_cleaned, publication_message = owner.discard_publication()
        cleanup_message = f"{workspace_message} {publication_message}"
        return BridgeResult(
            False,
            backend=result.backend,
            message=f"Cannot publish external Office output: {exc}",
            attempted_backend_ids=result.attempted_backend_ids,
            available_backend_ids=result.available_backend_ids,
            error_code="OFFICE_OUTPUT_PUBLISH_FAILED",
            cleanup_message=cleanup_message,
            cleanup_failed=not publication_cleaned,
        )

    result.output_path = str(output)
    return result

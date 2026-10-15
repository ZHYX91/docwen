"""Recycle validated files or directories and prove a recoverable payload exists."""

from __future__ import annotations

import os
import struct
import time
from pathlib import Path
from typing import Any, ClassVar

from tools.recycle_payload import assert_plain_chain, read_plain_bytes, require_payload_identity, snapshot_payload

_RECYCLE_FLAGS = 0x80000 | 0x20000000 | 0x100000 | 0x400 | 0x10 | 0x4


class RecycleUnconfirmedError(OSError):
    """A recycle was attempted; preserve exact candidates rather than retrying it."""

    def __init__(
        self,
        reason: str,
        *,
        path: Path,
        recovery: list[dict[str, str]],
        expected: dict[str, Any],
        source_absent: bool | None,
        observation_errors: list[str],
        operation_error: str | None = None,
    ):
        super().__init__(f"{reason}:{path}" + (f":{operation_error}" if operation_error else ""))
        self.recovery = recovery
        self.expected_identity = expected
        self.source_absent = source_absent
        self.observation_errors = observation_errors
        self.operation_error = operation_error


def _exists(path: Path) -> bool:
    try:
        path.lstat()
    except FileNotFoundError:
        return False
    return True


def _recycle_root(path: Path) -> Path:
    return Path(path.anchor) / "$Recycle.Bin"


def _original_path(content: bytes) -> str:
    version = struct.unpack_from("<Q", content)[0]
    if version not in {1, 2}:
        raise ValueError("unsupported_recycle_metadata_version")
    return content[28 if version == 2 else 24 :].decode("utf-16-le").split("\0")[0]


def _observe_after_recycle(path: Path, previous: set[str]) -> tuple[list[dict[str, str]], bool | None, list[str]]:
    errors = []
    try:
        candidates = [entry for entry in _recovery_entries(path) if entry["metadata"] not in previous]
    except (OSError, ValueError) as error:
        candidates = []
        errors.append(f"recovery_observation:{type(error).__name__}:{error}")
    try:
        absent = not _exists(path)
    except OSError as error:
        absent = None
        errors.append(f"source_observation:{type(error).__name__}:{error}")
    return candidates, absent, errors


def _recycle_guard():
    import winerror
    from win32com.server.exception import COMException
    from win32com.server.policy import DesignatedWrapPolicy
    from win32com.shell import shell

    class Guard(DesignatedWrapPolicy):
        _com_interfaces_: ClassVar[list[object]] = [shell.IID_IFileOperationProgressSink]
        _public_methods_: ClassVar[list[str]] = [
            "PreDeleteItem",
            "PostDeleteItem",
            "StartOperations",
            "FinishOperations",
            "UpdateProgress",
            "ResetTimer",
            "PauseTimer",
            "ResumeTimer",
        ]
        refused = False

        def __init__(self):
            self._wrap_(self)

        def PreDeleteItem(self, flags, item):
            # Raise a COM failure: returning an HRESULT integer from this Python
            # callback does not cancel the operation in pywin32.
            if not flags & 0x80:  # TSF_DELETE_RECYCLE_IF_POSSIBLE
                self.refused = True
                raise COMException(desc="permanent_deletion_refused", scode=winerror.E_ABORT)

        def _noop(self, *args):
            pass

        PostDeleteItem = StartOperations = FinishOperations = UpdateProgress = _noop
        ResetTimer = PauseTimer = ResumeTimer = _noop

    return Guard()


def _perform_recycle(path: Path) -> None:
    import pythoncom
    import pywintypes
    from win32com.shell import shell

    pythoncom.CoInitialize()
    try:
        operation = pythoncom.CoCreateInstance(
            shell.CLSID_FileOperation, None, pythoncom.CLSCTX_ALL, shell.IID_IFileOperation
        )
        # Request recycling and veto any permanent-delete operation before it runs.
        # Flags alone do not provide that veto for every Shell fallback.
        operation.SetOperationFlags(_RECYCLE_FLAGS)
        guard = _recycle_guard()
        operation.DeleteItem(
            shell.SHCreateItemFromParsingName(str(path), None, shell.IID_IShellItem),
            pythoncom.WrapObject(guard, shell.IID_IFileOperationProgressSink),
        )
        try:
            operation.PerformOperations()
        except pywintypes.com_error as error:
            if guard.refused:
                raise OSError(f"permanent_deletion_refused:{path}") from error
            raise OSError(f"recycle_failed:{path}:{error}") from error
        if guard.refused:
            raise OSError(f"permanent_deletion_refused:{path}")
        if operation.GetAnyOperationsAborted():
            raise OSError(f"recycle_aborted:{path}")
    finally:
        pythoncom.CoUninitialize()


def _recovery_entries(path: Path) -> list[dict[str, str]]:
    matches = []
    root = _recycle_root(path)
    if not _exists(root):
        return []
    try:
        assert_plain_chain(root)
    except OSError:
        return []
    for owner in root.iterdir():
        try:
            assert_plain_chain(owner)
            for metadata in owner.glob("$I*"):
                try:
                    content = read_plain_bytes(metadata, max_bytes=65536)
                    original = _original_path(content)
                    payload = metadata.with_name("$R" + metadata.name[2:])
                    if os.path.normcase(original) == os.path.normcase(str(path)) and _exists(payload):
                        matches.append({"metadata": str(metadata), "payload": str(payload)})
                except (OSError, ValueError, UnicodeError, struct.error):
                    continue
        except OSError:
            continue
    return matches


def verify_recovery_entry(
    path: Path,
    record: dict[str, Any],
    expected_identity: dict[str, Any],
    *,
    require_recorded_proof: bool = True,
) -> dict[str, Any]:
    """Verify current recovery. Never rewrite the earlier completed observation."""
    import hashlib

    expected = require_payload_identity(expected_identity)
    metadata = Path(record["metadata"])
    payload = Path(record["payload"])
    if (
        metadata.parent.parent != _recycle_root(path)
        or not metadata.name.startswith("$I")
        or payload != metadata.with_name("$R" + metadata.name[2:])
    ):
        raise OSError(f"recycle_recovery_boundary_mismatch:{path}")
    assert_plain_chain(metadata.parent)
    raw = read_plain_bytes(metadata, max_bytes=65536)
    if os.path.normcase(_original_path(raw)) != os.path.normcase(str(path)):
        raise OSError(f"recycle_metadata_original_changed:{path}")
    metadata_sha256 = hashlib.sha256(raw).hexdigest()
    metadata_identity = snapshot_payload(metadata)
    if metadata_identity["contentSha256"] != hashlib.sha256(b".\0" + raw + b"\0").hexdigest():
        raise OSError(f"recycle_metadata_changed_during_verification:{path}")
    if require_recorded_proof and (
        record.get("verification") != "identity-and-content-confirmed"
        or record.get("metadataSha256") != metadata_sha256
        or record.get("metadataIdentity") != metadata_identity
        or record.get("payloadIdentity") != expected
    ):
        raise OSError(f"recycle_recorded_proof_missing_or_changed:{path}")
    observed = snapshot_payload(payload)
    if observed != expected:
        raise OSError(f"recycle_payload_identity_or_content_changed:{path}:{payload}")
    if snapshot_payload(metadata) != metadata_identity:
        raise OSError(f"recycle_metadata_changed_during_verification:{path}")
    return {
        "metadata": str(metadata),
        "payload": str(payload),
        "metadataSha256": metadata_sha256,
        "metadataIdentity": metadata_identity,
        "payloadIdentity": observed,
        "verification": "identity-and-content-confirmed",
    }


def recycle_path(path: Path, *, expected_identity: dict[str, Any] | None = None) -> list[dict[str, Any]]:
    """Verify moved contents and identity, or fail without permanent deletion."""
    if os.name != "nt":
        raise OSError("recycle_requires_windows")
    if not path.is_absolute() or path == Path(path.anchor):
        raise ValueError("recycle_requires_absolute_nonroot_target")
    assert_plain_chain(path.parent)
    if path.is_symlink() or getattr(path.lstat(), "st_file_attributes", 0) & 0x400:
        raise ValueError("recycle_target_is_reparse_point")
    path = path.resolve(strict=True)
    before = snapshot_payload(path)
    if expected_identity is not None and before != require_payload_identity(expected_identity):
        raise OSError(f"recycle_source_identity_changed:{path}")
    previous = {entry["metadata"] for entry in _recovery_entries(path)}
    try:
        _perform_recycle(path)
    except OSError as error:
        candidates, absent, observation_errors = _observe_after_recycle(path, previous)
        raise RecycleUnconfirmedError(
            "recycle_operation_unconfirmed",
            path=path,
            recovery=candidates,
            expected=before,
            source_absent=absent,
            observation_errors=observation_errors,
            operation_error=f"{type(error).__name__}:{error}",
        ) from error
    candidates, absent, observation_errors = _observe_after_recycle(path, previous)
    if absent is not True or observation_errors:
        raise RecycleUnconfirmedError(
            "recycle_target_remains" if absent is False else "recycle_source_observation_unconfirmed",
            path=path,
            recovery=candidates,
            expected=before,
            source_absent=absent,
            observation_errors=observation_errors,
        )
    # Shell return can precede visibility of the metadata entry.
    for _ in range(30):
        entries, absent, observation_errors = _observe_after_recycle(path, previous)
        if observation_errors or absent is not True:
            raise RecycleUnconfirmedError(
                "recycle_recovery_observation_unconfirmed",
                path=path,
                recovery=entries,
                expected=before,
                source_absent=absent,
                observation_errors=observation_errors,
            )
        if entries:
            try:
                if len(entries) != 1:
                    raise OSError("recycle_recovery_not_unique")
                return [verify_recovery_entry(path, entries[0], before, require_recorded_proof=False)]
            except (OSError, ValueError, KeyError, struct.error) as error:
                raise RecycleUnconfirmedError(
                    "recycle_payload_verification_unconfirmed",
                    path=path,
                    recovery=entries,
                    expected=before,
                    source_absent=absent,
                    observation_errors=[f"payload_verification:{type(error).__name__}:{error}"],
                ) from error
        time.sleep(0.1)
    raise RecycleUnconfirmedError(
        "recycle_recovery_unverified",
        path=path,
        recovery=[],
        expected=before,
        source_absent=True,
        observation_errors=[],
    )


def recycle_directory(path: Path, *, expected_identity: dict[str, Any] | None = None) -> list[dict[str, Any]]:
    """Compatibility entry point for saved directory cleanup plans."""
    return recycle_path(path, expected_identity=expected_identity)


def recovery_entries(path: Path) -> list[dict[str, str]]:
    """Read current candidates, not proof that their payload still matches source."""
    return _recovery_entries(path)

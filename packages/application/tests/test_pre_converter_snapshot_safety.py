"""Adversarial contracts for Application-owned protective snapshots."""

from __future__ import annotations

import os
import stat
import threading
from pathlib import Path

import pytest

from docwen_core.cancellation import CancellationToken

pytestmark = pytest.mark.unit


def _successful_bridge(source: Path, calls: list[Path]):
    from docwen_core.office_bridge import BridgeResult

    def convert(input_path: str, output_path: str, **_kwargs: object) -> BridgeResult:
        protected = Path(input_path)
        calls.append(protected)
        Path(output_path).write_bytes(b"converted")
        assert source.read_bytes().startswith(b"SOURCE")
        return BridgeResult(True, output_path=output_path, backend="Fake Office")

    return convert


def test_copy_fault_never_publishes_partial_snapshot_or_overwrites_prior_file(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from docwen_application.preconversion import pre_converter

    source = tmp_path / "source.rtf"
    source.write_bytes(b"SOURCE-V1")
    staging = tmp_path / "staging"
    staging.mkdir()
    protected = staging / "input.rtf"
    protected.write_bytes(b"PRIOR-SNAPSHOT")
    bridge_calls: list[Path] = []

    def fail_after_write(_source, destination, _cancel) -> None:
        destination.write(b"torn")
        raise OSError("injected mid-copy failure")

    monkeypatch.setattr(pre_converter, "_copy_snapshot_stream", fail_after_write)
    monkeypatch.setattr(pre_converter, "convert_with_backend_priority", _successful_bridge(source, bridge_calls))

    outcome = pre_converter.pre_convert(str(source), "rtf", staging_dir=str(staging))

    assert isinstance(outcome, pre_converter.PreConversionFailure)
    assert outcome.error_type == "conversion_failed"
    assert outcome.diagnostic_code == "PRECONVERSION_INPUT_COPY_FAILED"
    assert protected.read_bytes() == b"PRIOR-SNAPSHOT"
    assert list(staging.glob(".input.rtf.docwen-snapshot-*.tmp")) == []
    assert source.read_bytes() == b"SOURCE-V1"
    assert bridge_calls == []


def test_success_atomically_replaces_prior_regular_snapshot_and_cleans_temp(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from docwen_application.preconversion import pre_converter

    source = tmp_path / "source.rtf"
    source.write_bytes(b"SOURCE-V1")
    os.utime(source, ns=(1_700_000_000_000_000_000, 1_700_000_001_000_000_000))
    staging = tmp_path / "staging"
    staging.mkdir()
    protected = staging / "input.rtf"
    protected.write_bytes(b"PRIOR-SNAPSHOT")
    bridge_calls: list[Path] = []
    monkeypatch.setattr(pre_converter, "convert_with_backend_priority", _successful_bridge(source, bridge_calls))

    outcome = pre_converter.pre_convert(str(source), "rtf", staging_dir=str(staging))

    assert isinstance(outcome, pre_converter.PreConversionResult)
    assert protected.read_bytes() == b"SOURCE-V1"
    assert protected.stat().st_mtime_ns == source.stat().st_mtime_ns
    assert stat.S_IMODE(protected.stat().st_mode) == stat.S_IMODE(source.stat().st_mode)
    assert list(staging.glob(".input.rtf.docwen-snapshot-*.tmp")) == []
    assert bridge_calls == [protected]


def test_publish_fault_preserves_prior_snapshot_and_cleans_private_temp(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from docwen_application.preconversion import pre_converter

    source = tmp_path / "source.rtf"
    source.write_bytes(b"SOURCE-V1")
    staging = tmp_path / "staging"
    staging.mkdir()
    protected = staging / "input.rtf"
    protected.write_bytes(b"PRIOR-SNAPSHOT")
    bridge_calls: list[Path] = []

    def fail_publish(_source_path, _destination_path) -> None:
        raise PermissionError("injected atomic publish failure")

    monkeypatch.setattr(pre_converter.os, "replace", fail_publish)
    monkeypatch.setattr(pre_converter, "convert_with_backend_priority", _successful_bridge(source, bridge_calls))

    outcome = pre_converter.pre_convert(str(source), "rtf", staging_dir=str(staging))

    assert isinstance(outcome, pre_converter.PreConversionFailure)
    assert outcome.diagnostic_code == "PRECONVERSION_INPUT_COPY_FAILED"
    assert protected.read_bytes() == b"PRIOR-SNAPSHOT"
    assert list(staging.glob(".input.rtf.docwen-snapshot-*.tmp")) == []
    assert bridge_calls == []


def test_source_cloud_placeholder_tag_is_not_misclassified_as_path_redirection(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from docwen_application.preconversion import pre_converter

    class ReparseStat:
        st_mode = stat.S_IFREG | 0o600
        st_file_attributes = 0x400
        st_reparse_tag = 0x9000001A

    monkeypatch.setattr(Path, "lstat", lambda _path: ReparseStat())

    assert pre_converter._is_name_surrogate(tmp_path / "cloud-placeholder.doc") is False


def test_source_change_during_snapshot_is_typed_and_never_bridged(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from docwen_application.preconversion import pre_converter

    source = tmp_path / "source.rtf"
    source.write_bytes(b"SOURCE-V1")
    staging = tmp_path / "staging"
    staging.mkdir()
    bridge_calls: list[Path] = []
    real_identity = pre_converter._snapshot_identity
    identity_calls = 0

    def changed_identity(source_stat):
        nonlocal identity_calls
        identity_calls += 1
        identity = real_identity(source_stat)
        if identity_calls == 2:
            return (*identity[:-1], identity[-1] + 1)
        return identity

    monkeypatch.setattr(pre_converter, "_snapshot_identity", changed_identity)
    monkeypatch.setattr(pre_converter, "convert_with_backend_priority", _successful_bridge(source, bridge_calls))

    outcome = pre_converter.pre_convert(str(source), "rtf", staging_dir=str(staging))

    assert isinstance(outcome, pre_converter.PreConversionFailure)
    assert outcome.error_type == "conversion_failed"
    assert outcome.diagnostic_code == "PRECONVERSION_SOURCE_CHANGED"
    assert not (staging / "input.rtf").exists()
    assert source.read_bytes() == b"SOURCE-V1"
    assert bridge_calls == []


def test_cancellation_during_snapshot_never_admits_partial_file(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from docwen_application.preconversion import pre_converter

    source = tmp_path / "source.rtf"
    source.write_bytes(b"SOURCE-" + b"X" * 64)
    staging = tmp_path / "staging"
    staging.mkdir()
    cancel_owner = CancellationToken()
    cancel = cancel_owner.view()
    bridge_calls: list[Path] = []
    real_copy = pre_converter._copy_snapshot_stream

    def copy_then_cancel(source_stream, destination_stream, token) -> None:
        destination_stream.write(source_stream.read(4))
        cancel_owner.cancel()
        real_copy(source_stream, destination_stream, token)

    monkeypatch.setattr(pre_converter, "_copy_snapshot_stream", copy_then_cancel)
    monkeypatch.setattr(pre_converter, "convert_with_backend_priority", _successful_bridge(source, bridge_calls))

    outcome = pre_converter.pre_convert(
        str(source),
        "rtf",
        staging_dir=str(staging),
        cancel=cancel,
    )

    assert isinstance(outcome, pre_converter.PreConversionFailure)
    assert outcome.cancelled is True
    assert not (staging / "input.rtf").exists()
    assert list(staging.glob(".input.rtf.docwen-snapshot-*.tmp")) == []
    assert bridge_calls == []


def test_cancellation_immediately_after_atomic_publication_skips_bridge_with_complete_snapshot(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from docwen_application.preconversion import pre_converter

    source = tmp_path / "source.rtf"
    source.write_bytes(b"SOURCE-V1")
    staging = tmp_path / "staging"
    staging.mkdir()
    cancel_owner = CancellationToken()
    cancel = cancel_owner.view()
    bridge_calls: list[Path] = []
    real_replace = os.replace

    def replace_then_cancel(source_path, destination_path) -> None:
        real_replace(source_path, destination_path)
        cancel_owner.cancel()

    monkeypatch.setattr(pre_converter.os, "replace", replace_then_cancel)
    monkeypatch.setattr(pre_converter, "convert_with_backend_priority", _successful_bridge(source, bridge_calls))

    outcome = pre_converter.pre_convert(
        str(source),
        "rtf",
        staging_dir=str(staging),
        cancel=cancel,
    )

    assert isinstance(outcome, pre_converter.PreConversionFailure)
    assert outcome.cancelled is True
    assert (staging / "input.rtf").read_bytes() == b"SOURCE-V1"
    assert bridge_calls == []


@pytest.mark.skipif(os.name != "nt", reason="NTFS junction contract")
def test_staging_junction_is_rejected_without_touching_target(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    from docwen_application.preconversion import pre_converter

    source = tmp_path / "source.rtf"
    source.write_bytes(b"SOURCE-V1")
    target = tmp_path / "unrelated"
    target.mkdir()
    sentinel = target / "sentinel.bin"
    sentinel.write_bytes(b"UNRELATED")
    junction = tmp_path / "staging-junction"
    _create_junction(junction, target)
    bridge_calls: list[Path] = []
    monkeypatch.setattr(pre_converter, "convert_with_backend_priority", _successful_bridge(source, bridge_calls))

    try:
        outcome = pre_converter.pre_convert(str(source), "rtf", staging_dir=str(junction))
    finally:
        os.rmdir(junction)

    assert isinstance(outcome, pre_converter.PreConversionFailure)
    assert outcome.diagnostic_code == "PRECONVERSION_UNSAFE_PATH"
    assert sentinel.read_bytes() == b"UNRELATED"
    assert sorted(path.name for path in target.iterdir()) == ["sentinel.bin"]
    assert bridge_calls == []


@pytest.mark.skipif(os.name != "nt", reason="NTFS junction contract")
def test_source_junction_is_rejected_without_reading_or_modifying_target(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from docwen_application.preconversion import pre_converter

    source_target = tmp_path / "source-target"
    source_target.mkdir()
    sentinel = source_target / "sentinel.bin"
    sentinel.write_bytes(b"UNRELATED")
    source_junction = tmp_path / "source.rtf"
    _create_junction(source_junction, source_target)
    staging = tmp_path / "staging"
    staging.mkdir()
    bridge_calls: list[Path] = []
    monkeypatch.setattr(
        pre_converter,
        "convert_with_backend_priority",
        _successful_bridge(sentinel, bridge_calls),
    )

    try:
        outcome = pre_converter.pre_convert(
            str(source_junction),
            "rtf",
            staging_dir=str(staging),
        )
    finally:
        os.rmdir(source_junction)

    assert isinstance(outcome, pre_converter.PreConversionFailure)
    assert outcome.diagnostic_code == "PRECONVERSION_UNSAFE_PATH"
    assert sentinel.read_bytes() == b"UNRELATED"
    assert list(staging.iterdir()) == []
    assert bridge_calls == []


@pytest.mark.skipif(os.name != "nt", reason="NTFS junction contract")
def test_existing_protected_junction_is_rejected_without_following_target(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from docwen_application.preconversion import pre_converter

    source = tmp_path / "source.rtf"
    source.write_bytes(b"SOURCE-V1")
    staging = tmp_path / "staging"
    staging.mkdir()
    target = tmp_path / "unrelated"
    target.mkdir()
    sentinel = target / "sentinel.bin"
    sentinel.write_bytes(b"UNRELATED")
    protected_junction = staging / "input.rtf"
    _create_junction(protected_junction, target)
    bridge_calls: list[Path] = []
    monkeypatch.setattr(pre_converter, "convert_with_backend_priority", _successful_bridge(source, bridge_calls))

    try:
        outcome = pre_converter.pre_convert(str(source), "rtf", staging_dir=str(staging))
    finally:
        os.rmdir(protected_junction)

    assert isinstance(outcome, pre_converter.PreConversionFailure)
    assert outcome.diagnostic_code == "PRECONVERSION_UNSAFE_PATH"
    assert sentinel.read_bytes() == b"UNRELATED"
    assert sorted(path.name for path in target.iterdir()) == ["sentinel.bin"]
    assert bridge_calls == []


@pytest.mark.skipif(os.name != "nt", reason="Windows share-mode contract")
def test_windows_snapshot_handle_blocks_writer_until_publication(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from docwen_application.preconversion import pre_converter

    source = tmp_path / "source.rtf"
    source.write_bytes(b"SOURCE-V1")
    staging = tmp_path / "staging"
    staging.mkdir()
    copy_started = threading.Event()
    allow_copy = threading.Event()
    bridge_calls: list[Path] = []
    real_copy = pre_converter._copy_snapshot_stream

    def paused_copy(source_stream, destination_stream, cancel) -> None:
        copy_started.set()
        assert allow_copy.wait(timeout=5)
        real_copy(source_stream, destination_stream, cancel)

    monkeypatch.setattr(pre_converter, "_copy_snapshot_stream", paused_copy)
    monkeypatch.setattr(pre_converter, "convert_with_backend_priority", _successful_bridge(source, bridge_calls))
    outcomes: list[object] = []
    worker = threading.Thread(
        target=lambda: outcomes.append(pre_converter.pre_convert(str(source), "rtf", staging_dir=str(staging)))
    )
    worker.start()
    assert copy_started.wait(timeout=5)

    try:
        with pytest.raises(PermissionError):
            source.write_bytes(b"SOURCE-V2")
        replacement = tmp_path / "replacement.rtf"
        replacement.write_bytes(b"SOURCE-V3")
        with pytest.raises(PermissionError):
            os.replace(replacement, source)
        assert replacement.read_bytes() == b"SOURCE-V3"
    finally:
        allow_copy.set()
        worker.join(timeout=10)

    assert not worker.is_alive()
    assert isinstance(outcomes[0], pre_converter.PreConversionResult)
    assert (staging / "input.rtf").read_bytes() == b"SOURCE-V1"
    assert source.read_bytes() == b"SOURCE-V1"
    assert bridge_calls == [staging / "input.rtf"]


@pytest.mark.skipif(os.name != "nt", reason="Windows share-mode contract")
def test_windows_exclusive_source_lock_is_a_typed_failure_without_stale_bridge(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from docwen_application.preconversion import pre_converter

    source = tmp_path / "source.rtf"
    source.write_bytes(b"SOURCE-V1")
    staging = tmp_path / "staging"
    staging.mkdir()
    (staging / "input.rtf").write_bytes(b"PRIOR-SNAPSHOT")
    bridge_calls: list[Path] = []
    monkeypatch.setattr(pre_converter, "convert_with_backend_priority", _successful_bridge(source, bridge_calls))

    handle = _open_exclusive_windows_handle(source)
    try:
        outcome = pre_converter.pre_convert(str(source), "rtf", staging_dir=str(staging))
    finally:
        _close_windows_handle(handle)

    assert isinstance(outcome, pre_converter.PreConversionFailure)
    assert outcome.error_type == "conversion_failed"
    assert outcome.diagnostic_code == "PRECONVERSION_INPUT_COPY_FAILED"
    assert source.read_bytes() == b"SOURCE-V1"
    assert (staging / "input.rtf").read_bytes() == b"PRIOR-SNAPSHOT"
    assert bridge_calls == []


def test_source_handle_is_released_before_bridge_and_backend_mutation_is_contained(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from docwen_application.preconversion import pre_converter
    from docwen_core.office_bridge import BridgeResult

    source = tmp_path / "source.rtf"
    source.write_bytes(b"SOURCE-V1")
    staging = tmp_path / "staging"
    staging.mkdir()

    def mutating_bridge(input_path: str, output_path: str, **_kwargs: object) -> BridgeResult:
        source.write_bytes(b"SOURCE-V2")
        protected = Path(input_path)
        assert protected.read_bytes() == b"SOURCE-V1"
        protected.write_bytes(b"BACKEND-MUTATION")
        Path(output_path).write_bytes(b"converted")
        return BridgeResult(True, output_path=output_path, backend="Fake Office")

    monkeypatch.setattr(pre_converter, "convert_with_backend_priority", mutating_bridge)

    outcome = pre_converter.pre_convert(str(source), "rtf", staging_dir=str(staging))

    assert isinstance(outcome, pre_converter.PreConversionResult)
    assert source.read_bytes() == b"SOURCE-V2"
    assert (staging / "input.rtf").read_bytes() == b"BACKEND-MUTATION"


def _create_junction(link: Path, target: Path) -> None:
    import subprocess

    completed = subprocess.run(
        ["cmd.exe", "/d", "/c", "mklink", "/J", str(link), str(target)],
        check=False,
        capture_output=True,
        text=True,
    )
    if completed.returncode != 0:
        pytest.fail(f"cannot create NTFS junction: {completed.stdout} {completed.stderr}")


def _open_exclusive_windows_handle(path: Path) -> int:
    import ctypes
    from ctypes import wintypes

    ctypes_api = vars(ctypes)
    win_dll = ctypes_api["WinDLL"]
    win_error = ctypes_api["WinError"]
    get_last_error = ctypes_api["get_last_error"]
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
    handle = create_file(str(path), 0x80000000, 0, None, 3, 0x80, None)
    if handle == ctypes.c_void_p(-1).value:
        raise win_error(get_last_error())
    return int(handle)


def _close_windows_handle(handle: int) -> None:
    import ctypes
    from ctypes import wintypes

    ctypes_api = vars(ctypes)
    win_dll = ctypes_api["WinDLL"]
    win_error = ctypes_api["WinError"]
    get_last_error = ctypes_api["get_last_error"]
    close_handle = win_dll("kernel32", use_last_error=True).CloseHandle
    close_handle.argtypes = [wintypes.HANDLE]
    close_handle.restype = wintypes.BOOL
    if not close_handle(handle):
        raise win_error(get_last_error())

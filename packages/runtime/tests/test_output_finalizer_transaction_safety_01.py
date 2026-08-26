"""Focused tests split from test_output_finalizer_transaction_safety.py."""

from __future__ import annotations

from ._output_finalizer_transaction_safety_support import (
    Any,
    CancellationRequested,
    CancellationToken,
    OutputFinalizer,
    OutputPolicy,
    Path,
    _artifact,
    _finalize_process,
    _hold_process_lock,
    errno,
    multiprocessing,
    os,
    pytest,
    sys,
    threading,
    time,
)

pytestmark = pytest.mark.integration


@pytest.mark.parametrize(
    "error",
    [
        IsADirectoryError("Existing output target is not a file: report.md"),
        FileNotFoundError("Staging artifact is not a file: staging.md"),
    ],
)
def test_public_exception_text_preserves_meaningful_manual_oserror(error: OSError) -> None:
    assert OutputFinalizer._public_exception_text(error) == str(error)


def test_public_exception_text_strips_internal_prefix_from_structured_oserror(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(sys, "platform", "win32")
    error = OSError(
        errno.EACCES,
        r"injected \\?\C:\private-detail",
        r"\\?\C:\private-source",
        None,
        r"\\?\UNC\server\share\private-destination",
    )

    message = OutputFinalizer._public_exception_text(error)

    assert "\\\\?\\" not in message
    assert "private-detail" in message
    assert "private-source" in message
    assert "private-destination" in message


def test_pre_cancelled_finalization_does_not_create_destination(tmp_path: Path) -> None:
    staging = tmp_path / "staging.md"
    staging.write_bytes(b"complete payload")
    output = tmp_path / "output"
    token = CancellationToken()
    token.cancel("red-pre-lock")

    with pytest.raises(CancellationRequested):
        OutputFinalizer().finalize(
            task_id="pre-cancelled",
            artifacts=[_artifact(staging)],
            policy=OutputPolicy(output_dir=str(output)),
            cancellation=token.view(),
        )

    assert not output.exists()


def test_overwrite_copy_fault_preserves_existing_destination(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    staging = tmp_path / "staging.md"
    staging.write_bytes(b"new complete payload")
    output = tmp_path / "output"
    output.mkdir()
    destination = output / "report.md"
    destination.write_bytes(b"old authoritative payload")

    def partial_then_fail(source: str, target: str, cancellation: object) -> None:
        del source
        del cancellation
        Path(target).write_bytes(b"torn")
        raise OSError("injected copy failure")

    monkeypatch.setattr(OutputFinalizer, "_copy_to_temp", staticmethod(partial_then_fail))

    result = OutputFinalizer().finalize(
        task_id="overwrite-copy-fault",
        artifacts=[_artifact(staging)],
        policy=OutputPolicy(output_dir=str(output), overwrite_mode="overwrite"),
    )

    assert result.success is False
    assert destination.read_bytes() == b"old authoritative payload"
    assert sorted(path.name for path in output.iterdir()) == ["report.md"]


def test_resolved_descendant_escape_is_rejected_without_mutation(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    staging = tmp_path / "staging.md"
    staging.write_bytes(b"must stay contained")
    output = tmp_path / "output"
    outside = tmp_path / "outside"
    output.mkdir()
    outside.mkdir()
    real_realpath = os.path.realpath

    def escaped_realpath(path: os.PathLike[str] | str) -> str:
        absolute = OutputFinalizer._logical_io_spelling(os.path.abspath(path))
        escaped_prefix = os.path.abspath(output / "escaped")
        if absolute == escaped_prefix or absolute.startswith(escaped_prefix + os.sep):
            relative = os.path.relpath(absolute, escaped_prefix)
            return os.path.abspath(outside / relative)
        return real_realpath(absolute)

    monkeypatch.setattr(os.path, "realpath", escaped_realpath)

    result = OutputFinalizer().finalize(
        task_id="resolved-escape",
        artifacts=[_artifact(staging, "escaped/report.md")],
        policy=OutputPolicy(output_dir=str(output)),
    )

    assert result.success is False
    assert result.error is not None
    assert result.error.diagnostic_code == "FINALIZER_FAILED"
    assert not (output / "escaped" / "report.md").exists()
    assert not (outside / "report.md").exists()


def test_cancellation_while_waiting_for_process_lock_has_no_output(tmp_path: Path) -> None:
    staging = tmp_path / "staging.md"
    staging.write_bytes(b"blocked payload")
    output = tmp_path / "output"
    context = multiprocessing.get_context("spawn")
    ready = context.Event()
    release = context.Event()
    holder = context.Process(target=_hold_process_lock, args=(str(output), ready, release))
    holder.start()
    assert ready.wait(10.0)
    token = CancellationToken()
    raised: list[BaseException] = []

    def finalize() -> None:
        try:
            OutputFinalizer().finalize(
                task_id="cancel-process-lock-wait",
                artifacts=[_artifact(staging)],
                policy=OutputPolicy(output_dir=str(output)),
                cancellation=token.view(),
            )
        except BaseException as exc:
            raised.append(exc)

    worker = threading.Thread(target=finalize)
    worker.start()
    time.sleep(0.15)
    token.cancel("cancel while process lock is held")
    worker.join(3.0)
    release.set()
    holder.join(10.0)

    assert not worker.is_alive()
    assert holder.exitcode == 0
    assert len(raised) == 1
    assert isinstance(raised[0], CancellationRequested)
    assert not output.exists()


def test_cancellation_during_temp_copy_removes_private_temp(tmp_path: Path, monkeypatch) -> None:
    staging = tmp_path / "staging.md"
    staging.write_bytes(b"complete payload")
    output = tmp_path / "output"
    token = CancellationToken()

    def cancel_during_copy(source: str, temp_path: str, cancellation: Any) -> None:
        del source
        Path(temp_path).write_bytes(b"private partial")
        token.cancel("cancel during copy")
        cancellation.check()

    monkeypatch.setattr(OutputFinalizer, "_copy_to_temp", staticmethod(cancel_during_copy))

    with pytest.raises(CancellationRequested):
        OutputFinalizer().finalize(
            task_id="cancel-copy",
            artifacts=[_artifact(staging)],
            policy=OutputPolicy(output_dir=str(output)),
            cancellation=token.view(),
        )

    assert not (output / "report.md").exists()
    assert list(output.iterdir()) == []


def test_cancellation_immediately_before_commit_publishes_nothing(tmp_path: Path, monkeypatch) -> None:
    staging = tmp_path / "staging.md"
    staging.write_bytes(b"prepared payload")
    output = tmp_path / "output"
    token = CancellationToken()
    real_prepare = OutputFinalizer._prepare_artifact

    def prepare_then_cancel(*args: Any, **kwargs: Any) -> Any:
        prepared = real_prepare(*args, **kwargs)
        token.cancel("cancel at precommit boundary")
        return prepared

    monkeypatch.setattr(OutputFinalizer, "_prepare_artifact", staticmethod(prepare_then_cancel))

    with pytest.raises(CancellationRequested):
        OutputFinalizer().finalize(
            task_id="cancel-precommit",
            artifacts=[_artifact(staging)],
            policy=OutputPolicy(output_dir=str(output)),
            cancellation=token.view(),
        )

    assert list(output.iterdir()) == []


def test_cancel_after_commit_does_not_rewrite_committed_success(tmp_path: Path, monkeypatch) -> None:
    staging = tmp_path / "staging.md"
    staging.write_bytes(b"committed payload")
    output = tmp_path / "output"
    token = CancellationToken()
    real_commit = OutputFinalizer._commit_prepared

    def commit_then_cancel(*args: Any, **kwargs: Any) -> Any:
        committed = real_commit(*args, **kwargs)
        token.cancel("late after commit")
        return committed

    monkeypatch.setattr(OutputFinalizer, "_commit_prepared", staticmethod(commit_then_cancel))

    result = OutputFinalizer().finalize(
        task_id="late-cancel",
        artifacts=[_artifact(staging)],
        policy=OutputPolicy(output_dir=str(output)),
        cancellation=token.view(),
    )

    assert token.is_cancelled is True
    assert result.success is True
    assert (output / "report.md").read_bytes() == b"committed payload"


def test_new_destination_is_absent_until_complete_temp_is_committed(tmp_path: Path, monkeypatch) -> None:
    staging = tmp_path / "staging.md"
    staging.write_bytes(b"complete payload")
    output = tmp_path / "output"
    real_copy = OutputFinalizer._copy_to_temp

    def observed_copy(source: str, temp_path: str, cancellation: Any) -> None:
        assert not (output / "report.md").exists()
        real_copy(source, temp_path, cancellation)
        assert Path(temp_path).read_bytes() == b"complete payload"
        assert not (output / "report.md").exists()

    monkeypatch.setattr(OutputFinalizer, "_copy_to_temp", staticmethod(observed_copy))

    result = OutputFinalizer().finalize(
        task_id="atomic-new",
        artifacts=[_artifact(staging)],
        policy=OutputPolicy(output_dir=str(output)),
    )

    assert result.success is True
    assert (output / "report.md").read_bytes() == b"complete payload"
    assert [path.name for path in output.iterdir()] == ["report.md"]


def test_multi_artifact_partial_keeps_only_complete_commits(tmp_path: Path) -> None:
    first = tmp_path / "first.md"
    first.write_bytes(b"first complete")
    missing = tmp_path / "missing.md"
    output = tmp_path / "output"
    artifacts = [_artifact(first, "first.md"), _artifact(missing, "missing.md")]
    artifacts[0].artifact_id = "first"
    artifacts[1].artifact_id = "missing"

    result = OutputFinalizer().finalize(
        task_id="partial-complete-only",
        artifacts=artifacts,
        policy=OutputPolicy(output_dir=str(output)),
    )

    assert result.success is False
    assert result.error is not None
    assert result.error.diagnostic_code == "FINALIZER_PARTIAL"
    assert (output / "first.md").read_bytes() == b"first complete"
    assert not (output / "missing.md").exists()
    assert [path.name for path in output.iterdir()] == ["first.md"]


def test_next_locked_run_removes_only_aged_reserved_stale_temp(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    staging = tmp_path / "staging.md"
    staging.write_bytes(b"next run")
    output = tmp_path / "output"
    output.mkdir()
    stale = output / ".__docwen-finalizer-interrupted"
    user_file = output / ".user-private"
    stale.write_bytes(b"stale private partial")
    user_file.write_bytes(b"keep me")
    observed_now = time.time()
    monkeypatch.setattr(time, "time", lambda: observed_now + (48 * 60 * 60))

    result = OutputFinalizer().finalize(
        task_id="stale-recovery",
        artifacts=[_artifact(staging)],
        policy=OutputPolicy(output_dir=str(output)),
    )

    assert result.success is True
    assert not stale.exists()
    assert user_file.read_bytes() == b"keep me"
    assert (output / "report.md").read_bytes() == b"next run"


def test_stale_cleanup_does_not_unlink_a_replaced_fresh_temp(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    parent = tmp_path / "output"
    parent.mkdir()
    target = parent / ".__docwen-finalizer-race"
    target.write_bytes(b"aged predecessor")
    observed_now = time.time()
    monkeypatch.setattr(time, "time", lambda: observed_now + (48 * 60 * 60))

    class SwappingEntry:
        name = target.name
        path = str(target)

        @staticmethod
        def is_file(*, follow_symlinks: bool) -> bool:
            assert follow_symlinks is False
            return True

        @staticmethod
        def stat(*, follow_symlinks: bool):
            assert follow_symlinks is False
            inspected = target.stat(follow_symlinks=False)
            target.unlink()
            target.write_bytes(b"fresh replacement")
            return inspected

    class Entries:
        def __enter__(self):
            return iter([SwappingEntry()])

        def __exit__(self, *_args: Any) -> None:
            return None

    monkeypatch.setattr(os, "scandir", lambda _path: Entries())

    OutputFinalizer._cleanup_stale_temps(str(parent))

    assert target.read_bytes() == b"fresh replacement"


def test_overlapping_output_roots_do_not_delete_a_live_private_temp(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    first_source = tmp_path / "first.md"
    second_source = tmp_path / "second.md"
    first_source.write_bytes(b"first complete")
    second_source.write_bytes(b"second complete")
    root = tmp_path / "output"
    first_ready = threading.Event()
    release_first = threading.Event()
    second_lock_attempted = threading.Event()
    first_result: list[Any] = []
    second_result: list[Any] = []
    real_copy = OutputFinalizer._copy_to_temp
    real_acquire = OutputFinalizer._acquire_thread_lock
    observed_now = time.time()
    monkeypatch.setattr(time, "time", lambda: observed_now + (48 * 60 * 60))

    def block_first_copy(source: str, target: str, cancellation: Any) -> None:
        real_copy(source, target, cancellation)
        if source == str(first_source):
            first_ready.set()
            if not release_first.wait(10.0):
                raise TimeoutError("second finalizer did not complete")

    monkeypatch.setattr(OutputFinalizer, "_copy_to_temp", staticmethod(block_first_copy))

    def observe_second_lock(output_lock: Any, cancellation: Any) -> None:
        if threading.current_thread().name == "overlap-second":
            second_lock_attempted.set()
        real_acquire(output_lock, cancellation)

    monkeypatch.setattr(OutputFinalizer, "_acquire_thread_lock", staticmethod(observe_second_lock))

    def run_first() -> None:
        first_result.append(
            OutputFinalizer().finalize(
                task_id="overlap-root",
                artifacts=[_artifact(first_source, os.path.join("nested", "first.md"))],
                policy=OutputPolicy(output_dir=str(root)),
            )
        )

    thread = threading.Thread(target=run_first)
    thread.start()
    assert first_ready.wait(10.0)

    def run_second() -> None:
        second_result.append(
            OutputFinalizer().finalize(
                task_id="overlap-nested",
                artifacts=[_artifact(second_source, "second.md")],
                policy=OutputPolicy(output_dir=str(root / "nested")),
            )
        )

    second_thread = threading.Thread(target=run_second, name="overlap-second")
    second_thread.start()
    try:
        assert second_lock_attempted.wait(10.0)
    finally:
        release_first.set()
        thread.join(10.0)
        second_thread.join(10.0)

    assert not thread.is_alive()
    assert not second_thread.is_alive()
    assert len(first_result) == 1
    assert len(second_result) == 1
    assert first_result[0].success is True, first_result[0].diagnostics
    assert second_result[0].success is True, second_result[0].diagnostics
    assert (root / "nested" / "first.md").read_bytes() == b"first complete"
    assert (root / "nested" / "second.md").read_bytes() == b"second complete"
    assert not any(path.name.startswith(".__docwen-finalizer-") for path in (root / "nested").iterdir())


def test_finalization_lock_set_includes_each_concrete_nested_parent(tmp_path: Path) -> None:
    output = tmp_path / "output"
    staging = tmp_path / "staging.md"
    staging.write_bytes(b"complete")
    artifact = _artifact(staging, os.path.join("nested", "report.md"))

    paths = OutputFinalizer._finalization_lock_paths(str(output), [artifact])
    keys = {OutputFinalizer._lock_key(path) for path in paths}

    assert keys == {
        OutputFinalizer._lock_key(str(output)),
        OutputFinalizer._lock_key(str(output / "nested")),
    }


def test_no_clobber_publish_uses_the_platform_atomic_primitive(
    tmp_path: Path,
    monkeypatch,
) -> None:
    staging = tmp_path / "staging.md"
    staging.write_bytes(b"portable atomic publish")
    output = tmp_path / "output"

    def unexpected_primitive(*_args: Any, **_kwargs: Any) -> None:
        raise AssertionError("wrong platform no-clobber primitive")

    if os.name == "nt":
        monkeypatch.setattr(os, "link", unexpected_primitive)
    else:
        monkeypatch.setattr(os, "rename", unexpected_primitive)

    result = OutputFinalizer().finalize(
        task_id="platform-no-clobber",
        artifacts=[_artifact(staging)],
        policy=OutputPolicy(output_dir=str(output)),
    )

    assert result.success is True
    assert (output / "report.md").read_bytes() == b"portable atomic publish"
    assert [path.name for path in output.iterdir()] == ["report.md"]


def test_two_real_processes_preserve_same_name_payloads(tmp_path: Path) -> None:
    sources = [tmp_path / "source-0.md", tmp_path / "source-1.md"]
    sources[0].write_bytes(b"process zero")
    sources[1].write_bytes(b"process one")
    output = tmp_path / "output"
    context = multiprocessing.get_context("spawn")
    start = context.Event()
    results = context.Queue()
    processes = [
        context.Process(
            target=_finalize_process,
            args=(str(source), str(output), f"process-{index}", start, results),
        )
        for index, source in enumerate(sources)
    ]
    for process in processes:
        process.start()
    start.set()
    reported = [results.get(timeout=15.0) for _ in processes]
    for process in processes:
        process.join(15.0)

    assert all(process.exitcode == 0 for process in processes)
    assert all(success for success, _name, _error in reported)
    assert sorted(name for _success, name, _error in reported) == ["report.md", "report_001.md"]
    assert {path.read_bytes() for path in output.iterdir()} == {b"process zero", b"process one"}


def test_external_rename_collision_at_commit_retries_suffix(tmp_path: Path, monkeypatch) -> None:
    staging = tmp_path / "staging.md"
    staging.write_bytes(b"docwen payload")
    output = tmp_path / "output"
    output.mkdir()
    (output / "report.md").write_bytes(b"original payload")
    real_commit = OutputFinalizer._commit_prepared

    def collide_then_commit(item: Any, output_dir: str, overwrite_mode: str) -> Any:
        Path(item.destination).write_bytes(b"external payload")
        return real_commit(item, output_dir, overwrite_mode)

    monkeypatch.setattr(OutputFinalizer, "_commit_prepared", staticmethod(collide_then_commit))

    result = OutputFinalizer().finalize(
        task_id="external-collision",
        artifacts=[_artifact(staging)],
        policy=OutputPolicy(output_dir=str(output), overwrite_mode="rename"),
    )

    assert result.success is True
    assert (output / "report.md").read_bytes() == b"original payload"
    assert (output / "report_001.md").read_bytes() == b"external payload"
    assert (output / "report_002.md").read_bytes() == b"docwen payload"

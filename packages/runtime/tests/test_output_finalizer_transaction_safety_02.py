"""Focused tests split from test_output_finalizer_transaction_safety.py."""

from __future__ import annotations

from ._output_finalizer_transaction_safety_support import (
    Any,
    ArtifactManifest,
    OutputFinalizer,
    OutputPolicy,
    Path,
    _artifact,
    _io_bytes,
    _io_names,
    _output_dir_with_total_length,
    _suggested_name_for_total_length,
    _utf16_units,
    errno,
    os,
    pytest,
    time,
)

pytestmark = pytest.mark.integration


@pytest.mark.skipif(os.name != "nt", reason="Win32 extended-path boundary")
@pytest.mark.parametrize("total_units", [248, 259])
def test_windows_public_length_output_directory_is_created_with_internal_namespace(
    tmp_path: Path,
    total_units: int,
) -> None:
    staging = tmp_path / "staging.md"
    staging.write_bytes(b"directory-boundary")
    output = _output_dir_with_total_length(tmp_path, total_units)

    result = OutputFinalizer().finalize(
        task_id=f"windows-output-dir-{total_units}",
        artifacts=[_artifact(staging)],
        policy=OutputPolicy(output_dir=str(output)),
    )

    assert result.success is True, result.diagnostics
    logical_path = result.artifacts[0].staging_path
    assert _utf16_units(os.path.abspath(output)) == total_units
    assert not logical_path.startswith("\\\\?\\")
    assert _io_bytes(logical_path) == b"directory-boundary"


@pytest.mark.skipif(os.name != "nt", reason="Win32 extended-path boundary")
@pytest.mark.parametrize("target_length", [259, 260])
def test_windows_publish_crosses_legacy_max_path_boundary(
    tmp_path: Path,
    target_length: int,
) -> None:
    staging = tmp_path / f"staging-{target_length}.md"
    payload = f"complete-{target_length}".encode()
    staging.write_bytes(payload)
    output = tmp_path / "output"
    suggested_name = _suggested_name_for_total_length(output, target_length)

    result = OutputFinalizer().finalize(
        task_id=f"windows-boundary-{target_length}",
        artifacts=[_artifact(staging, suggested_name)],
        policy=OutputPolicy(output_dir=str(output), overwrite_mode="error"),
    )

    assert result.success is True, result.diagnostics
    logical_path = result.artifacts[0].staging_path
    assert len(os.path.abspath(logical_path)) == target_length
    assert not logical_path.startswith("\\\\?\\")
    assert _io_bytes(logical_path) == payload
    assert result.metrics.output_bytes == len(payload)
    assert not any(name.startswith(".__docwen-finalizer-") for name in _io_names(Path(logical_path).parent))


@pytest.mark.skipif(os.name != "nt", reason="Win32 extended-path boundary")
def test_windows_publish_counts_non_bmp_utf16_units(tmp_path: Path) -> None:
    staging = tmp_path / "emoji-staging.md"
    staging.write_bytes(b"emoji-path-payload")
    output = tmp_path / "emoji-output"
    remaining_units = 260 - _utf16_units(os.path.abspath(output)) - 1
    if remaining_units < 5:
        pytest.skip("pytest temp root is already too long for the UTF-16 boundary")
    emoji_count = min(12, (remaining_units - 4) // 2)
    ascii_count = remaining_units - (emoji_count * 2) - len(".md")
    emoji = "\U0001f4c4"
    suggested_name = f"{emoji * emoji_count}{'r' * ascii_count}.md"
    assert _utf16_units(suggested_name) == remaining_units

    result = OutputFinalizer().finalize(
        task_id="windows-non-bmp-boundary",
        artifacts=[_artifact(staging, suggested_name)],
        policy=OutputPolicy(output_dir=str(output), overwrite_mode="error"),
    )

    assert result.success is True, result.diagnostics
    logical_path = result.artifacts[0].staging_path
    assert len(os.path.abspath(logical_path)) < 260
    assert _utf16_units(os.path.abspath(logical_path)) == 260
    assert not logical_path.startswith("\\\\?\\")
    assert _io_bytes(logical_path) == b"emoji-path-payload"


@pytest.mark.skipif(os.name != "nt", reason="Win32 namespace policy")
@pytest.mark.parametrize(
    "path",
    [
        r"\\.\PhysicalDrive0",
        r"\\?\GLOBALROOT\Device\HarddiskVolumeShadowCopy1\Windows",
        r"\\?\pipe\docwen",
        r"\\?\Volume{00000000-0000-0000-0000-000000000000}\report.md",
        r"\\?\C:relative.md",
        r"\\?\UNC\server",
        r"\\?\UNC\\share\report.md",
    ],
)
def test_windows_io_path_rejects_device_and_nonfilesystem_namespaces(path: str) -> None:
    with pytest.raises(ValueError, match="device namespace"):
        OutputFinalizer._io_path(path)


@pytest.mark.skipif(os.name != "nt", reason="Win32 extended-path boundary")
def test_windows_io_path_handles_local_unc_and_extended_namespaces() -> None:
    local_259 = "C:\\" + "a" * 256
    local_260 = f"{local_259}a"
    unc_prefix = "\\\\server\\share\\"
    unc_260 = unc_prefix + "u" * (260 - _utf16_units(unc_prefix))
    already_extended = r"\\?\C:\already-extended\report.md"
    already_extended_unc = r"\\?\UNC\server\share\report.md"

    assert str(OutputFinalizer._io_path(local_259)) == local_259
    assert str(OutputFinalizer._io_path(local_260)) == f"\\\\?\\{local_260}"
    assert str(OutputFinalizer._io_path(unc_260)) == f"\\\\?\\UNC\\{unc_260[2:]}"
    assert str(OutputFinalizer._io_path(already_extended)) == already_extended
    assert str(OutputFinalizer._io_path(already_extended_unc)) == already_extended_unc


@pytest.mark.skipif(os.name != "nt", reason="Win32 extended-path boundary")
def test_windows_forced_io_parent_reserves_space_for_private_child() -> None:
    from docwen_runtime.path_io import filesystem_path

    local_259 = "C:\\" + "a" * 256

    assert str(filesystem_path(local_259)) == local_259
    assert str(filesystem_path(local_259, force_extended=True)) == f"\\\\?\\{local_259}"


@pytest.mark.skipif(os.name != "nt", reason="Win32 extended-path boundary")
@pytest.mark.parametrize("overwrite_mode", ["rename", "overwrite", "skip"])
def test_windows_long_destination_preserves_collision_contracts(
    tmp_path: Path,
    overwrite_mode: str,
) -> None:
    output = tmp_path / "output"
    suggested_name = _suggested_name_for_total_length(output, 260)
    original = tmp_path / "original.md"
    replacement = tmp_path / "replacement.md"
    original.write_bytes(b"original")
    replacement.write_bytes(b"replacement")
    finalizer = OutputFinalizer()

    first = finalizer.finalize(
        task_id="windows-long-collision-original",
        artifacts=[_artifact(original, suggested_name)],
        policy=OutputPolicy(output_dir=str(output), overwrite_mode="error"),
    )
    second = finalizer.finalize(
        task_id=f"windows-long-collision-{overwrite_mode}",
        artifacts=[_artifact(replacement, suggested_name)],
        policy=OutputPolicy(output_dir=str(output), overwrite_mode=overwrite_mode),
    )

    assert first.success is True, first.diagnostics
    assert second.success is True, second.diagnostics
    original_logical = first.artifacts[0].staging_path
    replacement_logical = second.artifacts[0].staging_path
    assert not original_logical.startswith("\\\\?\\")
    assert not replacement_logical.startswith("\\\\?\\")
    if overwrite_mode == "rename":
        assert replacement_logical != original_logical
        assert Path(replacement_logical).stem.endswith("_001")
        assert _io_bytes(original_logical) == b"original"
        assert _io_bytes(replacement_logical) == b"replacement"
    elif overwrite_mode == "overwrite":
        assert replacement_logical == original_logical
        assert _io_bytes(original_logical) == b"replacement"
    else:
        assert replacement_logical == original_logical
        assert second.artifacts[0].metadata["skipped"] is True
        assert _io_bytes(original_logical) == b"original"
    assert not any(name.startswith(".__docwen-finalizer-") for name in _io_names(Path(original_logical).parent))


@pytest.mark.skipif(os.name != "nt", reason="Win32 extended-path boundary")
def test_windows_long_external_commit_collision_retries_second_suffix(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    output = tmp_path / "output"
    suggested_name = _suggested_name_for_total_length(output, 260)
    staging = tmp_path / "staging.md"
    staging.write_bytes(b"docwen payload")
    real_commit = OutputFinalizer._commit_prepared

    def collide_then_commit(item: Any, output_dir: str, overwrite_mode: str) -> Any:
        OutputFinalizer._io_path(item.destination).write_bytes(b"external base")
        base, extension = os.path.splitext(item.destination)
        OutputFinalizer._io_path(f"{base}_001{extension}").write_bytes(b"external suffix")
        return real_commit(item, output_dir, overwrite_mode)

    monkeypatch.setattr(OutputFinalizer, "_commit_prepared", staticmethod(collide_then_commit))
    result = OutputFinalizer().finalize(
        task_id="windows-long-external-collision",
        artifacts=[_artifact(staging, suggested_name)],
        policy=OutputPolicy(output_dir=str(output), overwrite_mode="rename"),
    )

    assert result.success is True, result.diagnostics
    logical_path = result.artifacts[0].staging_path
    assert Path(logical_path).stem.endswith("_002")
    assert not logical_path.startswith("\\\\?\\")
    assert _io_bytes(logical_path) == b"docwen payload"


@pytest.mark.skipif(os.name != "nt", reason="Win32 extended-path boundary")
def test_windows_long_output_root_recovers_stale_and_exception_temps(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    output = _output_dir_with_total_length(tmp_path, 270)
    staging = tmp_path / "staging.md"
    staging.write_bytes(b"first complete")
    first = OutputFinalizer().finalize(
        task_id="windows-long-root-create",
        artifacts=[_artifact(staging, "first.md")],
        policy=OutputPolicy(output_dir=str(output)),
    )
    assert first.success is True, first.diagnostics
    assert not first.artifacts[0].staging_path.startswith("\\\\?\\")
    assert _io_bytes(first.artifacts[0].staging_path) == b"first complete"

    stale = output / ".__docwen-finalizer-interrupted"
    user_file = output / ".user-private"
    OutputFinalizer._io_path(stale).write_bytes(b"stale")
    OutputFinalizer._io_path(user_file).write_bytes(b"keep")
    observed_now = time.time()
    monkeypatch.setattr(time, "time", lambda: observed_now + (48 * 60 * 60))
    failing_staging = tmp_path / "failing.md"
    failing_staging.write_bytes(b"new complete")

    def partial_then_fail(_source: str, target: str, _cancellation: object) -> None:
        OutputFinalizer._io_path(target).write_bytes(b"private partial")
        raise OSError("injected long-root copy failure")

    monkeypatch.setattr(OutputFinalizer, "_copy_to_temp", staticmethod(partial_then_fail))
    failed = OutputFinalizer().finalize(
        task_id="windows-long-root-cleanup",
        artifacts=[_artifact(failing_staging, "second.md")],
        policy=OutputPolicy(output_dir=str(output), overwrite_mode="overwrite"),
    )

    assert failed.success is False
    assert not OutputFinalizer._io_path(stale).exists()
    assert _io_bytes(user_file) == b"keep"
    assert not any(name.startswith(".__docwen-finalizer-") for name in _io_names(output))


@pytest.mark.skipif(os.name != "nt", reason="Win32 extended-path boundary")
def test_windows_internal_io_prefix_never_leaks_into_failure_surfaces(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    output = _output_dir_with_total_length(tmp_path, 270)
    staging = tmp_path / "staging.md"
    staging.write_bytes(b"complete")

    def fail_with_internal_paths(_source: str, _target: str, _cancellation: Any) -> None:
        raise OSError(
            errno.EXDEV,
            "injected cross-device failure",
            r"\\?\C:\private-source",
            None,
            r"\\?\C:\private-destination",
        )

    monkeypatch.setattr(OutputFinalizer, "_copy_to_temp", staticmethod(fail_with_internal_paths))
    result = OutputFinalizer().finalize(
        task_id="windows-prefix-diagnostic",
        artifacts=[_artifact(staging, "report.md")],
        policy=OutputPolicy(output_dir=str(output)),
    )

    assert result.success is False
    public_surface = repr((result.diagnostics, result.error, result.metrics, result.artifacts))
    assert "\\\\?\\" not in public_surface
    assert "private-source" in result.diagnostics[0].message
    assert "private-destination" in result.diagnostics[0].message


@pytest.mark.skipif(os.name != "nt", reason="Win32 extended-path boundary")
def test_windows_long_staging_copies_to_short_output(tmp_path: Path) -> None:
    staging_root = _output_dir_with_total_length(tmp_path, 270)
    staging = staging_root / "staging.md"
    OutputFinalizer._io_path(staging).parent.mkdir(parents=True, exist_ok=True)
    OutputFinalizer._io_path(staging).write_bytes(b"deep staging payload")
    output = tmp_path / "short-output"

    result = OutputFinalizer().finalize(
        task_id="windows-long-staging",
        artifacts=[_artifact(staging)],
        policy=OutputPolicy(output_dir=str(output)),
    )

    assert result.success is True, result.diagnostics
    assert _io_bytes(result.artifacts[0].staging_path) == b"deep staging payload"


@pytest.mark.skipif(os.name != "nt", reason="Win32 extended-path boundary")
def test_windows_long_same_dir_reuses_identical_retained_input(tmp_path: Path) -> None:
    work = _output_dir_with_total_length(tmp_path, 270)
    input_path = work / "source.png"
    OutputFinalizer._io_path(input_path).parent.mkdir(parents=True, exist_ok=True)
    OutputFinalizer._io_path(input_path).write_bytes(b"same retained bytes")
    primary = tmp_path / "source.md"
    retained = tmp_path / "retained.png"
    primary.write_text("![[source.png]]", encoding="utf-8")
    retained.write_bytes(b"same retained bytes")
    artifacts = [
        _artifact(primary, "source.md"),
        ArtifactManifest(
            artifact_id="retained",
            kind="image",
            staging_path=str(retained),
            suggested_name="source.png",
            media_type="image/png",
            is_primary=False,
        ),
    ]

    result = OutputFinalizer().finalize(
        task_id="windows-long-retained-input",
        artifacts=artifacts,
        policy=OutputPolicy(overwrite_mode="rename"),
        input_path=str(input_path),
    )

    assert result.success is True, result.diagnostics
    retained_result = next(artifact for artifact in result.artifacts if artifact.artifact_id == "retained")
    assert retained_result.staging_path == os.path.abspath(input_path)
    assert not retained_result.staging_path.startswith("\\\\?\\")
    assert retained_result.metadata["reused_input"] is True
    assert _io_bytes(input_path) == b"same retained bytes"

    changed = tmp_path / "changed.png"
    changed.write_bytes(b"different retained bytes")
    changed_artifacts = [
        _artifact(primary, "source.md"),
        ArtifactManifest(
            artifact_id="retained-changed",
            kind="image",
            staging_path=str(changed),
            suggested_name="source.png",
            media_type="image/png",
            is_primary=False,
        ),
    ]
    changed_result = OutputFinalizer().finalize(
        task_id="windows-long-retained-input-changed",
        artifacts=changed_artifacts,
        policy=OutputPolicy(overwrite_mode="rename"),
        input_path=str(input_path),
    )

    assert changed_result.success is False
    assert changed_result.error is not None
    assert changed_result.error.diagnostic_code == "FINALIZER_PARTIAL"
    assert _io_bytes(input_path) == b"same retained bytes"


def test_selected_root_resolution_is_the_containment_boundary(tmp_path: Path, monkeypatch) -> None:
    logical_root = tmp_path / "logical-root"
    physical_root = tmp_path / "physical-root"
    logical_target = logical_root / "nested" / "report.md"
    real_realpath = os.path.realpath

    def root_alias_realpath(path: os.PathLike[str] | str) -> str:
        absolute = os.path.abspath(path)
        logical = os.path.abspath(logical_root)
        if absolute == logical or absolute.startswith(logical + os.sep):
            relative = os.path.relpath(absolute, logical)
            return os.path.abspath(physical_root / relative)
        return real_realpath(absolute)

    monkeypatch.setattr(os.path, "realpath", root_alias_realpath)

    final_path, suggested = OutputFinalizer._safe_final_path(str(logical_root), "nested/report.md")

    assert final_path == str(logical_target)
    assert suggested == os.path.normpath("nested/report.md")


def test_skip_rechecks_containment_at_commit_boundary(tmp_path: Path, monkeypatch) -> None:
    staging = tmp_path / "staging.md"
    staging.write_bytes(b"new")
    output = tmp_path / "output"
    output.mkdir()
    existing = output / "report.md"
    existing.write_bytes(b"existing")
    commit_started = False
    real_commit = OutputFinalizer._commit_prepared
    real_ensure = OutputFinalizer._ensure_contained

    def guarded_ensure(output_dir: str, final_path: str) -> None:
        if commit_started:
            raise ValueError("simulated descendant swap")
        real_ensure(output_dir, final_path)

    def swap_then_commit(item: Any, output_dir: str, overwrite_mode: str) -> Any:
        nonlocal commit_started
        commit_started = True
        try:
            return real_commit(item, output_dir, overwrite_mode)
        finally:
            commit_started = False

    monkeypatch.setattr(OutputFinalizer, "_ensure_contained", staticmethod(guarded_ensure))
    monkeypatch.setattr(OutputFinalizer, "_commit_prepared", staticmethod(swap_then_commit))
    result = OutputFinalizer().finalize(
        task_id="skip-containment-recheck",
        artifacts=[_artifact(staging)],
        policy=OutputPolicy(output_dir=str(output), overwrite_mode="skip"),
    )

    assert result.success is False
    assert result.artifacts == []
    assert existing.read_bytes() == b"existing"


def test_reuse_rechecks_containment_at_commit_boundary(tmp_path: Path, monkeypatch) -> None:
    output = tmp_path / "output"
    output.mkdir()
    input_path = output / "source.png"
    input_path.write_bytes(b"same")
    staging = tmp_path / "retained.png"
    staging.write_bytes(b"same")
    retained = ArtifactManifest(
        artifact_id="retained",
        kind="image",
        staging_path=str(staging),
        suggested_name=input_path.name,
        media_type="image/png",
        is_primary=False,
    )
    commit_started = False
    real_commit = OutputFinalizer._commit_prepared
    real_ensure = OutputFinalizer._ensure_contained

    def guarded_ensure(output_dir: str, final_path: str) -> None:
        if commit_started:
            raise ValueError("simulated descendant swap")
        real_ensure(output_dir, final_path)

    def swap_then_commit(item: Any, output_dir: str, overwrite_mode: str) -> Any:
        nonlocal commit_started
        commit_started = True
        try:
            return real_commit(item, output_dir, overwrite_mode)
        finally:
            commit_started = False

    monkeypatch.setattr(OutputFinalizer, "_ensure_contained", staticmethod(guarded_ensure))
    monkeypatch.setattr(OutputFinalizer, "_commit_prepared", staticmethod(swap_then_commit))
    result = OutputFinalizer().finalize(
        task_id="reuse-containment-recheck",
        artifacts=[retained],
        policy=OutputPolicy(output_dir=str(output), overwrite_mode="rename"),
        input_path=str(input_path),
    )

    assert result.success is False
    assert result.artifacts == []
    assert input_path.read_bytes() == b"same"

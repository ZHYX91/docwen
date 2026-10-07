"""Real directory-alias publication authorization regressions."""

from __future__ import annotations

from tests.support.subprocess_runner import run_subprocess

from ._output_finalizer_support import (
    ARTIFACT_KIND_PRIMARY,
    ArtifactManifest,
    OutputFinalizer,
    OutputPolicy,
    Path,
    os,
    pytest,
)
from ._output_finalizer_support import finalizer as finalizer

pytestmark = pytest.mark.integration


def _directory_alias(link: Path, target: Path) -> None:
    if os.name == "nt":
        completed = run_subprocess(
            ["cmd.exe", "/d", "/c", "mklink", "/J", str(link), str(target.resolve(strict=True))],
        )
        assert completed.returncode == 0, f"{completed.stdout}{completed.stderr}"
    else:
        link.symlink_to(target, target_is_directory=True)


class TestOutputFinalizerAliases:
    @pytest.mark.parametrize("alias_input", [False, True])
    def test_implicit_output_refuses_directory_alias_of_input(
        self,
        finalizer: OutputFinalizer,
        tmp_path: Path,
        alias_input: bool,
    ) -> None:
        source_dir = tmp_path / "source"
        source_dir.mkdir()
        alias = tmp_path / "alias"
        _directory_alias(alias, source_dir)
        input_path = source_dir / "source.png"
        input_path.write_bytes(b"original source bytes")
        staged_path = tmp_path / "generated.png"
        staged_path.write_bytes(b"converted bytes")
        artifact = ArtifactManifest(
            artifact_id="primary",
            kind=ARTIFACT_KIND_PRIMARY,
            staging_path=str(staged_path),
            suggested_name=input_path.name,
            is_primary=True,
        )

        result = finalizer.finalize(
            task_id="aliased-source-collision",
            artifacts=[artifact],
            policy=OutputPolicy(output_dir=str(source_dir if alias_input else alias), overwrite_mode="overwrite"),
            input_path=str(alias / input_path.name if alias_input else input_path),
        )

        assert not result.success
        assert result.artifacts == []
        assert input_path.read_bytes() == b"original source bytes"
        assert not list(source_dir.glob(".__docwen-finalizer-*"))

    @pytest.mark.parametrize("same_bytes", [False, True])
    def test_retained_resource_alias_reuses_only_identical_input(
        self, finalizer: OutputFinalizer, tmp_path: Path, same_bytes: bool
    ) -> None:
        source_dir = tmp_path / "source"
        source_dir.mkdir()
        alias = tmp_path / "alias"
        _directory_alias(alias, source_dir)
        input_path = source_dir / "source.png"
        input_path.write_bytes(b"original source bytes")
        staged_path = tmp_path / "retained.png"
        staged_path.write_bytes(input_path.read_bytes() if same_bytes else b"changed resource bytes")
        artifact = ArtifactManifest(
            artifact_id="retained",
            kind="resource",
            staging_path=str(staged_path),
            suggested_name=input_path.name,
            is_primary=False,
        )
        result = finalizer.finalize(
            task_id="retained-alias",
            artifacts=[artifact],
            policy=OutputPolicy(output_dir=str(alias), overwrite_mode="overwrite"),
            input_path=str(input_path),
        )
        assert result.success is same_bytes
        assert input_path.read_bytes() == b"original source bytes"
        if same_bytes:
            assert result.artifacts[0].metadata["reused_input"] is True
        else:
            assert not result.artifacts

    @pytest.mark.parametrize("primary, explicit", [(False, False), (True, False), (True, True)])
    def test_commit_rechecks_retargeted_directory_alias(self, tmp_path: Path, primary: bool, explicit: bool) -> None:
        source_dir = tmp_path / "source"
        output_dir = tmp_path / "output"
        source_dir.mkdir()
        output_dir.mkdir()
        alias = tmp_path / "alias"
        _directory_alias(alias, output_dir)
        input_path = source_dir / "source.png"
        input_path.write_bytes(b"original source bytes")
        staged_path = tmp_path / "generated.png"
        staged_path.write_bytes(b"converted bytes")
        artifact = ArtifactManifest(
            artifact_id="primary",
            kind=ARTIFACT_KIND_PRIMARY if primary else "resource",
            staging_path=str(staged_path),
            suggested_name=input_path.name,
            is_primary=primary,
        )
        prepared = OutputFinalizer._prepare_artifact(
            artifact, str(alias), "overwrite", str(input_path), None, allow_primary_input_replacement=explicit
        )
        assert prepared.temp_path
        prepared.temp_path = str(Path(prepared.temp_path).resolve(strict=True))
        # Remove only the alias itself; neither directory's contents are removed.
        if os.name == "nt":
            alias.rmdir()
        else:
            alias.unlink()
        _directory_alias(alias, source_dir)
        try:
            with pytest.raises(ValueError, match="must not replace its input file"):
                OutputFinalizer._commit_prepared(prepared, str(alias), "overwrite")
            assert input_path.read_bytes() == b"original source bytes"
            assert not (output_dir / input_path.name).exists()
        finally:
            if prepared.temp_path:
                Path(prepared.temp_path).unlink(missing_ok=True)

    def test_explicit_canonical_input_alias_allows_in_place_output(self, tmp_path: Path) -> None:
        source_dir = tmp_path / "source"
        source_dir.mkdir()
        alias = tmp_path / "alias"
        _directory_alias(alias, source_dir)
        input_path = source_dir / "source.png"
        input_path.write_bytes(b"original source bytes")
        staged_path = tmp_path / "generated.png"
        staged_path.write_bytes(b"converted bytes")
        artifact = ArtifactManifest(
            artifact_id="primary",
            kind=ARTIFACT_KIND_PRIMARY,
            staging_path=str(staged_path),
            suggested_name=input_path.name,
            is_primary=True,
        )
        result = OutputFinalizer().finalize(
            task_id="explicit-canonical-in-place",
            artifacts=[artifact],
            policy=OutputPolicy(output_path=str(alias / input_path.name), overwrite_mode="overwrite"),
            input_path=str(input_path),
        )
        assert result.success
        assert input_path.read_bytes() == b"converted bytes"

    def test_explicit_in_place_alias_cannot_publish_to_a_changed_target(self, tmp_path: Path) -> None:
        source_dir = tmp_path / "source"
        other_dir = tmp_path / "other"
        source_dir.mkdir()
        other_dir.mkdir()
        input_path = source_dir / "source.png"
        other_path = other_dir / input_path.name
        input_path.write_bytes(b"original source")
        other_path.write_bytes(b"unrelated file")
        alias = tmp_path / "alias"
        _directory_alias(alias, source_dir)
        staged_path = tmp_path / "generated.png"
        staged_path.write_bytes(b"converted bytes")
        artifact = ArtifactManifest(
            artifact_id="primary",
            kind=ARTIFACT_KIND_PRIMARY,
            staging_path=str(staged_path),
            suggested_name=input_path.name,
            is_primary=True,
        )
        prepared = OutputFinalizer._prepare_artifact(
            artifact, str(alias), "overwrite", str(input_path), None, allow_primary_input_replacement=True
        )
        assert prepared.temp_path
        prepared.temp_path = str(Path(prepared.temp_path).resolve(strict=True))
        if os.name == "nt":
            alias.rmdir()
        else:
            alias.unlink()
        _directory_alias(alias, other_dir)
        try:
            with pytest.raises(ValueError, match="in-place output target changed"):
                OutputFinalizer._commit_prepared(prepared, str(alias), "overwrite")
            assert input_path.read_bytes() == b"original source"
            assert other_path.read_bytes() == b"unrelated file"
        finally:
            if prepared.temp_path:
                Path(prepared.temp_path).unlink(missing_ok=True)

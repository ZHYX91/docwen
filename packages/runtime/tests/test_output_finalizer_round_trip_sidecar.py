"""Final placement keeps one DOCX and its public sidecar as a matched pair."""

from __future__ import annotations

import contextlib
import threading
from collections.abc import Iterator
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from typing import Any

import pytest

from docwen_core.models.artifact import ArtifactManifest
from docwen_core.models.request import OutputPolicy
from docwen_core.round_trip_sidecar import (
    DOCX_MEDIA_TYPE,
    ROUND_TRIP_SIDECAR_MEDIA_TYPE,
    ROUND_TRIP_SIDECAR_OWNER_METADATA,
    ROUND_TRIP_SIDECAR_SCHEMA,
    ROUND_TRIP_SIDECAR_SCHEMA_METADATA,
)
from docwen_runtime.output.finalizer import OutputFinalizer

pytestmark = pytest.mark.integration


def _artifacts(staging: Path) -> list[ArtifactManifest]:
    staging.mkdir()
    docx = staging / "artifact.docx"
    sidecar = staging / "artifact.docwen"
    docx.write_bytes(b"new-docx")
    sidecar.write_bytes(b"new-sidecar")
    primary = ArtifactManifest(
        artifact_id="task-docx",
        kind="primary",
        staging_path=str(docx),
        suggested_name="source.docx",
        media_type=DOCX_MEDIA_TYPE,
        is_primary=True,
    )
    companion = ArtifactManifest(
        artifact_id="task-sidecar",
        kind="auxiliary",
        staging_path=str(sidecar),
        suggested_name="source.docx.docwen",
        media_type=ROUND_TRIP_SIDECAR_MEDIA_TYPE,
        metadata={
            ROUND_TRIP_SIDECAR_SCHEMA_METADATA: ROUND_TRIP_SIDECAR_SCHEMA,
            ROUND_TRIP_SIDECAR_OWNER_METADATA: primary.artifact_id,
        },
    )
    return [primary, companion]


def test_exact_output_path_renames_sidecar_to_final_docx_basename(tmp_path: Path) -> None:
    output = tmp_path / "output"
    result = OutputFinalizer().finalize(
        "task",
        _artifacts(tmp_path / "staging"),
        OutputPolicy(output_path=str(output / "chosen.docx"), overwrite_mode="error"),
    )

    assert result.success, result.error
    assert [Path(item.staging_path).name for item in result.artifacts] == [
        "chosen.docx",
        "chosen.docx.docwen",
    ]
    assert (output / "chosen.docx").read_bytes() == b"new-docx"
    assert (output / "chosen.docx.docwen").read_bytes() == b"new-sidecar"


def test_rename_policy_selects_one_free_basename_for_both_files(tmp_path: Path) -> None:
    output = tmp_path / "output"
    output.mkdir()
    (output / "source.docx.docwen").write_bytes(b"unrelated-existing-sidecar")

    result = OutputFinalizer().finalize(
        "task",
        _artifacts(tmp_path / "staging"),
        OutputPolicy(output_dir=str(output), overwrite_mode="rename"),
    )

    assert result.success, result.error
    assert [Path(item.staging_path).name for item in result.artifacts] == [
        "source_001.docx",
        "source_001.docx.docwen",
    ]
    assert (output / "source.docx.docwen").read_bytes() == b"unrelated-existing-sidecar"


def test_concurrent_rename_recomputes_the_matched_basename_under_the_output_lock(tmp_path: Path) -> None:
    output = tmp_path / "output"
    arrivals = threading.Barrier(2)

    class BarrierFinalizer(OutputFinalizer):
        @classmethod
        @contextlib.contextmanager
        def _finalization_locks(
            cls,
            paths: tuple[str, ...],
            cancellation: Any,
        ) -> Iterator[None]:
            arrivals.wait(timeout=5)
            with super()._finalization_locks(paths, cancellation):
                yield

    def finalize(index: int):
        return BarrierFinalizer().finalize(
            f"task-{index}",
            _artifacts(tmp_path / f"staging-{index}"),
            OutputPolicy(output_dir=str(output), overwrite_mode="rename"),
        )

    with ThreadPoolExecutor(max_workers=2) as executor:
        results = list(executor.map(finalize, range(2)))

    assert all(result.success for result in results)
    assert sorted(tuple(Path(item.staging_path).name for item in result.artifacts) for result in results) == [
        ("source.docx", "source.docx.docwen"),
        ("source_001.docx", "source_001.docx.docwen"),
    ]
    assert sorted(path.name for path in output.iterdir()) == [
        "source.docx",
        "source.docx.docwen",
        "source_001.docx",
        "source_001.docx.docwen",
    ]


@pytest.mark.parametrize("overwrite_mode", ["error", "skip"])
def test_error_and_skip_do_not_publish_half_a_pair(tmp_path: Path, overwrite_mode: str) -> None:
    output = tmp_path / "output"
    output.mkdir()
    (output / "source.docx").write_bytes(b"existing-docx")

    result = OutputFinalizer().finalize(
        "task",
        _artifacts(tmp_path / f"staging-{overwrite_mode}"),
        OutputPolicy(output_dir=str(output), overwrite_mode=overwrite_mode),
    )

    assert not result.success
    assert result.artifacts == []
    assert (output / "source.docx").read_bytes() == b"existing-docx"
    assert not (output / "source.docx.docwen").exists()

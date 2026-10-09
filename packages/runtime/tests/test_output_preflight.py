"""Output probes must fail before conversion and leave user entries intact."""

from __future__ import annotations

import errno
import os
from pathlib import Path

import pytest

from docwen_core.models.request import OutputPolicy
from docwen_runtime.output.finalizer import AtomicPublishUnavailable, OutputFinalizer, OutputPreflightError

pytestmark = pytest.mark.integration


@pytest.mark.parametrize("grouped", [False, True])
@pytest.mark.parametrize("nested", [False, True])
def test_preflight_probes_selected_filesystem_without_creating_outputs(tmp_path: Path, grouped: bool, nested: bool):
    existing = tmp_path / "keep.txt"
    existing.write_bytes(b"unchanged")
    output = tmp_path / "new" / "nested" if nested else tmp_path

    OutputFinalizer().preflight(OutputPolicy(output_dir=str(output)), group_outputs=grouped)

    assert list(tmp_path.iterdir()) == [existing]
    assert existing.read_bytes() == b"unchanged"


def test_preflight_uses_explicit_output_parent(tmp_path: Path):
    parent = tmp_path / "chosen"
    parent.mkdir()
    output = parent / "existing.txt"
    output.write_bytes(b"original")
    OutputFinalizer().preflight(OutputPolicy(output_path=str(output), overwrite_mode="overwrite"))
    assert list(parent.iterdir()) == [output]
    assert output.read_bytes() == b"original"


def test_read_only_operation_never_probes_output(monkeypatch, tmp_path: Path):
    def unexpected(*args, **kwargs):
        pytest.fail("Read-only operation attempted an output probe")

    monkeypatch.setattr("docwen_runtime.output.finalizer.tempfile.TemporaryDirectory", unexpected)
    OutputFinalizer().preflight(OutputPolicy(output_dir=str(tmp_path), write_artifacts=False))


def test_unsupported_publication_cleans_probe_and_preserves_diagnostic(monkeypatch, tmp_path: Path):
    def unsupported(source, destination):
        assert Path(source).parent.parent == tmp_path
        raise AtomicPublishUnavailable(errno.EINVAL, "Choose another output folder")

    monkeypatch.setattr(OutputFinalizer, "_publish_directory_no_clobber", staticmethod(unsupported))
    with pytest.raises(AtomicPublishUnavailable, match="Choose another output folder"):
        OutputFinalizer().preflight(OutputPolicy(output_dir=str(tmp_path)), group_outputs=True)
    assert list(tmp_path.iterdir()) == []


def test_probe_rejects_a_primitive_that_silently_replaces_targets(monkeypatch, tmp_path: Path):
    monkeypatch.setattr(OutputFinalizer, "_publish_no_clobber", staticmethod(os.replace))
    with pytest.raises(AtomicPublishUnavailable, match="cannot prevent replacement"):
        OutputFinalizer().preflight(OutputPolicy(output_dir=str(tmp_path)))
    assert list(tmp_path.iterdir()) == []


def test_invalid_output_parent_is_actionable_and_unchanged(tmp_path: Path):
    output = tmp_path / "not-a-directory"
    output.write_bytes(b"original")
    with pytest.raises(OutputPreflightError, match="permissions or choose another folder"):
        OutputFinalizer().preflight(OutputPolicy(output_dir=str(output)))
    assert output.read_bytes() == b"original"

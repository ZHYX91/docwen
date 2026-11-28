"""Integrity tests for the pinned PyMuPDF Layout resource contract."""

from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace

import pytest

from docwen_runtime import pymupdf_layout_resources as resources

pytestmark = pytest.mark.unit


def _use_manifest(
    monkeypatch: pytest.MonkeyPatch,
    manifest: tuple[resources.PymupdfLayoutResourceSpec, ...],
) -> None:
    monkeypatch.setattr(resources, "_current_pymupdf_layout_resource_manifest", lambda: manifest)


def test_pinned_manifest_matches_the_installed_locked_distribution() -> None:
    verification = resources.verify_installed_pymupdf_layout_distribution()
    manifest = resources.pymupdf_layout_resource_manifest()

    assert resources.PYMUPDF_LAYOUT_DISTRIBUTION == "pymupdf-layout"
    assert resources.PYMUPDF_LAYOUT_DISTRIBUTION_VERSION == "1.28.2"
    assert len(manifest) == 19
    assert verification == resources.PymupdfLayoutResourceVerification(
        available=True,
        reason=None,
        resource_types=("onnx", "yaml"),
        resource_count=19,
    )


def test_platform_manifest_selection_is_explicit_and_unknown_platforms_are_not_supported() -> None:
    manifest = resources.PYMUPDF_LAYOUT_RESOURCE_MANIFEST
    for platform in ("win32", "linux", "darwin"):
        assert resources._pymupdf_layout_resource_manifest_for_platform(platform) is manifest
    assert resources._pymupdf_layout_resource_manifest_for_platform("freebsd14") is None
    assert len({spec.relative_path for spec in manifest}) == len(manifest)


@pytest.mark.parametrize(
    ("corruption", "expected_reason"),
    (
        ("size", "required_resource_size_mismatch"),
        ("hash", "required_resource_hash_mismatch"),
    ),
)
def test_manifest_rejects_size_and_hash_corruption_without_exposing_path(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
    corruption: str,
    expected_reason: str,
) -> None:
    manifest = resources.PYMUPDF_LAYOUT_RESOURCE_MANIFEST
    _use_manifest(monkeypatch, manifest)
    resource_root = tmp_path / "private-user"
    for spec in manifest:
        resource_file = resource_root.joinpath(*spec.relative_path.split("/"))
        resource_file.parent.mkdir(parents=True, exist_ok=True)
        resource_file.touch()
    if corruption == "hash":
        first = manifest[0]
        resource_root.joinpath(*first.relative_path.split("/")).write_bytes(bytes(first.size))

    verification = resources.verify_pymupdf_layout_resource_root(resource_root)

    assert verification.available is False
    assert verification.reason == expected_reason
    assert str(tmp_path) not in repr(verification)


def test_unknown_platform_fails_closed_before_reading_the_distribution(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    monkeypatch.setattr(resources.sys, "platform", "freebsd14")

    def unexpected_distribution_lookup(_name: str) -> None:
        raise AssertionError("unsupported platforms must not trust installed distribution bytes")

    monkeypatch.setattr(resources.importlib.metadata, "distribution", unexpected_distribution_lookup)

    verification = resources.verify_pymupdf_layout_resource_root(tmp_path / "private-user")
    installed_verification = resources.verify_installed_pymupdf_layout_distribution()

    assert verification == resources.PymupdfLayoutResourceVerification(
        available=False,
        reason="unsupported_resource_platform",
        resource_types=(),
        resource_count=0,
    )
    assert installed_verification == verification
    with pytest.raises(RuntimeError, match=r"^unsupported_resource_platform$"):
        resources.pymupdf_layout_resource_manifest()


def test_distribution_upgrade_requires_an_explicit_manifest_update(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(
        resources.importlib.metadata,
        "distribution",
        lambda _name: SimpleNamespace(version="1.28.0"),
    )

    verification = resources.verify_installed_pymupdf_layout_distribution()

    assert verification.available is False
    assert verification.reason == "distribution_version_mismatch"


def test_distribution_file_list_must_match_the_pinned_manifest(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(
        resources.importlib.metadata,
        "distribution",
        lambda _name: SimpleNamespace(version="1.28.2", files=()),
    )

    verification = resources.verify_installed_pymupdf_layout_distribution()

    assert verification.available is False
    assert verification.reason == "distribution_resource_manifest_mismatch"

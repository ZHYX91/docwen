"""Pinned integrity contract for PyMuPDF Layout model resources.

The dependency deliberately does not auto-upgrade this contract.  A
``pymupdf-layout`` upgrade must update the version and every resource digest in
this module so source, build, packaged, and runtime checks all agree on the
same bytes.
"""

from __future__ import annotations

import hashlib
import importlib.metadata
import sys
from dataclasses import dataclass
from pathlib import Path, PurePosixPath

PYMUPDF_LAYOUT_DISTRIBUTION = "pymupdf-layout"
PYMUPDF_LAYOUT_DISTRIBUTION_VERSION = "1.27.2.2"
PYMUPDF_LAYOUT_SOURCE_RESOURCE_ROOT = PurePosixPath("pymupdf/layout/resources")


@dataclass(frozen=True, slots=True)
class PymupdfLayoutResourceSpec:
    """One immutable file in the pinned PyMuPDF Layout resource set."""

    relative_path: str
    size: int
    sha256: str


@dataclass(frozen=True, slots=True)
class PymupdfLayoutResourceVerification:
    """Path-free resource verification result safe for public projections."""

    available: bool
    reason: str | None
    resource_types: tuple[str, ...]
    resource_count: int


PYMUPDF_LAYOUT_WINDOWS_RESOURCE_MANIFEST = (
    PymupdfLayoutResourceSpec(
        relative_path="onnx/feature_imf1.onnx",
        size=1_116_432,
        sha256="fac3e029a8c4d1c3e6cc9a6195bca8429175771db3c82c3d42f6ea9e9b8b4520",
    ),
    PymupdfLayoutResourceSpec(
        relative_path="onnx/layout_imf1.onnx",
        size=6_137_495,
        sha256="80a39fbd859999d812fdafe35b6aa1850f22be13057625dca86d59e58f8858e8",
    ),
    PymupdfLayoutResourceSpec(
        relative_path="onnx/layout_imf1.yaml",
        size=4_314,
        sha256="009b475fae8a8f1ae1e8e8276f31b3479f6b3438f7952cdd5eaf8dd45a4b8bac",
    ),
    PymupdfLayoutResourceSpec(
        relative_path="onnx/layout_rf2.4.1+imf1.onnx",
        size=6_453_771,
        sha256="77db5779c368084abfaed1fa08522c753776561d15b971a0227099a473e2871c",
    ),
    PymupdfLayoutResourceSpec(
        relative_path="onnx/layout_rf2.4.1+imf1.yaml",
        size=7_755,
        sha256="fcdbbd1166a625fe126c18ee30ff8b7274f4f6286141747e4a6449d46ec51f45",
    ),
    PymupdfLayoutResourceSpec(
        relative_path="onnx/layout_rf2.4.1.onnx",
        size=5_997_803,
        sha256="1db0a751714404222c80c23f4613e611c2bef0e7b10e0ec8338b6aeb0e4eee02",
    ),
    PymupdfLayoutResourceSpec(
        relative_path="onnx/layout_rf2.4.1.yaml",
        size=7_742,
        sha256="975d864374a3ce333c948a16e5083f2a01e1688145748106b38228bc882ca22a",
    ),
)

PYMUPDF_LAYOUT_POSIX_RESOURCE_MANIFEST = (
    PymupdfLayoutResourceSpec(
        relative_path="onnx/feature_imf1.onnx",
        size=1_116_432,
        sha256="fac3e029a8c4d1c3e6cc9a6195bca8429175771db3c82c3d42f6ea9e9b8b4520",
    ),
    PymupdfLayoutResourceSpec(
        relative_path="onnx/layout_imf1.onnx",
        size=6_137_495,
        sha256="80a39fbd859999d812fdafe35b6aa1850f22be13057625dca86d59e58f8858e8",
    ),
    PymupdfLayoutResourceSpec(
        relative_path="onnx/layout_imf1.yaml",
        size=4_100,
        sha256="09c6e01dd3ee99f5efe4889cba8ed22e48c608bd00ed8a6304c751e8445dbc2c",
    ),
    PymupdfLayoutResourceSpec(
        relative_path="onnx/layout_rf2.4.1+imf1.onnx",
        size=6_453_771,
        sha256="77db5779c368084abfaed1fa08522c753776561d15b971a0227099a473e2871c",
    ),
    PymupdfLayoutResourceSpec(
        relative_path="onnx/layout_rf2.4.1+imf1.yaml",
        size=7_406,
        sha256="c8198e97a1c58ed0cd765c0208ccf5408fb1563715f8f134096eab1ad9e72316",
    ),
    PymupdfLayoutResourceSpec(
        relative_path="onnx/layout_rf2.4.1.onnx",
        size=5_997_803,
        sha256="1db0a751714404222c80c23f4613e611c2bef0e7b10e0ec8338b6aeb0e4eee02",
    ),
    PymupdfLayoutResourceSpec(
        relative_path="onnx/layout_rf2.4.1.yaml",
        size=7_394,
        sha256="c0ebe52f9eea9ff2387d2b5317c1e84065b631e4f4ffed33f0cacd28232d15bd",
    ),
)


def _pymupdf_layout_resource_manifest_for_platform(
    platform: str,
) -> tuple[PymupdfLayoutResourceSpec, ...] | None:
    if platform == "win32":
        return PYMUPDF_LAYOUT_WINDOWS_RESOURCE_MANIFEST
    if platform in {"linux", "darwin"}:
        return PYMUPDF_LAYOUT_POSIX_RESOURCE_MANIFEST
    return None


def _current_pymupdf_layout_resource_manifest() -> tuple[PymupdfLayoutResourceSpec, ...] | None:
    return _pymupdf_layout_resource_manifest_for_platform(sys.platform)


def pymupdf_layout_resource_manifest() -> tuple[PymupdfLayoutResourceSpec, ...]:
    """Return the exact raw-byte manifest for the running supported platform."""

    manifest = _current_pymupdf_layout_resource_manifest()
    if manifest is None:
        raise RuntimeError("unsupported_resource_platform")
    return manifest


def pymupdf_layout_resource_paths() -> tuple[str, ...]:
    """Return the stable, relative resource paths from the pinned contract."""

    return tuple(spec.relative_path for spec in pymupdf_layout_resource_manifest())


def _resource_types(paths: tuple[str, ...]) -> tuple[str, ...]:
    return tuple(sorted({PurePosixPath(path).suffix.lower().removeprefix(".") for path in paths}))


def _result(
    *,
    available: bool,
    reason: str | None,
    present_paths: tuple[str, ...],
) -> PymupdfLayoutResourceVerification:
    return PymupdfLayoutResourceVerification(
        available=available,
        reason=reason,
        resource_types=_resource_types(present_paths),
        resource_count=len(present_paths),
    )


def _resource_path(resource_root: Path, relative_path: str) -> Path:
    return resource_root.joinpath(*PurePosixPath(relative_path).parts)


def verify_pymupdf_layout_resource_root(resource_root: Path) -> PymupdfLayoutResourceVerification:
    """Verify one resource root against the exact pinned bytes.

    The result contains stable reason codes and aggregate counts only.  It never
    includes the supplied filesystem path.
    """

    manifest = _current_pymupdf_layout_resource_manifest()
    if manifest is None:
        return _result(available=False, reason="unsupported_resource_platform", present_paths=())

    present_paths = tuple(
        spec.relative_path for spec in manifest if _resource_path(resource_root, spec.relative_path).is_file()
    )
    if len(present_paths) != len(manifest):
        return _result(available=False, reason="required_resource_missing", present_paths=present_paths)

    for spec in manifest:
        digest = hashlib.sha256()
        size = 0
        try:
            with _resource_path(resource_root, spec.relative_path).open("rb") as handle:
                while chunk := handle.read(1024 * 1024):
                    size += len(chunk)
                    digest.update(chunk)
        except OSError:
            return _result(available=False, reason="required_resource_unreadable", present_paths=present_paths)
        if size != spec.size:
            return _result(available=False, reason="required_resource_size_mismatch", present_paths=present_paths)
        if digest.hexdigest() != spec.sha256:
            return _result(available=False, reason="required_resource_hash_mismatch", present_paths=present_paths)

    return _result(available=True, reason=None, present_paths=present_paths)


def verify_installed_pymupdf_layout_distribution() -> PymupdfLayoutResourceVerification:
    """Verify the installed distribution version, file list, sizes, and hashes."""

    if _current_pymupdf_layout_resource_manifest() is None:
        return _result(available=False, reason="unsupported_resource_platform", present_paths=())

    try:
        distribution = importlib.metadata.distribution(PYMUPDF_LAYOUT_DISTRIBUTION)
    except importlib.metadata.PackageNotFoundError:
        return _result(available=False, reason="distribution_not_available", present_paths=())

    if distribution.version != PYMUPDF_LAYOUT_DISTRIBUTION_VERSION:
        return _result(available=False, reason="distribution_version_mismatch", present_paths=())

    distribution_files = distribution.files
    if distribution_files is None:
        return _result(available=False, reason="distribution_manifest_missing", present_paths=())

    actual_paths: list[str] = []
    for package_path in distribution_files:
        normalized_path = PurePosixPath(str(package_path).replace("\\", "/"))
        try:
            resource_path = normalized_path.relative_to(PYMUPDF_LAYOUT_SOURCE_RESOURCE_ROOT)
        except ValueError:
            continue
        if resource_path.parts:
            actual_paths.append(resource_path.as_posix())

    expected_paths = pymupdf_layout_resource_paths()
    if set(actual_paths) != set(expected_paths) or len(actual_paths) != len(expected_paths):
        return _result(
            available=False,
            reason="distribution_resource_manifest_mismatch",
            present_paths=tuple(sorted(actual_paths)),
        )

    try:
        resource_root = Path(str(distribution.locate_file(PYMUPDF_LAYOUT_SOURCE_RESOURCE_ROOT)))
    except (OSError, TypeError, ValueError):
        return _result(available=False, reason="distribution_resource_unreadable", present_paths=expected_paths)
    return verify_pymupdf_layout_resource_root(resource_root)


__all__ = [
    "PYMUPDF_LAYOUT_DISTRIBUTION",
    "PYMUPDF_LAYOUT_DISTRIBUTION_VERSION",
    "PYMUPDF_LAYOUT_POSIX_RESOURCE_MANIFEST",
    "PYMUPDF_LAYOUT_SOURCE_RESOURCE_ROOT",
    "PYMUPDF_LAYOUT_WINDOWS_RESOURCE_MANIFEST",
    "PymupdfLayoutResourceSpec",
    "PymupdfLayoutResourceVerification",
    "pymupdf_layout_resource_manifest",
    "pymupdf_layout_resource_paths",
    "verify_installed_pymupdf_layout_distribution",
    "verify_pymupdf_layout_resource_root",
]

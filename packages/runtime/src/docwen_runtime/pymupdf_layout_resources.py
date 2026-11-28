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
PYMUPDF_LAYOUT_DISTRIBUTION_VERSION = "1.28.2"
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


PYMUPDF_LAYOUT_RESOURCE_MANIFEST = (
    PymupdfLayoutResourceSpec(
        relative_path="onnx/feature_imf1.onnx",
        size=1_264_599,
        sha256="27e30224174d760ca28536b7ca498133745b6b73ad091d67b2acb3617ce1ae51",
    ),
    PymupdfLayoutResourceSpec(
        relative_path="onnx/feature_imf2.onnx",
        size=1_330_260,
        sha256="da2b70759311a4c5b22f427845ee8201f41d1945d609bcc68e299d7bbea4877e",
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
        size=8_195_105,
        sha256="9602ca3f4e4f34b50551ccd2af81bea4af7010bd61b2cc10e77de6f7b125cb20",
    ),
    PymupdfLayoutResourceSpec(
        relative_path="onnx/layout_rf2.4.1+imf1.yaml",
        size=6_561,
        sha256="afdb4f7ec55900e5fb26239a9ceb78c814b6726c05c5a47eec1efb7beb9ae6c3",
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
    PymupdfLayoutResourceSpec(
        relative_path="onnx/table_grid_model_v1.onnx",
        size=683_875,
        sha256="ca1bbdcf9ed6f322899ca70cc5043f3e6939aa2f9edcdd3a65168d28e717e05f",
    ),
    PymupdfLayoutResourceSpec(
        relative_path="onnx/table_grid_model_v1_old.onnx",
        size=683_875,
        sha256="ca1bbdcf9ed6f322899ca70cc5043f3e6939aa2f9edcdd3a65168d28e717e05f",
    ),
    PymupdfLayoutResourceSpec(
        relative_path="onnx/table_grid_model_v1a.onnx",
        size=1_071_242,
        sha256="82b1eed0bada8872156f46ec9ffc067909d6942c4a2d1f684f2da926ef4b5417",
    ),
    PymupdfLayoutResourceSpec(
        relative_path="onnx/table_grid_model_v1t.onnx",
        size=3_904_284,
        sha256="2f10e54d183a7a912fe985070c4d752a8b121e400909250e6e1a75a37bf7d310",
    ),
    PymupdfLayoutResourceSpec(
        relative_path="onnx/table_grid_model_v2_conn.onnx",
        size=145_910,
        sha256="0097e1986f0dee4cd174fb73725dfe31607a91ed7a8a74f7cc649b4148881745",
    ),
    PymupdfLayoutResourceSpec(
        relative_path="onnx/table_grid_model_v2_grid.onnx",
        size=1_067_089,
        sha256="53c03e63b4882adf99466736ca2b3f22150a6826d398730952cf00320274b363",
    ),
    PymupdfLayoutResourceSpec(
        relative_path="onnx/table_grid_model_v2c.onnx",
        size=7_019_379,
        sha256="f83756e4a1986274aada2b457fd19ea0ded0303617636cb8f0d456ed29f0a698",
    ),
    PymupdfLayoutResourceSpec(
        relative_path="onnx/table_grid_model_v3.onnx",
        size=1_649_152,
        sha256="805212c764f0eb461fd6f419c5892d0a03b4ae727a1f83cbabd21e6ae807a88a",
    ),
    PymupdfLayoutResourceSpec(
        relative_path="onnx/table_grid_model_v4.onnx",
        size=3_652_791,
        sha256="9c41b156c94cdba4a9b474b5d9aab74288f1460207645dac5a38368013d32fe4",
    ),
    PymupdfLayoutResourceSpec(
        relative_path="onnx/table_grid_model_v4_do.onnx",
        size=3_652_791,
        sha256="9c41b156c94cdba4a9b474b5d9aab74288f1460207645dac5a38368013d32fe4",
    ),
    PymupdfLayoutResourceSpec(
        relative_path="onnx/table_grid_model_v4_ep.onnx",
        size=4_833_323,
        sha256="ec5d393e6fce44ac692ea75e11f0183ea4dcfca1ea434abc38395a1396cd815e",
    ),
)


def _pymupdf_layout_resource_manifest_for_platform(
    platform: str,
) -> tuple[PymupdfLayoutResourceSpec, ...] | None:
    if platform in {"win32", "linux", "darwin"}:
        return PYMUPDF_LAYOUT_RESOURCE_MANIFEST
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
    "PYMUPDF_LAYOUT_RESOURCE_MANIFEST",
    "PYMUPDF_LAYOUT_SOURCE_RESOURCE_ROOT",
    "PymupdfLayoutResourceSpec",
    "PymupdfLayoutResourceVerification",
    "pymupdf_layout_resource_manifest",
    "pymupdf_layout_resource_paths",
    "verify_installed_pymupdf_layout_distribution",
    "verify_pymupdf_layout_resource_root",
]

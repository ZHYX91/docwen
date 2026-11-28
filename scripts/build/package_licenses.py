"""Collect the licenses of binaries that the release redistributes."""

from __future__ import annotations

import re
import shutil
import subprocess
from importlib import metadata
from pathlib import Path

_PROJECT_NOTICES = ("LICENSE", "LICENSE_THIRD_PARTY.txt", "NOTICE.txt")
_QT_NOTICES = ("GPL-3.0-only.txt", "LGPL-3.0-only.txt", "Qt-GPL-exception-1.0.txt")
_DIST_INFO_NOTICES = {
    "easyofd": ("LICENSE",),
    "img2pdf": ("LICENSE",),
    "pillow": ("LICENSE",),
    "pillow-heif": ("LICENSE.txt", "LICENSES_bundled.txt"),
    "psutil": ("LICENSE",),
    "pymupdf": ("COPYING",),
    "pymupdf-layout": ("COPYING",),
    "pymupdf4llm": ("COPYING",),
}
_PACKAGE_NOTICES = {
    "onnxruntime": ("onnxruntime/LICENSE", "onnxruntime/ThirdPartyNotices.txt"),
    "opencv-python-headless": ("cv2/LICENSE.txt", "cv2/LICENSE-3RD-PARTY.txt"),
}


def _dist_info_notice(dist: metadata.Distribution, filename: str) -> Path:
    matches = [
        dist.locate_file(path)
        for path in dist.files or ()
        if path.name == filename and path.parts[0].endswith(".dist-info")
    ]
    if len(matches) != 1:
        raise RuntimeError(f"Expected one file for {dist.name}: {filename}")
    return Path(str(matches[0]))


def _copy_notice(source: Path, destination: Path) -> None:
    if not source.is_file() or source.stat().st_size == 0:
        raise RuntimeError(f"Missing release license material: {source}")
    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, destination)


def _source_entry(source_root: Path, release_version: str) -> str:
    try:
        revision = subprocess.check_output(
            ["git", "rev-parse", "--verify", "HEAD"], cwd=source_root, stderr=subprocess.DEVNULL, text=True
        ).strip()
    except (OSError, subprocess.CalledProcessError):
        revision = release_version
    if not re.fullmatch(r"(?:[0-9a-f]{40}|[0-9]+\.[0-9]+\.[0-9]+)", revision):
        raise RuntimeError("Release source could not be identified by commit or version")
    text = (source_root / "SOURCE_CODE.txt").read_text(encoding="utf-8")
    return (
        f"DocWen {release_version}\n"
        f"Corresponding project source: https://github.com/ZHYX91/docwen/tree/{revision}\n"
        f"Download project source: https://github.com/ZHYX91/docwen/archive/{revision}.tar.gz\n\n"
        f"{text}"
    )


def stage_package_licenses(source_root: Path, payload: Path, release_version: str, *, with_gui: bool) -> None:
    """Keep version-matched notices outside PyInstaller's import graph."""
    for filename in _PROJECT_NOTICES:
        _copy_notice(source_root / filename, payload / filename)
    (payload / "SOURCE_CODE.txt").write_text(
        _source_entry(source_root, release_version), encoding="utf-8", newline="\n"
    )

    licenses = payload / "licenses"
    for name, filename in (("rapidocr", "LICENSE"), ("x265", "COPYING")):
        _copy_notice(source_root / "licenses" / name / filename, licenses / name / filename)
    dist_info_notices = dict(_DIST_INFO_NOTICES)
    if with_gui:
        dist_info_notices["pyside6-fluent-widgets"] = ("LICENSE",)
        for filename in _QT_NOTICES:
            _copy_notice(source_root / "licenses" / "qt" / filename, licenses / "qt" / filename)

    for name, filenames in dist_info_notices.items():
        dist = metadata.distribution(name)
        for filename in filenames:
            _copy_notice(_dist_info_notice(dist, filename), licenses / name / filename)

    for name, members in _PACKAGE_NOTICES.items():
        dist = metadata.distribution(name)
        for member in members:
            _copy_notice(Path(str(dist.locate_file(member))), licenses / name / Path(member).name)

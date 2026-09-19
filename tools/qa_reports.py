"""Coverage arguments and compact reports for the owned QA runtime."""

from __future__ import annotations

import hashlib
import json
import stat
import tomllib
from pathlib import Path

VISIBILITY_REPORTS = (
    "skip_report.json",
    "not_collected_report.json",
    "slow_report.json",
    "subprocess_report.json",
    "missing_marker_report.json",
)


def coverage_arguments(repo_root: Path, runtime: Path) -> list[str]:
    config = tomllib.loads((repo_root / "pyproject.toml").read_text(encoding="utf-8"))
    threshold = config["tool"]["coverage"]["report"]["fail_under"]
    return [
        "--cov",
        "--cov-config=pyproject.toml",
        "--cov-report=term-missing:skip-covered",
        f"--cov-report=xml:{runtime / 'coverage.xml'}",
        f"--cov-fail-under={threshold}",
    ]


def coverage_checks(runtime: Path) -> list[list[str]]:
    report = str(runtime / "coverage.xml")
    return [
        ["tools/check_coverage_source_manifest.py", report],
        ["tools/check_core_coverage.py", report, "--soft-gate"],
        ["tools/check_gui_coverage.py", report],
    ]


def export_reports(runtime: Path, output: Path, *, coverage: bool, exit_code: int) -> None:
    """Copy only named reports; never preserve test fixtures, cache, or credentials."""

    if output.is_relative_to(runtime) or runtime.is_relative_to(output):
        raise ValueError("compact reports and raw runtime must be separate directories")
    report_root = runtime / "reports"
    if report_root.exists():
        info = report_root.lstat()
        if report_root.is_symlink() or getattr(info, "st_file_attributes", 0) & 0x400:
            raise ValueError("QA reports cannot traverse a link or reparse point")
    sources = [(report_root / name, name) for name in VISIBILITY_REPORTS]
    if coverage:
        sources.append((runtime / "coverage.xml", "coverage.xml"))
    records = {}
    output.mkdir()  # Exclusive; validation must already have checked the owning parent.
    for source, name in sources:
        if not source.exists():
            if exit_code == 0:
                raise ValueError(f"successful QA report missing: {name}")
            continue
        info = source.lstat()
        if not stat.S_ISREG(info.st_mode) or source.is_symlink() or getattr(info, "st_file_attributes", 0) & 0x400:
            raise ValueError(f"QA report must be a regular file: {name}")
        content = source.read_bytes()
        with (output / name).open("xb") as stream:
            stream.write(content)
        records[name] = {"bytes": len(content), "sha256": hashlib.sha256(content).hexdigest()}
    summary = {"schema": "docwen-qa-reports-v1", "exitCode": exit_code, "coverage": coverage, "reports": records}
    with (output / "summary.json").open("x", encoding="utf-8") as stream:
        stream.write(json.dumps(summary, indent=2) + "\n")

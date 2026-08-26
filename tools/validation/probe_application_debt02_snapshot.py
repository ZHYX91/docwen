"""Run the VIS-166 real-backend protective-snapshot acceptance probe."""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import os
import subprocess
import time
import zipfile
from pathlib import Path
from typing import Any

from docwen_application.preconversion.pre_converter import PreConversionResult, pre_convert

_RTF_BYTES = (
    rb"{\rtf1\ansi\deff0{\fonttbl{\f0 Calibri;}}"
    rb"\f0\fs24 DocWen DEBT-02 stable protective snapshot.\par}"
)
_TRACKED_PROCESSES = {"soffice.exe", "soffice.bin"}


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest().upper()


def _office_processes() -> dict[int, str]:
    if os.name != "nt":
        return {}
    completed = subprocess.run(
        ["tasklist.exe", "/fo", "csv", "/nh"],
        check=False,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    processes: dict[int, str] = {}
    for row in csv.reader(completed.stdout.splitlines()):
        if len(row) < 2 or row[0].casefold() not in _TRACKED_PROCESSES:
            continue
        try:
            processes[int(row[1])] = row[0]
        except ValueError:
            continue
    return processes


def _wait_for_process_cleanup(before: dict[int, str], timeout_s: float = 10.0) -> dict[int, str]:
    deadline = time.monotonic() + timeout_s
    while True:
        remaining = {pid: name for pid, name in _office_processes().items() if pid not in before}
        if not remaining or time.monotonic() >= deadline:
            return remaining
        time.sleep(0.25)


def run_probe(output_root: Path, libreoffice_program: Path) -> dict[str, Any]:
    output_root.mkdir(parents=True, exist_ok=False)
    source_dir = output_root / "source"
    staging_dir = output_root / "staging"
    source_dir.mkdir()
    staging_dir.mkdir()
    source = source_dir / "representative.rtf"
    source.write_bytes(_RTF_BYTES)
    source_hash_before = _sha256(source)

    prior_path = os.environ.get("PATH", "")
    os.environ["PATH"] = f"{libreoffice_program}{os.pathsep}{prior_path}"
    before_processes = _office_processes()
    try:
        outcome = pre_convert(
            str(source),
            "rtf",
            staging_dir=str(staging_dir),
            backend_priority=["libreoffice"],
        )
    finally:
        os.environ["PATH"] = prior_path

    remaining_processes = _wait_for_process_cleanup(before_processes)
    protected = staging_dir / "input.rtf"
    output = staging_dir / "representative.docx"
    output_members: list[str] = []
    if output.exists() and zipfile.is_zipfile(output):
        with zipfile.ZipFile(output) as archive:
            output_members = sorted(archive.namelist())

    result = {
        "backend": outcome.backend if isinstance(outcome, PreConversionResult) else None,
        "failure": None
        if isinstance(outcome, PreConversionResult)
        else {
            "message": outcome.message,
            "error_type": outcome.error_type,
            "diagnostic_code": outcome.diagnostic_code,
        },
        "libreoffice_program": str(libreoffice_program.resolve()),
        "output": {
            "exists": output.exists(),
            "is_docx_zip": zipfile.is_zipfile(output) if output.exists() else False,
            "sha256": _sha256(output) if output.exists() else None,
            "size": output.stat().st_size if output.exists() else 0,
            "has_document_xml": "word/document.xml" in output_members,
        },
        "protected_snapshot": {
            "exists": protected.exists(),
            "sha256": _sha256(protected) if protected.exists() else None,
        },
        "source": {
            "sha256_before": source_hash_before,
            "sha256_after": _sha256(source),
        },
        "new_office_processes_after_wait": remaining_processes,
    }
    result["pass"] = bool(
        isinstance(outcome, PreConversionResult)
        and outcome.backend == "LibreOffice"
        and result["source"]["sha256_before"] == result["source"]["sha256_after"]
        and result["source"]["sha256_after"] == result["protected_snapshot"]["sha256"]
        and result["output"]["is_docx_zip"]
        and result["output"]["has_document_xml"]
        and not remaining_processes
    )
    (output_root / "probe-result.json").write_text(
        json.dumps(result, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    return result


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--output-root", type=Path, required=True)
    parser.add_argument("--libreoffice-program", type=Path, required=True)
    arguments = parser.parse_args()
    result = run_probe(arguments.output_root, arguments.libreoffice_program)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if result["pass"] else 1


if __name__ == "__main__":
    raise SystemExit(main())

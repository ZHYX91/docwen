from __future__ import annotations

import json
import os
import tempfile
from dataclasses import dataclass
from importlib.util import find_spec
from pathlib import Path

from docwen.errors import ExitCode


@dataclass(slots=True)
class DoctorCheck:
    name: str
    ok: bool
    details: str | None = None


def _check_module(name: str) -> DoctorCheck:
    spec = find_spec(name)
    return DoctorCheck(
        name=f"module:{name}", ok=spec is not None, details=None if spec is not None else "not_installed"
    )


def _check_path_writable(path: str) -> DoctorCheck:
    p = Path(path)
    if not p.exists():
        return DoctorCheck(name=f"writable:{path}", ok=False, details="not_found")
    if not p.is_dir():
        return DoctorCheck(name=f"writable:{path}", ok=False, details="not_directory")
    ok = os.access(str(p), os.W_OK)
    return DoctorCheck(name=f"writable:{path}", ok=ok, details=None if ok else "not_writable")


def run_doctor(*, json_mode: bool = False) -> int:
    checks: list[DoctorCheck] = []

    checks.append(_check_module("fitz"))
    checks.append(_check_module("PIL"))
    checks.append(_check_module("rapidocr_onnxruntime"))

    checks.append(_check_path_writable(tempfile.gettempdir()))

    try:
        from docwen.config.config_manager import config_manager

        config_manager.get_output_config_block()
        checks.append(DoctorCheck(name="config:load", ok=True))
    except Exception as e:
        checks.append(DoctorCheck(name="config:load", ok=False, details=str(e) or "error"))

    ok = all(c.ok for c in checks if not c.name.startswith("module:"))

    data = {
        "ok": ok,
        "checks": [{"name": c.name, "ok": c.ok, "details": c.details} for c in checks],
    }

    if json_mode:
        from docwen.cli.executor import make_json_envelope

        output = make_json_envelope(command="doctor", success=ok, data=data)
        print(json.dumps(output, ensure_ascii=False, indent=2))
    else:
        print("doctor")
        for item in data["checks"]:
            status = "OK" if item["ok"] else "FAIL"
            details = f" ({item['details']})" if item.get("details") else ""
            print(f"- {status}: {item['name']}{details}")

    return int(ExitCode.OK if ok else ExitCode.UNKNOWN_ERROR)

"""Content checks against the actual extracted macOS release executables."""

from __future__ import annotations

import json
import os
import subprocess
from pathlib import Path
from typing import Any

from scripts.release.macos_functional_cases import identity, make_cases, make_inputs, require, verify_case


def run_cases(binary: Path, root: Path, repo: Path) -> list[dict[str, Any]]:
    root.mkdir()
    inputs = make_inputs(root / "inputs", repo)
    originals = {name: identity(path) for name, path in inputs.items()}
    environment = os.environ.copy()
    for key in ("PYTHONPATH", "PYTHONHOME", "DOCWEN_CONFIG_DIR", "DOCWEN_LOG_DIR"):
        environment.pop(key, None)
    environment.update(DOCWEN_DATA_DIR=str(root / "profile"), DOCWEN_RUNTIME_ROOT=str(root / "runtime"))
    results: list[dict[str, Any]] = []
    discovery = subprocess.run(
        [str(binary), "info", "--json"], env=environment, capture_output=True, text=True, timeout=60
    )
    require(discovery.returncode == 0, "packaged info failed")
    capabilities = {item["id"]: item for item in json.loads(discovery.stdout)["data"]["capabilities"]}
    for name in ("cli.convert", "cli.validate", "cli.number.markdown", "cli.merge", "cli.split.pdf"):
        require(
            capabilities[name]["current_platform_supported"] and capabilities[name]["available"], f"blocked: {name}"
        )
    for scope in ("core", "office"):
        for case in make_cases(inputs, root / scope, scope=scope):
            result = subprocess.run(
                [str(binary), *case.arguments, "--json"],
                cwd=root,
                env=environment,
                capture_output=True,
                text=True,
                encoding="utf-8",
                timeout=180,
            )
            require(
                result.returncode == case.expected_exit,
                f"{case.name}: exit {result.returncode}: {result.stderr[-3000:]}",
            )
            payload = json.loads(result.stdout)
            require(payload.get("success") is (case.expected_exit == 0), f"{case.name}: wrong success state")
            outputs = verify_case(case, payload, inputs)
            require({name: identity(path) for name, path in inputs.items()} == originals, "source input changed")
            results.append(
                {"name": case.name, "exitCode": result.returncode, "outputs": outputs, "inputsUnchanged": True}
            )
            print(f"macOS package: {binary.parent.name}/{case.name}: passed", flush=True)
    return results

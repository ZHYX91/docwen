#!/usr/bin/env python3
"""Post-closure CLI smoke test.

Runs key docwen CLI commands via subprocess — no controller arg, letting the
runtime initialise itself through docwen_bundle.cli_entry.main.  Produces a
JSON result document keyed by ``passed`` (critical-only) and ``results``.

Usage::

    python tools/validation/run_post_closure_cli_smoke.py
    python tools/validation/run_post_closure_cli_smoke.py --json-output result.json
"""

from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
import tempfile
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
PYTHON = sys.executable
SRC = str(REPO_ROOT / "packages" / "bundle" / "src")
SAMPLE_MD = "samples/sample.md"

_OPTIONAL_TESTS: frozenset[str] = frozenset()


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------


def _trim(s: str, *, tail: int = 4000) -> str:
    """Return the last *tail* characters of *s* so we keep JSON slices compact."""
    if len(s) <= tail:
        return s
    return "…" + s[-(tail - 1) :]


def _run_cmd(
    test_name: str,
    cli_args: list[str],
    *,
    optional: bool | None = None,
    timeout: int = 120,
    env_extra: dict[str, str] | None = None,
) -> dict:
    """Invoke ``docwen_bundle.cli_entry.main(cli_args)`` in a subprocess.

    Uses ``python -c`` because the module has no ``__main__`` block.
    CLI arguments are serialised to JSON to avoid cross-platform quoting
    issues, then decoded inside the inline script.
    """
    args_json = json.dumps(cli_args)
    script = (
        "import sys, json;"
        f"sys.path.insert(0, {SRC!r});"
        "from docwen_bundle.cli_entry import main;"
        "sys.exit(main(json.loads(sys.argv[1])))"
    )
    cmd = [PYTHON, "-c", script, args_json]
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    if env_extra:
        env.update(env_extra)

    # Build a human-readable command label (like ``docwen <args>``)
    label = "docwen " + " ".join(cli_args)

    try:
        r = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            encoding="utf-8",
            timeout=timeout,
            cwd=str(REPO_ROOT),
            env=env,
        )
        return {
            "name": test_name,
            "command": label,
            "returncode": r.returncode,
            "passed": r.returncode == 0,
            "optional": _OPTIONAL_TESTS.__contains__(test_name) if optional is None else optional,
            "stdout": _trim(r.stdout or ""),
            "stderr": _trim(r.stderr or ""),
        }
    except subprocess.TimeoutExpired:
        return {
            "name": test_name,
            "command": label,
            "returncode": -1,
            "passed": False,
            "optional": _OPTIONAL_TESTS.__contains__(test_name) if optional is None else optional,
            "stdout": "",
            "stderr": f"TIMEOUT after {timeout}s",
        }
    except Exception as exc:
        return {
            "name": test_name,
            "command": label,
            "returncode": -1,
            "passed": False,
            "optional": _OPTIONAL_TESTS.__contains__(test_name) if optional is None else optional,
            "stdout": "",
            "stderr": f"{type(exc).__name__}: {exc}",
        }


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="Post-closure CLI smoke test")
    ap.add_argument("--json-output", help="Also write JSON to this file path")
    args = ap.parse_args(argv)

    results: list[dict] = []
    with tempfile.TemporaryDirectory(prefix="docwen-post-closure-") as temp_dir:
        output_docx = str(Path(temp_dir) / "sample.docx")
        output_md = str(Path(temp_dir) / "roundtrip.md")

        results.append(_run_cmd("help", ["--help"]))
        results.append(_run_cmd("doctor", ["doctor", "--json"]))
        results.append(
            _run_cmd(
                "convert_md_to_docx",
                ["convert", SAMPLE_MD, "--to", "docx", "--output", output_docx],
            )
        )
        results.append(
            _run_cmd(
                "convert_docx_to_md",
                ["convert", output_docx, "--to", "md", "--output", output_md],
            )
        )
        results.append(_run_cmd("inspect", ["inspect", SAMPLE_MD]))
        results.append(_run_cmd("resources_formats", ["resources", "list", "formats"]))
        results.append(_run_cmd("resources_formats_json", ["resources", "list", "formats", "--json"]))

    # overall verdict — only critical tests count
    passed = all(r["passed"] for r in results if not r.get("optional"))

    output: dict = {"passed": passed, "results": results}
    json_str = json.dumps(output, ensure_ascii=False, indent=2)
    print(json_str)

    if args.json_output:
        Path(args.json_output).write_text(json_str, encoding="utf-8")

    return 0 if passed else 1


if __name__ == "__main__":
    sys.exit(main())

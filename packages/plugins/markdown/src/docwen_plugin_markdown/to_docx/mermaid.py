"""Local Mermaid CLI rendering for Markdown-to-DOCX conversion."""

from __future__ import annotations

import json
import os
import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import Any


class MermaidRenderError(RuntimeError):
    """Raised when a Mermaid diagram cannot be rendered locally."""


def _resolve_mermaid_cli() -> str:
    """Resolve a local mmdc executable without downloading anything."""

    configured = os.environ.get("DOCWEN_MERMAID_CLI", "").strip()
    if configured:
        resolved = shutil.which(configured)
        if resolved:
            return resolved
        explicit = Path(configured).expanduser()
        if explicit.is_file():
            return str(explicit)
        raise MermaidRenderError("DOCWEN_MERMAID_CLI does not point to an executable Mermaid CLI")

    resolved = shutil.which("mmdc")
    if resolved:
        return resolved
    raise MermaidRenderError(
        "Mermaid CLI (mmdc) is unavailable; install @mermaid-js/mermaid-cli or set DOCWEN_MERMAID_CLI"
    )


def _stderr_excerpt(value: str, *, limit: int = 1200) -> str:
    normalized = " ".join(value.split())
    if len(normalized) <= limit:
        return normalized
    return normalized[-limit:]


def render_mermaid_png(
    source: str,
    *,
    work_dir: str | Path,
    cancellation: Any = None,
    timeout_seconds: float = 30.0,
) -> bytes:
    """Render one Mermaid definition to PNG bytes with the local mmdc.

    The function never downloads Mermaid or contacts an online rendering
    service. securityLevel=strict is passed to Mermaid. Temporary source,
    config and output files are request-local and removed before returning.
    """

    if not source.strip():
        raise MermaidRenderError("Mermaid source is empty")
    if cancellation is not None:
        cancellation.check()

    executable = _resolve_mermaid_cli()
    root = Path(work_dir)
    root.mkdir(parents=True, exist_ok=True)
    job_dir = Path(tempfile.mkdtemp(prefix="docwen-mermaid-", dir=root))
    try:
        input_path = job_dir / "diagram.mmd"
        output_path = job_dir / "diagram.png"
        config_path = job_dir / "mermaid-config.json"
        input_path.write_text(source, encoding="utf-8", newline="")
        config_path.write_text(
            json.dumps(
                {
                    "securityLevel": "strict",
                    "theme": "default",
                    "flowchart": {"htmlLabels": False},
                },
                ensure_ascii=True,
                separators=(",", ":"),
            ),
            encoding="utf-8",
        )

        command = [
            executable,
            "--input",
            str(input_path),
            "--output",
            str(output_path),
            "--backgroundColor",
            "white",
            "--width",
            "1600",
            "--configFile",
            str(config_path),
        ]
        try:
            completed = subprocess.run(
                command,
                cwd=job_dir,
                stdin=subprocess.DEVNULL,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=timeout_seconds,
                check=False,
            )
        except subprocess.TimeoutExpired as exc:
            raise MermaidRenderError(f"Mermaid rendering exceeded {timeout_seconds:g} seconds") from exc
        except OSError as exc:
            raise MermaidRenderError(f"Mermaid CLI could not be started ({type(exc).__name__})") from exc

        if cancellation is not None:
            cancellation.check()
        if completed.returncode != 0:
            detail = _stderr_excerpt(completed.stderr or completed.stdout or "")
            suffix = f": {detail}" if detail else ""
            raise MermaidRenderError(f"Mermaid CLI exited with status {completed.returncode}{suffix}")
        if not output_path.is_file() or output_path.stat().st_size == 0:
            raise MermaidRenderError("Mermaid CLI produced no PNG output")
        return output_path.read_bytes()
    finally:
        shutil.rmtree(job_dir, ignore_errors=True)

"""Read-only discovery of the optional local Mermaid rendering installation."""

from __future__ import annotations

import json
import os
import re
import shutil
from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class MermaidRuntime:
    available: bool
    reason: str
    node: str = ""
    cli_entry: str = ""
    cli_version: str = ""
    mermaid_version: str = ""


def _package(path: Path, name: str) -> dict:
    try:
        data = json.loads((path / "package.json").read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}
    return data if isinstance(data, dict) and data.get("name") == name else {}


def _compatible(version: str) -> bool:
    match = re.fullmatch(r"11\.(\d+)\.(\d+)", version)
    return match is not None and int(match[1]) >= 16


def inspect_mermaid_runtime(cli_path: str = "") -> MermaidRuntime:
    """Inspect package versions without starting Node or a browser.

    Accept official npm global/local layouts. Browser availability is checked
    when an image conversion runs. CLI and Mermaid 11.16+ support swimlanes.
    """
    configured = cli_path.strip() or os.environ.get("DOCWEN_MERMAID_CLI", "").strip()
    command = shutil.which(configured or "mmdc")
    if command is None and configured:
        explicit = Path(configured).expanduser()
        command = str(explicit) if explicit.is_file() else None
    if command is None and not configured:
        common = [Path.home() / ".local/bin/mmdc", Path("/usr/local/bin/mmdc")]
        appdata = os.environ.get("APPDATA")
        if appdata:
            common.insert(0, Path(appdata) / "npm/mmdc.cmd")
        command = next((str(path) for path in common if path.is_file()), None)
    if command is None:
        return MermaidRuntime(False, "cli_unavailable")
    shim = Path(command).resolve()
    candidates = [shim.parent.parent, shim.parent / "node_modules/@mermaid-js/mermaid-cli"]
    candidates.append(shim.parent.parent / "@mermaid-js/mermaid-cli")
    cli_root = next((path for path in candidates if _package(path, "@mermaid-js/mermaid-cli")), None)
    if cli_root is None:
        return MermaidRuntime(False, "cli_package_unavailable")
    cli_version = str(_package(cli_root, "@mermaid-js/mermaid-cli").get("version", ""))
    core_version = ""
    for parent in (cli_root, *cli_root.parents):
        core = _package(parent / "node_modules/mermaid", "mermaid")
        if core:
            core_version = str(core.get("version", ""))
            break
    if not _compatible(cli_version) or not _compatible(core_version):
        return MermaidRuntime(False, "unsupported_version", cli_version=cli_version, mermaid_version=core_version)
    node_name = "node.exe" if os.name == "nt" else "node"
    adjacent_node = shim.parent / node_name
    node = str(adjacent_node) if adjacent_node.is_file() else shutil.which("node")
    if node is None:
        return MermaidRuntime(False, "node_unavailable", cli_version=cli_version, mermaid_version=core_version)
    entry = cli_root / "src/cli.js"
    if not entry.is_file():
        return MermaidRuntime(False, "cli_entry_unavailable", cli_version=cli_version, mermaid_version=core_version)
    return MermaidRuntime(True, "ready", node, str(entry), cli_version, core_version)

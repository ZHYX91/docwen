"""Evaluate the frozen POLICY-03=A Presentation payload oracle.

This runner executes only the current project. Historical Tk/old-PySide6
results are reused from VIS-129; their source/payload omissions are immutable
provenance, not a reason to rerun those routes.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
from pathlib import Path
from typing import Any

from docwen_core.models.file_ref import FileRef
from docwen_core.models.request import ConversionRequest, OutputPolicy
from docwen_plugin_presentation import PresentationPlugin
from docwen_runtime.engine.route_resolver import RouteResolver
from docwen_runtime.engine.task_manager import TaskManager
from docwen_runtime.output.finalizer import OutputFinalizer
from docwen_runtime.plugin_registry.registry import PluginRegistry
from docwen_runtime.workspace.manager import WorkspaceManager

SOURCES: dict[str, dict[str, Any]] = {
    "bar-chart.pptx": {
        "bytes": 44_410,
        "sha256": "79e1d218bfb2903e8dc8425a6b1997d9c1976f5a5f025bada85b0c47b5777969",
        "payloads": {
            "chart_workbook": {
                "bytes": 8_813,
                "sha256": "89673f803b955c3f553900ddfd406a80babe5a543e8f8401bfa7b2b7834cae22",
                "media_type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            }
        },
    },
    "EmbeddedAudio.pptx": {
        "bytes": 90_047,
        "sha256": "c5ae4274e2bf5504a56aef9c8d7c5d2381ece69c9bf68f7749ad5eae3e675edb",
        "payloads": {
            "audio": {
                "bytes": 52_079,
                "sha256": "0244590f2b4bcb62352b574e78bea940e8d89cfa69823b5208ef4c43e0abcb44",
                "media_type": "audio/mpeg",
            },
            "audio_poster": {
                "bytes": 4_717,
                "sha256": "b0151c2c2e3cf64bc37a7bb9d8b8b98d4c4fccf7b6af4c08c4f847a79f9db0da",
                "media_type": "image/png",
            },
        },
    },
    "EmbeddedVideo.pptx": {
        "bytes": 201_418,
        "sha256": "7940e3b1a339db11f00b65399a2fe77e0e85a5da3a30ac8d6c8a0a77527b2ab2",
        "payloads": {
            "video": {
                "bytes": 101_799,
                "sha256": "21c3b5d779abe3bc2ee886a6d2455202800537fa31fe367d11563da16cbf8040",
                "media_type": "video/mp4",
            },
            "video_poster": {
                "bytes": 65_215,
                "sha256": "f5516c6cae484df63ce03db77fb69b778660916b9207de5a4e04aa5e3b72908d",
                "media_type": "image/png",
            },
        },
    },
}


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _office_processes() -> list[dict[str, Any]]:
    try:
        import psutil
    except ImportError:
        return []
    result: list[dict[str, Any]] = []
    for process in psutil.process_iter(["pid", "name"]):
        try:
            name = str(process.info.get("name") or "")
            if name.lower() in {"powerpnt.exe", "soffice.exe"}:
                result.append({"pid": process.pid, "name": name})
        except (psutil.AccessDenied, psutil.NoSuchProcess):
            continue
    return sorted(result, key=lambda item: (item["name"], item["pid"]))


def _markdown_targets(markdown: str) -> list[str]:
    return re.findall(r"!?\[[^\]]*\]\(([^)]+)\)", markdown)


def _manager(workspace_root: Path) -> tuple[TaskManager, WorkspaceManager]:
    registry = PluginRegistry()
    registry.register(PresentationPlugin())
    workspace = WorkspaceManager(root_dir=str(workspace_root))
    manager = TaskManager(
        registry,
        RouteResolver(registry),
        workspace,
        OutputFinalizer(),
    )
    return manager, workspace


def _run_case(
    manager: TaskManager,
    source: Path,
    output_dir: Path,
    *,
    request_id: str,
) -> dict[str, Any]:
    output_dir.mkdir(parents=True, exist_ok=True)
    result = manager.execute_single(
        ConversionRequest(
            request_id=request_id,
            input_refs=[
                FileRef(
                    path=str(source),
                    format="pptx",
                    category="presentation",
                    size_bytes=source.stat().st_size,
                )
            ],
            target_format="md",
            output_policy=OutputPolicy(output_dir=str(output_dir)),
            options={
                "export_notes": True,
                "to_md_keep_images": True,
                "to_md_enable_ocr": False,
                "image_mode": "file",
                "image_link_style": "markdown_embed",
                "locale": "zh_CN",
                "yaml_key_labels": {"title": "标题"},
            },
        )
    )
    artifacts: list[dict[str, Any]] = []
    for artifact in result.artifacts:
        path = Path(artifact.staging_path)
        artifacts.append(
            {
                "kind": artifact.kind,
                "suggested_name": artifact.suggested_name,
                "final_name": path.name,
                "media_type": artifact.media_type,
                "metadata": artifact.metadata,
                "bytes": path.stat().st_size,
                "sha256": _sha256(path),
                "path": str(path),
            }
        )
    primary = next((item for item in artifacts if item["kind"] == "primary"), None)
    markdown = Path(primary["path"]).read_text(encoding="utf-8") if primary else ""
    return {
        "success": result.success,
        "diagnostics": [diagnostic.to_dict() for diagnostic in result.diagnostics],
        "error": result.error.to_dict() if result.error else None,
        "artifacts": artifacts,
        "markdown": markdown,
        "targets": _markdown_targets(markdown),
    }


def _evaluate_case(name: str, case: dict[str, Any]) -> dict[str, Any]:
    expected = SOURCES[name]
    artifacts = case["artifacts"]
    payload_checks: dict[str, Any] = {}
    for payload_name, payload_expected in expected["payloads"].items():
        matches = [artifact for artifact in artifacts if artifact["metadata"].get("payload") == payload_name]
        payload_checks[payload_name] = {
            "count": len(matches),
            "exact": bool(matches)
            and matches[0]["bytes"] == payload_expected["bytes"]
            and matches[0]["sha256"] == payload_expected["sha256"]
            and matches[0]["media_type"] == payload_expected["media_type"],
            "linked": bool(matches) and matches[0]["final_name"] in case["targets"],
        }

    diagnostic_codes = [item["code"] for item in case["diagnostics"]]
    chart_rows = (
        "| 1st Qtr | 8.200000000000001 |",
        "| 2nd Qtr | 3.2 |",
        "| 3rd Qtr | 1.4 |",
        "| 4th Qtr | 1.2 |",
    )
    chart_projection = None
    if name == "bar-chart.pptx":
        positions = [case["markdown"].find(row) for row in chart_rows]
        chart_projection = {
            "tokens_present": "### Chart: Sales" in case["markdown"] and all(position >= 0 for position in positions),
            "ordered": positions == sorted(positions),
            "table_contiguous": "| Category | Sales |\n| --- | --- |" in case["markdown"],
            "snapshot_warning_count": diagnostic_codes.count("PPTX-CHART-SNAPSHOT-UNAVAILABLE"),
            "snapshot_artifact_count": sum(
                1 for artifact in artifacts if artifact["metadata"].get("payload") == "chart_snapshot"
            ),
        }

    passed = (
        case["success"] is True
        and case["error"] is None
        and all(check["exact"] and check["linked"] for check in payload_checks.values())
        and (
            chart_projection is None
            or (
                chart_projection["tokens_present"]
                and chart_projection["ordered"]
                and chart_projection["table_contiguous"]
                and chart_projection["snapshot_warning_count"] == 1
                and chart_projection["snapshot_artifact_count"] == 0
            )
        )
    )
    return {
        "passed": passed,
        "payloads": payload_checks,
        "chart_projection": chart_projection,
        "diagnostic_codes": diagnostic_codes,
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source-root", type=Path, required=True)
    parser.add_argument("--output-root", type=Path, required=True)
    args = parser.parse_args()

    output_root = args.output_root.resolve()
    if output_root.exists() and any(output_root.iterdir()):
        raise SystemExit(f"output root must be empty: {output_root}")
    output_root.mkdir(parents=True, exist_ok=True)
    workspace_root = output_root / "workspace"
    manager, workspace = _manager(workspace_root)
    source_before = {
        name: {
            "bytes": (args.source_root / name).stat().st_size,
            "sha256": _sha256(args.source_root / name),
        }
        for name in SOURCES
    }
    processes_before = _office_processes()
    cases: dict[str, Any] = {}
    try:
        for name in SOURCES:
            cases[name] = _run_case(
                manager,
                args.source_root / name,
                output_root / "outputs" / Path(name).stem,
                request_id=f"vis204-{Path(name).stem}",
            )

        batch_root = output_root / "same-basename-inputs"
        batch_output = output_root / "same-basename-output"
        batch_cases: dict[str, Any] = {}
        for index, name in enumerate(SOURCES, 1):
            source_copy = batch_root / str(index) / "same.pptx"
            source_copy.parent.mkdir(parents=True, exist_ok=True)
            shutil.copyfile(args.source_root / name, source_copy)
            batch_cases[name] = _run_case(
                manager,
                source_copy,
                batch_output,
                request_id=f"vis204-same-{index}",
            )
    finally:
        workspace.cleanup_all()
        shutil.rmtree(workspace_root, ignore_errors=True)

    source_after = {
        name: {
            "bytes": (args.source_root / name).stat().st_size,
            "sha256": _sha256(args.source_root / name),
        }
        for name in SOURCES
    }
    source_expected = {
        name: before["bytes"] == SOURCES[name]["bytes"] and before["sha256"] == SOURCES[name]["sha256"]
        for name, before in source_before.items()
    }
    evaluations = {name: _evaluate_case(name, case) for name, case in cases.items()}
    batch_evaluations = {name: _evaluate_case(name, case) for name, case in batch_cases.items()}
    batch_primary_names = [
        next(artifact["final_name"] for artifact in batch_cases[name]["artifacts"] if artifact["kind"] == "primary")
        for name in SOURCES
    ]
    batch_reachable = all(
        all((batch_output / target).is_file() for target in batch_cases[name]["targets"]) for name in SOURCES
    )
    processes_after = _office_processes()
    result = {
        "stage": "VIS-2026-07-23-204",
        "policy": "POLICY-03=A",
        "source_root": str(args.source_root.resolve()),
        "output_root": str(output_root),
        "source_before": source_before,
        "source_after": source_after,
        "source_expected": source_expected,
        "source_immutable": source_before == source_after,
        "cases": cases,
        "evaluations": evaluations,
        "same_basename": {
            "cases": batch_cases,
            "evaluations": batch_evaluations,
            "primary_names": batch_primary_names,
            "primary_names_unique": len(set(batch_primary_names)) == len(batch_primary_names),
            "all_targets_reachable": batch_reachable,
        },
        "office_processes": {
            "before": processes_before,
            "after": processes_after,
            "added": [item for item in processes_after if item not in processes_before],
        },
    }
    result["passed"] = (
        all(source_expected.values())
        and result["source_immutable"]
        and all(item["passed"] for item in evaluations.values())
        and all(item["passed"] for item in batch_evaluations.values())
        and result["same_basename"]["primary_names_unique"]
        and batch_reachable
        and not result["office_processes"]["added"]
    )
    result_path = output_root / "result.json"
    result_path.write_text(
        json.dumps(result, ensure_ascii=False, indent=2) + os.linesep,
        encoding="utf-8",
    )
    print(json.dumps({"passed": result["passed"], "result": str(result_path)}))
    if not result["passed"]:
        raise SystemExit(1)


if __name__ == "__main__":
    main()

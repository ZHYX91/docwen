"""Capture FA-08 image final-artifact parity evidence.

The stage contract and binary corpus live outside all three repositories.  The
orchestrator validates that frozen contract, executes five production routes
for four images in each project, and records format/resource/render/OCR
projections without writing to either reference repository.
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import io
import json
import os
import re
import shutil
import subprocess
import sys
import time
from pathlib import Path
from typing import Any

from PIL import Image, ImageChops, ImageOps, ImageStat

PROJECTS = ("docwen-ref-tk", "docwen-ref-pyside6", "docwen-current")
ROUTES = ("jpg", "webp", "pdf-original", "markdown-file", "markdown-image-md")
IMAGE_SUFFIXES = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tif", ".tiff", ".webp", ".heic", ".heif"}
MODEL_NAMES = (
    "ch_PP-OCRv4_det_infer.onnx",
    "ch_PP-OCRv4_rec_infer.onnx",
    "ch_ppocr_mobile_v2.0_cls_infer.onnx",
)
MINIMUM_DISPLAY_ORIENTED_RGB_SIMILARITY = 0.95


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest().lower()


def _write_json(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def _print_json(value: Any) -> None:
    sys.stdout.buffer.write((json.dumps(value, ensure_ascii=False, indent=2) + "\n").encode("utf-8"))


def _actual_format(path: Path) -> str:
    with Image.open(path) as image:
        value = (image.format or path.suffix.lstrip(".")).lower()
    return "jpg" if value in {"jpeg", "mpo"} else value


def _orientation(image: Image.Image) -> int | None:
    try:
        value = image.getexif().get(274)
        return int(value) if value is not None else None
    except Exception:
        return None


def _image_projection(path: Path) -> dict[str, Any]:
    with Image.open(path) as image:
        exif = image.getexif()
        selected = {str(tag): exif.get(tag) for tag in (271, 272, 274, 305, 306, 36867) if exif.get(tag) is not None}
        display = ImageOps.exif_transpose(image.copy())
        projection = {
            "sha256": _sha256(path),
            "bytes": path.stat().st_size,
            "format": image.format,
            "mode": image.mode,
            "raw_size": list(image.size),
            "display_size": list(display.size),
            "frames": int(getattr(image, "n_frames", 1)),
            "orientation": _orientation(image),
            "exif_tag_count": len(exif),
            "selected_exif": selected,
            "icc_bytes": len(image.info.get("icc_profile", b"") or b""),
            "duration_ms": image.info.get("duration"),
            "loop": image.info.get("loop"),
        }
        if int(getattr(image, "n_frames", 1)) > 1:
            frame_means: list[list[float]] = []
            for index in range(int(getattr(image, "n_frames", 1))):
                image.seek(index)
                stat = ImageStat.Stat(image.convert("RGB"))
                frame_means.append([round(value, 3) for value in stat.mean])
            projection["frame_rgb_means"] = frame_means
        display.close()
        return projection


def _display_rgb(path: Path, frame: int = 0) -> Image.Image:
    with Image.open(path) as image:
        image.seek(frame)
        return ImageOps.exif_transpose(image.copy()).convert("RGB")


def _rgb_similarity(expected: Image.Image, actual: Image.Image) -> dict[str, Any]:
    same_dimensions = expected.size == actual.size
    if not same_dimensions:
        return {
            "same_dimensions": False,
            "expected_size": list(expected.size),
            "actual_size": list(actual.size),
            "score": None,
        }
    difference = ImageChops.difference(expected, actual)
    mean = sum(ImageStat.Stat(difference).mean) / 3.0
    return {
        "same_dimensions": True,
        "expected_size": list(expected.size),
        "actual_size": list(actual.size),
        "mean_absolute_rgb_channel_error": round(mean, 6),
        "score": round(1.0 - mean / 255.0, 9),
    }


def _metadata_fidelity(source: dict[str, Any], output: dict[str, Any], similarity: dict[str, Any]) -> dict[str, Any]:
    source_orientation = source.get("orientation")
    output_orientation = output.get("orientation")
    orientation_met = source_orientation in (None, 1) or output_orientation == source_orientation
    if source_orientation not in (None, 1) and output_orientation is None:
        orientation_met = bool(
            similarity.get("same_dimensions") and similarity.get("score", 0) >= MINIMUM_DISPLAY_ORIENTED_RGB_SIMILARITY
        )
    source_selected = {key: value for key, value in source.get("selected_exif", {}).items() if key != "274"}
    output_selected = output.get("selected_exif", {})
    selected_exif_met = all(output_selected.get(key) == value for key, value in source_selected.items())
    icc_met = not source.get("icc_bytes") or output.get("icc_bytes") == source.get("icc_bytes")
    return {
        "orientation_met": orientation_met,
        "selected_exif_met": selected_exif_met,
        "icc_met": icc_met,
        "source_orientation": source_orientation,
        "output_orientation": output_orientation,
        "source_selected_exif": source_selected,
        "output_selected_exif": output_selected,
        "source_icc_bytes": source.get("icc_bytes"),
        "output_icc_bytes": output.get("icc_bytes"),
        "all_met": orientation_met and selected_exif_met and icc_met,
    }


def _pdf_projection(path: Path, source: Path) -> dict[str, Any]:
    import fitz

    with Image.open(source) as source_image:
        source_frames = int(getattr(source_image, "n_frames", 1))
    with fitz.open(path) as document:
        pages: list[dict[str, Any]] = []
        for index in range(document.page_count):
            page = document[index]
            rect = page.rect
            images = page.get_images(full=True)
            source_rgb = _display_rgb(source, index) if index < source_frames else None
            embedded: list[dict[str, Any]] = []
            for info in images:
                extracted = document.extract_image(info[0])
                embedded_item = {
                    "xref": info[0],
                    "width": info[2],
                    "height": info[3],
                    "colorspace": info[5],
                    "extension": extracted.get("ext"),
                    "bytes": len(extracted.get("image", b"")),
                    "sha256": hashlib.sha256(extracted.get("image", b"")).hexdigest(),
                }
                if source_rgb is not None:
                    try:
                        with Image.open(io.BytesIO(extracted.get("image", b""))) as embedded_image:
                            embedded_rgb = ImageOps.exif_transpose(embedded_image.copy()).convert("RGB")
                        embedded_item["decoded_similarity_to_source_frame"] = _rgb_similarity(source_rgb, embedded_rgb)
                        embedded_rgb.close()
                    except Exception as exc:
                        embedded_item["decoded_projection_error"] = str(exc)
                embedded.append(embedded_item)
            page_item: dict[str, Any] = {
                "index": index,
                "rect": [round(rect.x0, 6), round(rect.y0, 6), round(rect.x1, 6), round(rect.y1, 6)],
                "rotation": page.rotation,
                "embedded_images": embedded,
            }
            if index < source_frames and rect.width and rect.height:
                assert source_rgb is not None
                zoom = source_rgb.width / rect.width
                pixmap = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
                rendered = Image.frombytes("RGB", (pixmap.width, pixmap.height), pixmap.samples)
                page_item["render"] = _rgb_similarity(source_rgb, rendered)
                page_item["source_aspect"] = round(source_rgb.width / source_rgb.height, 9)
                page_item["page_aspect"] = round(rect.width / rect.height, 9)
                page_item["aspect_delta"] = round(
                    abs(source_rgb.width / source_rgb.height - rect.width / rect.height), 9
                )
                rendered.close()
            if source_rgb is not None:
                source_rgb.close()
            pages.append(page_item)
    return {
        "sha256": _sha256(path),
        "bytes": path.stat().st_size,
        "source_frame_count": source_frames,
        "page_count": len(pages),
        "pages": pages,
    }


def _strip_yaml(text: str) -> str:
    normalized = text.replace("\r\n", "\n").replace("\r", "\n")
    if normalized.startswith("---\n"):
        end = normalized.find("\n---\n", 4)
        if end >= 0:
            normalized = normalized[end + 5 :]
    return normalized


def _ocr_projection(md_files: list[Path], primary: Path | None) -> dict[str, Any]:
    sidecars = [path for path in md_files if primary is None or path.resolve() != primary.resolve()]
    lines: list[str] = []
    for sidecar in sidecars:
        for raw_line in _strip_yaml(sidecar.read_text(encoding="utf-8", errors="replace")).splitlines():
            line = raw_line.strip()
            if re.fullmatch(r"!?\[\[[^]]+\]\]", line) or re.fullmatch(r"!?\[[^]]*\]\([^)]+\)", line):
                continue
            if line.startswith(">"):
                line = line[1:].strip()
            if not line or (line.startswith("**") and line.endswith("**")):
                continue
            lines.append(line)
    return {
        "sidecars": [str(path) for path in sidecars],
        "normalized_text": "\n".join(lines).strip(),
        "normalized_characters": len("\n".join(lines).strip()),
    }


def _markdown_links(path: Path) -> list[str]:
    text = path.read_text(encoding="utf-8", errors="replace")
    targets: list[str] = []
    targets.extend(match.strip() for match in re.findall(r"!?\[\[([^\]|#]+)", text))
    targets.extend(match.strip() for match in re.findall(r"!?\[[^\]]*\]\(([^)]+)\)", text))
    return targets


def _markdown_projection(final_dir: Path, source: Path, primary_hint: str | None) -> dict[str, Any]:
    files = sorted(path for path in final_dir.rglob("*") if path.is_file())
    md_files = [path for path in files if path.suffix.lower() == ".md"]
    primary = None
    if primary_hint:
        hinted = Path(primary_hint).resolve()
        if hinted.exists() and hinted.suffix.lower() == ".md":
            primary = hinted
    if primary is None and md_files:
        primary = min(md_files, key=lambda path: ("_ocr" in path.stem.lower(), len(path.name), path.name))
    links: dict[str, Any] = {}
    link_style_counts = {"wiki_embed": 0, "wiki_link": 0, "markdown_embed": 0, "markdown_link": 0}
    all_reachable = True
    no_absolute_links = True
    for md_file in md_files:
        content = md_file.read_text(encoding="utf-8", errors="replace")
        link_style_counts["wiki_embed"] += len(re.findall(r"!\[\[", content))
        link_style_counts["wiki_link"] += len(re.findall(r"(?<!!)\[\[", content))
        link_style_counts["markdown_embed"] += len(re.findall(r"!\[[^]]*\]\(", content))
        link_style_counts["markdown_link"] += len(re.findall(r"(?<!!)\[[^]]*\]\(", content))
        projected: list[dict[str, Any]] = []
        for target in _markdown_links(md_file):
            link_path = Path(target)
            is_absolute = link_path.is_absolute()
            reachable = False if is_absolute else (md_file.parent / link_path).resolve().exists()
            all_reachable = all_reachable and reachable
            no_absolute_links = no_absolute_links and not is_absolute
            projected.append({"target": target, "absolute": is_absolute, "reachable": reachable})
        links[str(md_file)] = projected
    retained_images = [path for path in files if path.suffix.lower() in IMAGE_SUFFIXES]
    retained = [
        {
            "path": str(path),
            "sha256": _sha256(path),
            "source_byte_identical": _sha256(path) == _sha256(source),
            "projection": _image_projection(path),
        }
        for path in retained_images
    ]
    return {
        "primary": str(primary) if primary else None,
        "markdown_files": [str(path) for path in md_files],
        "retained_images": retained,
        "all_links_reachable": all_reachable,
        "no_absolute_links": no_absolute_links,
        "link_style_counts": link_style_counts,
        "case_owned_naming": all(source.stem in path.name for path in files),
        "links": links,
        "ocr": _ocr_projection(md_files, primary),
    }


def _task_processes() -> list[dict[str, str]]:
    completed = subprocess.run(
        ["tasklist", "/fo", "csv", "/nh"],
        check=False,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    tracked = {"python.exe", "pythonw.exe", "docwen.exe", "soffice.exe", "soffice.bin"}
    return [
        {"name": row[0], "pid": row[1]}
        for row in csv.reader(completed.stdout.splitlines())
        if len(row) >= 2 and row[0].lower() in tracked
    ]


def _worker(args: argparse.Namespace) -> int:
    source = Path(args.worker_source).resolve()
    final_dir = Path(args.worker_final).resolve()
    result_path = Path(args.worker_result).resolve()
    final_dir.mkdir(parents=True, exist_ok=True)
    actual = _actual_format(source)

    from docwen.utils import ocr_utils, workspace_manager

    ocr_utils.get_configured_ocr_language = lambda: "chinese"
    ocr_utils.reset_ocr()

    if args.worker_project == "tk":
        workspace_manager.get_output_directory = lambda *_args, **_kwargs: str(final_dir)
        if args.worker_route in {"jpg", "webp"}:
            from docwen.services.strategies.image.format_conversion import ImageFormatConversionStrategy

            result = ImageFormatConversionStrategy().execute(
                str(source),
                options={"actual_format": actual, "target_format": args.worker_route, "compress_mode": "lossless"},
            )
        elif args.worker_route == "pdf-original":
            from docwen.services.strategies.image.to_pdf import ImageToPdfStrategy

            result = ImageToPdfStrategy().execute(
                str(source), options={"actual_format": actual, "quality_mode": "original"}
            )
        else:
            from docwen.config.config_manager import config_manager
            from docwen.services.strategies.image.to_markdown import ImageToMarkdownStrategy

            config_manager.get_markdown_link_style_settings = lambda: {
                "image_link_style": "wiki_embed",
                "md_file_link_style": "wiki_embed",
            }
            result = ImageToMarkdownStrategy().execute(
                str(source),
                options={
                    "actual_format": actual,
                    "extract_image": True,
                    "extract_ocr": args.worker_route == "markdown-image-md",
                    "to_md_image_extraction_mode": "file",
                    "to_md_ocr_placement_mode": "image_md",
                },
            )
    else:
        from docwen.services.options import CommonOptions, ImageOpOptions, ToMarkdownOptions
        from docwen.utils.workspace_manager import OutputPersistencePolicy

        common = CommonOptions(actual_format=actual, headless=True)

        def set_policy(strategy: Any) -> None:
            strategy._build_output_policy = lambda file_path, **_kwargs: OutputPersistencePolicy.from_input(
                file_path, custom_output_dir=str(final_dir), keep_intermediates=False
            )

        if args.worker_route in {"jpg", "webp"}:
            from docwen.services.strategies.image.format_conversion import ImageFormatConversionStrategy

            strategy = ImageFormatConversionStrategy()
            strategy.target_format = args.worker_route
            set_policy(strategy)
            result = strategy.execute(str(source), common=common, image_op=ImageOpOptions(compress_mode="lossless"))
        elif args.worker_route == "pdf-original":
            from docwen.services.strategies.image.to_pdf import ImageToPdfStrategy

            strategy = ImageToPdfStrategy()
            set_policy(strategy)
            result = strategy.execute(str(source), common=common, image_op=ImageOpOptions(quality_mode="original"))
        else:
            from docwen.config.config_manager import config_manager
            from docwen.services.strategies.image.to_markdown import ImageToMarkdownStrategy

            config_manager.get_markdown_link_style_settings = lambda: {
                "image_link_style": "wiki_embed",
                "md_file_link_style": "wiki_embed",
            }
            strategy = ImageToMarkdownStrategy()
            set_policy(strategy)
            result = strategy.execute(
                str(source),
                common=common,
                md=ToMarkdownOptions(
                    extract_image=True,
                    extract_ocr=args.worker_route == "markdown-image-md",
                    to_md_image_extraction_mode="file",
                    to_md_ocr_placement_mode="image_md",
                ),
            )

    payload = {
        "success": bool(result.success),
        "message": str(result.message),
        "output_path": str(result.output_path) if result.output_path else None,
        "error_code": getattr(result, "error_code", None),
        "details": str(getattr(result, "details", "") or ""),
    }
    _write_json(result_path, payload)
    return 0 if result.success else 1


def _json_from_stdout(stdout: str) -> dict[str, Any] | None:
    decoder = json.JSONDecoder()
    for index, char in enumerate(stdout):
        if char != "{":
            continue
        try:
            value, _ = decoder.raw_decode(stdout[index:])
        except json.JSONDecodeError:
            continue
        if isinstance(value, dict):
            return value
    return None


def _run_route(
    *, project: str, root: Path, executable: Path, case: str, route: str, frozen_source: Path, evidence_root: Path
) -> dict[str, Any]:
    route_dir = evidence_root / "runs" / project / case.lower() / route
    if route_dir.exists():
        if not route_dir.resolve().is_relative_to(evidence_root.resolve()):
            raise RuntimeError(f"refusing to replace evidence outside root: {route_dir}")
        shutil.rmtree(route_dir)
    final_dir = route_dir / "final"
    final_dir.mkdir(parents=True)
    source = route_dir / frozen_source.name
    shutil.copyfile(frozen_source, source)
    worker_result = route_dir / "worker-result.json"
    before_processes = _task_processes()

    if project == "docwen-current":
        target = route.split("-", 1)[0]
        command = [str(executable), "run", str(source), "--to", target, "--output", str(final_dir), "--json", "--quiet"]
        if route == "pdf-original":
            command.extend(["--quality-mode", "original"])
        elif route.startswith("markdown"):
            command.extend(["--extract-img", "--image-mode", "file", "--image-link-style", "wiki_embed"])
            if route == "markdown-image-md":
                command.extend(["--ocr", "--ocr-language", "chinese", "--ocr-placement", "image_md", "--lang", "zh_CN"])
    else:
        command = [
            str(root / ".venv" / "Scripts" / "python.exe"),
            str(Path(__file__).resolve()),
            "--worker-project",
            "tk" if project == "docwen-ref-tk" else "pyside",
            "--worker-route",
            route,
            "--worker-source",
            str(source),
            "--worker-final",
            str(final_dir),
            "--worker-result",
            str(worker_result),
        ]

    started = time.monotonic()
    completed = subprocess.run(command, cwd=root, capture_output=True, timeout=600)
    elapsed = round(time.monotonic() - started, 3)
    stdout = completed.stdout.decode("utf-8", errors="replace")
    stderr = completed.stderr.decode("utf-8", errors="replace")
    immediate_processes = _task_processes()
    after_processes = immediate_processes
    deadline = time.monotonic() + 10.0
    while any(item not in before_processes for item in after_processes) and time.monotonic() < deadline:
        time.sleep(0.5)
        after_processes = _task_processes()
    worker_payload = json.loads(worker_result.read_text(encoding="utf-8")) if worker_result.exists() else None
    current_payload = _json_from_stdout(stdout) if project == "docwen-current" else None
    result_success = (
        bool(worker_payload and worker_payload.get("success"))
        if worker_payload is not None
        else bool(completed.returncode == 0 and current_payload and current_payload.get("success"))
    )
    primary_hint = worker_payload.get("output_path") if worker_payload else None
    if primary_hint is None and current_payload:
        primary_hint = (current_payload.get("data") or {}).get("output_file")
    files = sorted(path for path in final_dir.rglob("*") if path.is_file())
    artifact_inventory = [
        {
            "path": str(path),
            "relative_path": path.relative_to(route_dir).as_posix(),
            "suffix": path.suffix.lower(),
            "bytes": path.stat().st_size,
            "sha256": _sha256(path),
        }
        for path in files
    ]
    projection: dict[str, Any] | None = None
    if route in {"jpg", "webp"}:
        expected_suffix = ".jpg" if route == "jpg" else ".webp"
        candidates = [path for path in files if path.suffix.lower() == expected_suffix]
        if len(candidates) == 1:
            source_rgb = _display_rgb(source)
            output_rgb = _display_rgb(candidates[0])
            projection = {
                "artifact": _image_projection(candidates[0]),
                "display_similarity": _rgb_similarity(source_rgb, output_rgb),
            }
            projection["metadata_fidelity"] = _metadata_fidelity(
                _image_projection(source), projection["artifact"], projection["display_similarity"]
            )
            source_rgb.close()
            output_rgb.close()
    elif route == "pdf-original":
        candidates = [path for path in files if path.suffix.lower() == ".pdf"]
        if len(candidates) == 1:
            projection = _pdf_projection(candidates[0], source)
    else:
        projection = _markdown_projection(final_dir, source, primary_hint)

    return {
        "command": command,
        "returncode": completed.returncode,
        "elapsed_seconds": elapsed,
        "stdout": stdout,
        "stderr": stderr,
        "worker_result": worker_payload,
        "current_json": current_payload,
        "success": result_success,
        "artifact_inventory": artifact_inventory,
        "final_placement": all(path.resolve().is_relative_to(route_dir.resolve()) for path in files),
        "projection": projection,
        "processes": {
            "before": before_processes,
            "immediate": immediate_processes,
            "after": after_processes,
            "residue_added": [item for item in after_processes if item not in before_processes],
        },
    }


def _model_manifest(roots: dict[str, Path]) -> dict[str, Any]:
    return {
        project: {
            name: {
                "path": str(root / "models" / "rapidocr" / name),
                "exists": (root / "models" / "rapidocr" / name).is_file(),
                "bytes": (root / "models" / "rapidocr" / name).stat().st_size
                if (root / "models" / "rapidocr" / name).is_file()
                else None,
                "sha256": _sha256(root / "models" / "rapidocr" / name)
                if (root / "models" / "rapidocr" / name).is_file()
                else None,
            }
            for name in MODEL_NAMES
        }
        for project, root in roots.items()
    }


def _validate_contract(contract_path: Path) -> tuple[dict[str, Any], dict[str, Path]]:
    contract = json.loads(contract_path.read_text(encoding="utf-8"))
    if contract.get("contract_id") != "CONTRACT-2026-07-22-001/FA-08":
        raise RuntimeError("unexpected FA-08 contract id")
    if contract.get("expected_route_execution_count") != 60:
        raise RuntimeError("FA-08 execution count is not frozen at 60")
    if [item.get("route_id") for item in contract.get("routes_per_input_per_project", [])] != list(ROUTES):
        raise RuntimeError("FA-08 route set differs from frozen route set")
    thresholds = contract.get("quality_thresholds_frozen_before_execution", {})
    if thresholds.get("minimum_display_oriented_rgb_similarity") != MINIMUM_DISPLAY_ORIENTED_RGB_SIMILARITY:
        raise RuntimeError("FA-08 display-oriented RGB threshold differs from the frozen probe threshold")
    inputs: dict[str, Path] = {}
    for item in contract.get("inputs", []):
        path = (contract_path.parent / item["path"]).resolve()
        if not path.is_file() or _sha256(path) != item["sha256"].lower() or path.stat().st_size != item["bytes"]:
            raise RuntimeError(f"frozen input validation failed: {path}")
        inputs[item["case_id"].split("-")[-1]] = path
    if set(inputs) != {"B1", "B2", "N1", "N2"}:
        raise RuntimeError(f"unexpected FA-08 cases: {sorted(inputs)}")
    return contract, inputs


def _route_ok(route: str, value: dict[str, Any]) -> bool:
    if not value["success"] or not value["final_placement"] or value["projection"] is None:
        return False
    if value["processes"]["residue_added"]:
        return False
    projection = value["projection"]
    if route in {"jpg", "webp"}:
        similarity = projection["display_similarity"]
        return bool(
            similarity["same_dimensions"]
            and similarity["score"] is not None
            and similarity["score"] >= MINIMUM_DISPLAY_ORIENTED_RGB_SIMILARITY
            and projection["metadata_fidelity"]["all_met"]
        )
    if route == "pdf-original":
        if projection["page_count"] != projection["source_frame_count"] or not projection["pages"]:
            return False
        return all(
            page.get("render", {}).get("same_dimensions")
            and page["render"].get("score") is not None
            and page["render"]["score"] >= MINIMUM_DISPLAY_ORIENTED_RGB_SIMILARITY
            for page in projection["pages"]
        )
    retained = projection["retained_images"]
    return bool(
        projection["primary"]
        and projection["all_links_reachable"]
        and projection["no_absolute_links"]
        and projection["case_owned_naming"]
        and projection["link_style_counts"]["wiki_embed"] > 0
        and not projection["link_style_counts"]["wiki_link"]
        and not projection["link_style_counts"]["markdown_embed"]
        and not projection["link_style_counts"]["markdown_link"]
        and retained
        and any(item["source_byte_identical"] for item in retained)
    )


def _reproject_route(route: str, value: dict[str, Any], source: Path) -> None:
    files = [Path(item["path"]).resolve() for item in value["artifact_inventory"]]
    if route in {"jpg", "webp"}:
        expected_suffix = ".jpg" if route == "jpg" else ".webp"
        candidates = [path for path in files if path.suffix.lower() == expected_suffix and path.is_file()]
        value["projection"] = None
        if len(candidates) == 1:
            source_rgb = _display_rgb(source)
            output_rgb = _display_rgb(candidates[0])
            value["projection"] = {
                "artifact": _image_projection(candidates[0]),
                "display_similarity": _rgb_similarity(source_rgb, output_rgb),
            }
            value["projection"]["metadata_fidelity"] = _metadata_fidelity(
                _image_projection(source),
                value["projection"]["artifact"],
                value["projection"]["display_similarity"],
            )
            source_rgb.close()
            output_rgb.close()
        return
    if route == "pdf-original":
        candidates = [path for path in files if path.suffix.lower() == ".pdf" and path.is_file()]
        value["projection"] = _pdf_projection(candidates[0], source) if len(candidates) == 1 else None
        return
    primary_hint = (value.get("worker_result") or {}).get("output_path")
    if primary_hint is None:
        primary_hint = ((value.get("current_json") or {}).get("data") or {}).get("output_file")
    final_dirs = {path.parent for path in files}
    common = Path(os.path.commonpath([str(path) for path in final_dirs])) if final_dirs else None
    value["projection"] = _markdown_projection(common, source, primary_hint) if common else None


def _route_failure_reasons(route: str, value: dict[str, Any]) -> list[str]:
    reasons: list[str] = []
    if not value.get("success"):
        reasons.append("route_failed")
    if not value.get("final_placement"):
        reasons.append("final_placement")
    if value.get("processes", {}).get("residue_added"):
        reasons.append("process_residue")
    projection = value.get("projection")
    if projection is None:
        reasons.append("missing_projection")
        return reasons
    if route in {"jpg", "webp"}:
        similarity = projection["display_similarity"]
        if (
            not similarity.get("same_dimensions")
            or similarity.get("score") is None
            or similarity["score"] < MINIMUM_DISPLAY_ORIENTED_RGB_SIMILARITY
        ):
            reasons.append("display_oriented_rgb_fidelity")
        metadata = projection["metadata_fidelity"]
        if not metadata["orientation_met"]:
            reasons.append("orientation_fidelity")
        if not metadata["selected_exif_met"]:
            reasons.append("selected_exif_fidelity")
        if not metadata["icc_met"]:
            reasons.append("icc_fidelity")
    elif route == "pdf-original":
        if projection["page_count"] != projection["source_frame_count"]:
            reasons.append("pdf_frame_page_count")
        for page in projection["pages"]:
            render = page.get("render", {})
            if (
                not render.get("same_dimensions")
                or render.get("score") is None
                or render["score"] < MINIMUM_DISPLAY_ORIENTED_RGB_SIMILARITY
            ):
                reasons.append(f"pdf_page_{page['index'] + 1}_render_fidelity")
    else:
        if not projection.get("primary"):
            reasons.append("markdown_primary")
        if not projection.get("all_links_reachable"):
            reasons.append("markdown_link_reachability")
        if not projection.get("no_absolute_links"):
            reasons.append("markdown_absolute_link")
        if not projection.get("case_owned_naming"):
            reasons.append("markdown_case_owned_naming")
        styles = projection.get("link_style_counts", {})
        if (
            not styles.get("wiki_embed")
            or styles.get("wiki_link")
            or styles.get("markdown_embed")
            or styles.get("markdown_link")
        ):
            reasons.append("markdown_link_style")
        if not any(item["source_byte_identical"] for item in projection.get("retained_images", [])):
            reasons.append("markdown_retained_source_identity")
    return reasons


def _normalized_route_projection(route: str, value: dict[str, Any]) -> dict[str, Any] | None:
    projection = value.get("projection")
    if projection is None:
        return None
    common = {
        "success": value.get("success"),
        "final_placement": value.get("final_placement"),
        "artifact_suffixes": sorted(item["suffix"] for item in value.get("artifact_inventory", [])),
        "residue": value.get("processes", {}).get("residue_added", []),
    }
    if route in {"jpg", "webp"}:
        artifact = projection["artifact"]
        common["artifact"] = {
            key: artifact[key]
            for key in (
                "format",
                "mode",
                "raw_size",
                "display_size",
                "frames",
                "orientation",
                "selected_exif",
                "icc_bytes",
            )
        }
        common["display_similarity"] = projection["display_similarity"]
        common["metadata_fidelity"] = projection["metadata_fidelity"]
        return common
    if route == "pdf-original":
        common["source_frame_count"] = projection["source_frame_count"]
        common["page_count"] = projection["page_count"]
        common["pages"] = [
            {
                "index": page["index"],
                "rect": page["rect"],
                "rotation": page["rotation"],
                "render": page.get("render"),
                "embedded_images": [
                    {
                        key: image.get(key)
                        for key in (
                            "width",
                            "height",
                            "colorspace",
                            "extension",
                            "bytes",
                            "sha256",
                            "decoded_similarity_to_source_frame",
                        )
                    }
                    for image in page["embedded_images"]
                ],
            }
            for page in projection["pages"]
        ]
        return common
    common["markdown"] = {
        "markdown_file_count": len(projection["markdown_files"]),
        "retained_image_count": len(projection["retained_images"]),
        "retained_source_identity": [item["source_byte_identical"] for item in projection["retained_images"]],
        "retained_image_projection": [
            {
                key: item["projection"][key]
                for key in (
                    "format",
                    "mode",
                    "raw_size",
                    "display_size",
                    "frames",
                    "orientation",
                    "selected_exif",
                    "icc_bytes",
                )
            }
            for item in projection["retained_images"]
        ],
        "all_links_reachable": projection["all_links_reachable"],
        "no_absolute_links": projection["no_absolute_links"],
        "link_style_counts": projection["link_style_counts"],
        "case_owned_naming": projection["case_owned_naming"],
        "ocr_text": projection["ocr"]["normalized_text"],
    }
    return common


def _summarize(projects: dict[str, Any], inputs: dict[str, Path]) -> dict[str, Any]:
    route_acceptance = {
        project: {
            case: {route: _route_ok(route, projects[project][case][route]) for route in ROUTES} for case in inputs
        }
        for project in PROJECTS
    }
    ocr_comparisons: dict[str, Any] = {}
    for case in ("N1", "N2"):
        texts = {
            project: projects[project][case]["markdown-image-md"]["projection"]["ocr"]["normalized_text"]
            for project in PROJECTS
        }
        baseline = texts["docwen-ref-tk"]
        ocr_comparisons[case] = {
            "texts": texts,
            "all_nonempty": all(bool(text.strip()) for text in texts.values()),
            "all_equal": all(text == baseline for text in texts.values()),
        }
    image_projection_equality: dict[str, Any] = {}
    for case in inputs:
        image_projection_equality[case] = {}
        for route in ("jpg", "webp"):
            projections = {
                project: projects[project][case][route]["projection"]["artifact"]
                if projects[project][case][route]["projection"]
                else None
                for project in PROJECTS
            }
            normalized = {
                project: (
                    {
                        key: value[key]
                        for key in ("format", "mode", "raw_size", "display_size", "frames", "orientation", "icc_bytes")
                    }
                    if value
                    else None
                )
                for project, value in projections.items()
            }
            image_projection_equality[case][route] = {
                "projections": normalized,
                "all_equal": len({json.dumps(value, sort_keys=True) for value in normalized.values()}) == 1,
            }
    all_routes_ok = all(
        route_acceptance[project][case][route] for project in PROJECTS for case in inputs for route in ROUTES
    )
    failure_reasons = {
        project: {
            case: {
                route: _route_failure_reasons(route, projects[project][case][route])
                for route in ROUTES
                if not route_acceptance[project][case][route]
            }
            for case in inputs
        }
        for project in PROJECTS
    }
    shared_failures: dict[str, Any] = {}
    for case in inputs:
        for route in ROUTES:
            if all(not route_acceptance[project][case][route] for project in PROJECTS):
                shared_failures[f"{case}/{route}"] = {
                    project: failure_reasons[project][case][route] for project in PROJECTS
                }
    cross_project_route_equality: dict[str, Any] = {}
    for case in inputs:
        cross_project_route_equality[case] = {}
        for route in ROUTES:
            normalized = {
                project: _normalized_route_projection(route, projects[project][case][route]) for project in PROJECTS
            }
            baseline = normalized["docwen-ref-tk"]
            cross_project_route_equality[case][route] = {
                "all_equal": all(value == baseline for value in normalized.values()),
                "normalized": normalized,
            }
    acceptance = {
        "expected_execution_count": True,
        "all_routes_meet_frozen_thresholds": all_routes_ok,
        "n1_n2_ocr_nonempty_and_equal": all(
            value["all_nonempty"] and value["all_equal"] for value in ocr_comparisons.values()
        ),
        "image_mandatory_projections_equal": all(
            item["all_equal"] for case in image_projection_equality.values() for item in case.values()
        ),
        "all_route_mandatory_projections_equal": all(
            item["all_equal"] for case in cross_project_route_equality.values() for item in case.values()
        ),
    }
    return {
        "route_acceptance": route_acceptance,
        "ocr_comparisons": ocr_comparisons,
        "image_projection_equality": image_projection_equality,
        "failure_reasons": failure_reasons,
        "shared_failures": shared_failures,
        "cross_project_route_equality": cross_project_route_equality,
        "acceptance": acceptance,
    }


def _orchestrate(args: argparse.Namespace) -> int:
    contract_path = Path(args.stage_contract).resolve()
    evidence_root = contract_path.parent
    contract, inputs = _validate_contract(contract_path)
    roots = {
        "docwen-ref-tk": Path(args.tk_root).resolve(),
        "docwen-ref-pyside6": Path(args.pyside_root).resolve(),
        "docwen-current": Path(args.current_root).resolve(),
    }
    model_manifest = _model_manifest(roots)
    if not all(item["exists"] for project in model_manifest.values() for item in project.values()):
        raise RuntimeError("required Chinese OCR model is missing")
    input_manifest = {
        "stage_contract_path": str(contract_path),
        "stage_contract_sha256": _sha256(contract_path),
        "stage_contract": contract,
        "source_projections": {case: _image_projection(path) for case, path in inputs.items()},
        "ocr_models": model_manifest,
    }
    _write_json(evidence_root / "probe-input-manifest.json", input_manifest)
    if args.prepare_only:
        _print_json(input_manifest)
        return 0

    repair_reruns: list[dict[str, Any]] = []
    if args.reproject_only:
        prior = json.loads((evidence_root / "probe-result.json").read_text(encoding="utf-8"))
        projects = prior["projects"]
        execution_count = prior["execution_count"]
        for project in PROJECTS:
            for case, source in inputs.items():
                for route in ROUTES:
                    _reproject_route(route, projects[project][case][route], source)
    elif args.rerun_route:
        prior = json.loads((evidence_root / "probe-result.json").read_text(encoding="utf-8"))
        projects = prior["projects"]
        execution_count = prior["execution_count"]
        for spec in args.rerun_route:
            try:
                project, case, route = spec.split(":", 2)
            except ValueError as exc:
                raise RuntimeError(f"invalid --rerun-route value: {spec}") from exc
            if project not in PROJECTS or case not in inputs or route not in ROUTES:
                raise RuntimeError(f"invalid --rerun-route slot: {spec}")
            root = roots[project]
            old = projects[project][case][route]
            new = _run_route(
                project=project,
                root=root,
                executable=root / ".venv" / "Scripts" / "docwen.exe",
                case=case,
                route=route,
                frozen_source=inputs[case],
                evidence_root=evidence_root,
            )
            projects[project][case][route] = new
            repair_reruns.append(
                {
                    "slot": spec,
                    "reason": "targeted post-fix replacement of the canonical route slot",
                    "superseded": {
                        "success": old.get("success"),
                        "artifact_inventory": old.get("artifact_inventory"),
                        "failure_reasons": _route_failure_reasons(route, old),
                    },
                    "replacement": {
                        "success": new.get("success"),
                        "artifact_inventory": new.get("artifact_inventory"),
                        "failure_reasons": _route_failure_reasons(route, new),
                    },
                }
            )
        _write_json(evidence_root / "repair-reruns.json", repair_reruns)
    else:
        projects = {}
        execution_count = 0
        for project in PROJECTS:
            root = roots[project]
            executable = root / ".venv" / "Scripts" / "docwen.exe"
            projects[project] = {}
            for case, source in inputs.items():
                projects[project][case] = {}
                for route in ROUTES:
                    value = _run_route(
                        project=project,
                        root=root,
                        executable=executable,
                        case=case,
                        route=route,
                        frozen_source=source,
                        evidence_root=evidence_root,
                    )
                    projects[project][case][route] = value
                    execution_count += 1
                    print(
                        f"[{execution_count:02d}/60] {project} {case} {route}: success={value['success']}", flush=True
                    )

    summary = _summarize(projects, inputs)
    summary["acceptance"]["expected_execution_count"] = execution_count == 60
    result = {
        "probe_id": "FA-08-B1-B2-N1-N2-2026-07-22",
        "contract_id": contract["contract_id"],
        "stage_id": contract["stage_id"],
        "stage_contract_sha256": input_manifest["stage_contract_sha256"],
        "execution_count": execution_count,
        "repair_route_execution_count": len(repair_reruns),
        "repair_reruns": repair_reruns,
        "projects": projects,
        **summary,
    }
    result["pass"] = all(result["acceptance"].values())
    _write_json(evidence_root / "probe-result.json", result)
    _print_json(result)
    return 0 if result["pass"] else 1


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser()
    parser.add_argument("--worker-project", choices=("tk", "pyside"))
    parser.add_argument("--worker-route", choices=ROUTES)
    parser.add_argument("--worker-source")
    parser.add_argument("--worker-final")
    parser.add_argument("--worker-result")
    parser.add_argument("--current-root", default=".")
    parser.add_argument("--tk-root", default="../docwen-ref-tk")
    parser.add_argument("--pyside-root", default="../docwen-ref-pyside6")
    parser.add_argument("--stage-contract", required=False)
    parser.add_argument("--prepare-only", action="store_true")
    parser.add_argument("--reproject-only", action="store_true")
    parser.add_argument("--rerun-route", action="append", default=[])
    return parser


def main() -> int:
    args = _parser().parse_args()
    if args.worker_project:
        return _worker(args)
    if not args.stage_contract:
        raise SystemExit("--stage-contract is required")
    return _orchestrate(args)


if __name__ == "__main__":
    raise SystemExit(main())

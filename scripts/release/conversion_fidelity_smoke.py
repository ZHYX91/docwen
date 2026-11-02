"""Synthetic file-content regressions against one packaged CLI, without product imports.

The caller owns the new evidence directory and the candidate's provenance receipt.
This module never launches source entry points or cleans up another run's files.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import queue
import re
import subprocess
import sys
import threading
import time
from pathlib import Path
from typing import IO, Any
from xml.etree import ElementTree
from zipfile import ZipFile

import openpyxl
from lxml import html
from markdown_it import MarkdownIt
from PIL import Image
from pptx import Presentation
from pptx.util import Inches

XLSX_MEDIA = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
LITERAL = "<b>literal</b> [manual](https://example.com) *not italic* &copy;"
BLOCK_TEXT = "# literal heading\n- not a bullet\n1. not ordered\n---\n$$price$$"
NONFINITE_TEXT = ("NaN", "nan", "Infinity", "-Infinity", "1e309")
FRAME_COLORS = ((255, 0, 0, 0), (0, 255, 0, 96), (0, 0, 255, 255), (10, 20, 30, 255))
TERMINALS = {"task/completed", "task/failed", "task/cancelled"}
MAX_FRAME_BYTES = 16 * 1024 * 1024


def require(condition: bool, message: str) -> None:
    if not condition:
        raise RuntimeError(message)


def identity(path: Path) -> dict[str, Any]:
    raw = path.read_bytes()
    return {"path": str(path), "size_bytes": len(raw), "sha256": hashlib.sha256(raw).hexdigest()}


def packaged_binary(directory: Path) -> Path:
    binary = directory.resolve(strict=True) / ("DocWenCLI.exe" if os.name == "nt" else "DocWenCLI")
    with binary.open("rb") as stream:
        magic = stream.read(4)
    require(magic[:2] == b"MZ" if os.name == "nt" else magic == b"\x7fELF", "expected_native_packaged_cli")
    return binary


def isolated_environment(work: Path) -> dict[str, str]:
    env = os.environ.copy()
    for name in ("PYTHONPATH", "PYTHONHOME", "DOCWEN_WORKSPACE_ROOT", "DOCWEN_RUNTIME_ROOT"):
        env.pop(name, None)
    env.update(PYTHONUTF8="1", PYTHONIOENCODING="utf-8", DOCWEN_LOG_TO_TEMP="")
    for key, directory in (("CONFIG", "config"), ("DATA", "data"), ("LOG", "logs")):
        path = work / directory
        path.mkdir()
        env[f"DOCWEN_{key}_DIR"] = str(path)
    # Same internal isolation hooks as the native packaged verifier; no user's
    # saved compression mode, config, logs or workspace can affect these cases.
    (work / "config" / "image.toml").write_text('compress_mode = "lossless"\n', encoding="utf-8")
    governed = work / "workspace"
    governed.mkdir()
    (governed / "README.md").write_text("# DocWen 本地工作区\n", encoding="utf-8")
    for name in ("runtime", "system"):
        (governed / "temp" / name).mkdir(parents=True)
    env["DOCWEN_WORKSPACE_ROOT"] = str(governed)
    env["DOCWEN_RUNTIME_ROOT"] = str(governed / "temp" / "runtime")
    for key in ("TEMP", "TMP", "TMPDIR"):
        env[key] = str(governed / "temp" / "system")
    return env


def _exact(stream: IO[bytes], length: int) -> bytes:
    parts = bytearray()
    while len(parts) < length:
        part = stream.read(length - len(parts))
        if not part:
            raise EOFError("machine_frame_truncated")
        parts.extend(part)
    return bytes(parts)


def read_frame(stream: IO[bytes]) -> dict[str, Any] | None:
    header = bytearray()
    while not header.endswith(b"\r\n\r\n"):
        byte = stream.read(1)
        if not byte and not header:
            return None
        require(bool(byte) and len(header) < 64, "machine_frame_header_invalid")
        header.extend(byte)
    match = re.fullmatch(rb"Content-Length: ([1-9][0-9]*)\r\n\r\n", header)
    require(match is not None, "machine_frame_header_invalid")
    assert match is not None
    length = int(match[1])
    require(length <= MAX_FRAME_BYTES, "machine_frame_exceeds_wire_limit")
    value = json.loads(_exact(stream, length).decode("utf-8"))
    require(isinstance(value, dict), "machine_frame_not_object")
    return value


class Machine:
    """One owned process; bounded reads and complete framed output evidence."""

    def __init__(self, binary: Path, work: Path, env: dict[str, str], timeout: float):
        self.timeout = timeout
        self.messages: list[dict[str, Any]] = []
        self.pending: queue.Queue[dict[str, Any] | BaseException | None] = queue.Queue()
        self.request_id = 0
        self.tasks: list[str] = []
        self.stderr = (work / "machine.stderr.txt").open("wb")
        self.requests = (work / "machine.requests.jsonl").open("w", encoding="utf-8")
        self.responses = (work / "machine.responses.jsonl").open("w", encoding="utf-8")
        try:
            self.process = subprocess.Popen(
                [str(binary), "serve", "--stdio"],
                cwd=work,
                env=env,
                stdin=subprocess.PIPE,
                stdout=subprocess.PIPE,
                stderr=self.stderr,
            )
        except BaseException:
            self.stderr.close()
            self.requests.close()
            self.responses.close()
            raise
        self.reader = threading.Thread(target=self._read, daemon=True)
        self.reader.start()

    def _read(self) -> None:
        assert self.process.stdout is not None
        try:
            while (message := read_frame(self.process.stdout)) is not None:
                self.responses.write(json.dumps(message, ensure_ascii=False) + "\n")
                self.responses.flush()
                self.pending.put(message)
            self.pending.put(None)
        except BaseException as exc:
            self.pending.put(exc)

    def _next(self, deadline: float) -> dict[str, Any]:
        remaining = deadline - time.monotonic()
        if remaining <= 0:
            raise TimeoutError("packaged_machine_response_timeout")
        try:
            value = self.pending.get(timeout=remaining)
        except queue.Empty as exc:
            raise TimeoutError("packaged_machine_response_timeout") from exc
        if time.monotonic() >= deadline:
            raise TimeoutError("packaged_machine_response_timeout")
        if isinstance(value, BaseException):
            raise value
        require(value is not None, "packaged_machine_unexpected_eof")
        assert value is not None
        self.messages.append(value)
        return value

    def request(self, method: str, params: dict[str, Any]) -> dict[str, Any]:
        self.request_id += 1
        request = {"jsonrpc": "2.0", "id": self.request_id, "method": method, "params": params}
        self.requests.write(json.dumps(request, ensure_ascii=False) + "\n")
        self.requests.flush()
        body = json.dumps(request, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
        deadline = time.monotonic() + self.timeout
        self._send(f"Content-Length: {len(body)}\r\n\r\n".encode("ascii") + body, deadline)
        while True:
            response = self._next(deadline)
            if response.get("id") == self.request_id:
                require("error" not in response and isinstance(response.get("result"), dict), f"{method}:{response}")
                return response["result"]
            require(response.get("method") == "task/progress", f"unexpected_machine_response:{response}")

    def _send(self, frame: bytes, deadline: float) -> None:
        # A stopped reader can fill the OS pipe even while our response reader
        # remains live. Bound the write with the same request deadline.
        failures: list[BaseException] = []

        def write() -> None:
            try:
                assert self.process.stdin is not None
                self.process.stdin.write(frame)
                self.process.stdin.flush()
            except BaseException as exc:
                failures.append(exc)

        writer = threading.Thread(target=write, daemon=True)
        writer.start()
        writer.join(timeout=max(0.0, deadline - time.monotonic()))
        if writer.is_alive() or time.monotonic() >= deadline:
            # Release a blocked pipe before close() attempts to close stdin.
            self._stop()
            writer.join(timeout=5)
            raise TimeoutError("packaged_machine_request_timeout")
        if failures:
            raise failures[0]

    def _stop(self) -> None:
        if self.process.poll() is None:
            self.process.terminate()
            try:
                self.process.wait(timeout=5)
            except subprocess.TimeoutExpired:
                self.process.kill()
                self.process.wait(timeout=5)

    def task(
        self, case: str, inputs: list[Path], staging: Path, capability: str, options: dict[str, Any]
    ) -> dict[str, Any]:
        staging.mkdir()
        media = {".xlsx": XLSX_MEDIA, ".tif": "image/tiff", ".png": "image/png"}
        refs = []
        for index, path in enumerate(inputs):
            ref = identity(path)
            refs.append(
                {
                    "input_id": f"input.{index}",
                    "role": "source",
                    "kind": "resource",
                    "logical_path": path.name,
                    "locator": {"kind": "local_path", "path": str(path)},
                    "media_type": media[path.suffix],
                    "size_bytes": ref["size_bytes"],
                    "sha256": ref["sha256"],
                }
            )
        plan = self.request(
            "task/plan",
            {
                "capability_id": capability,
                "inputs": refs,
                "options": options,
                "output": {
                    "staging_root": {"kind": "local_path", "path": str(staging)},
                    "staging_policy": "require_empty",
                },
            },
        )
        accepted = self.request("task/execute", {"plan_id": plan["plan_id"]})
        require(accepted.get("state") == "accepted", f"{case}:not_accepted")
        task_id = accepted["task_id"]
        self.tasks.append(task_id)
        deadline = time.monotonic() + self.timeout
        while True:
            message = self._next(deadline)
            require(message.get("params", {}).get("task_id") == task_id, f"{case}:task_identity_mismatch")
            if message.get("method") in TERMINALS:
                return message
            require(message.get("method") == "task/progress", f"{case}:unexpected_notification")

    def close(self) -> None:
        try:
            if self.process.stdin is not None:
                self.process.stdin.close()
            self.process.wait(timeout=15)
        finally:
            self._stop()
            self.reader.join(timeout=5)
            self.stderr.close()
            self.requests.close()
            self.responses.close()
            if self.process.stdout is not None:
                self.process.stdout.close()
        require(self.process.returncode == 0, "packaged_machine_exit_failed")
        while not self.pending.empty():
            value = self.pending.get_nowait()
            if isinstance(value, BaseException):
                raise value
            if value is not None:
                self.messages.append(value)
        for task_id in self.tasks:
            messages = [value for value in self.messages if value.get("params", {}).get("task_id") == task_id]
            require(sum(value.get("method") in TERMINALS for value in messages) == 1, "machine_terminal_not_unique")
            require(messages[-1].get("method") in TERMINALS, "machine_progress_after_terminal")
            sequences = [value["params"]["sequence"] for value in messages]
            require(sequences == sorted(set(sequences)), "machine_sequence_not_strictly_increasing")


def output_from_terminal(terminal: dict[str, Any], staging: Path, media_type: str) -> Path:
    require(terminal.get("method") == "task/completed", f"conversion_not_completed:{terminal}")
    bundle = terminal["params"]["bundle"]
    require(bundle.get("schema") == "docwen.artifact_bundle.v3", "bundle_schema_mismatch")
    require(bundle.get("task_id") == terminal["params"]["task_id"], "bundle_task_mismatch")
    artifacts = bundle["artifacts"]
    require(len(artifacts) == 1 and artifacts[0]["media_type"] == media_type, "unexpected_bundle_artifacts")
    artifact = artifacts[0]
    locator = artifact["locator"]
    require(isinstance(locator, str) and "\\" not in locator and not Path(locator).is_absolute(), "invalid_locator")
    path = (staging / locator).resolve(strict=True)
    require(path.is_relative_to(staging.resolve()), "artifact_outside_staging")
    actual = identity(path)
    require(all(actual[key] == artifact[key] for key in ("size_bytes", "sha256")), "artifact_integrity_mismatch")
    return path


def write_sheet(path: Path, cells: dict[str, Any]) -> None:
    workbook = openpyxl.Workbook()
    assert workbook.active is not None
    for coordinate, value in cells.items():
        workbook.active[coordinate] = value
    workbook.save(path)
    workbook.close()


def verify_sheet(path: Path, expected: dict[str, Any]) -> None:
    workbook = openpyxl.load_workbook(path, data_only=True)
    try:
        assert workbook.active is not None
        for coordinate, value in expected.items():
            require(workbook.active[coordinate].value == value, f"xlsx_value_mismatch:{coordinate}")
        with ZipFile(path) as archive:
            root = ElementTree.fromstring(archive.read("xl/worksheets/sheet1.xml"))
            require(len(root.findall(".//{*}c")) == len(expected), "xlsx_unexpected_materialized_cells")
    finally:
        workbook.close()


def write_fixtures(directory: Path) -> dict[str, list[Path]]:
    directory.mkdir()
    fixtures: dict[str, list[Path]] = {}
    fixtures["diagnostics64"] = [directory / f"part-{index:02d}.xlsx" for index in range(1, 65)]
    for index, path in enumerate(fixtures["diagnostics64"], 1):
        write_sheet(path, {"A1": index})
    for case in ("nonfinite", "overflow", "sparse"):
        fixtures[case] = [directory / f"{case}-{side}.xlsx" for side in (0, 1)]
        for side, path in enumerate(fixtures[case]):
            if case == "nonfinite":
                cells = {
                    f"{chr(65 + i)}{row}": text if row == side + 1 else 1
                    for i, text in enumerate(NONFINITE_TEXT)
                    for row in (1, 2)
                }
                # Pin coordinate alignment: otherwise the merger's existing
                # value-match fallback can align the two swapped rows.
                cells["F3"] = "coordinate-anchor"
            else:
                cells = {"A1": 1e308} if case == "overflow" else {"A1": 1, "SR512": 2}
            write_sheet(path, cells)
    tiff = directory / "three-pages.tif"
    frames = [Image.new("RGBA", (2, 2), color) for color in FRAME_COLORS[:3]]
    try:
        frames[0].save(tiff, save_all=True, append_images=frames[1:])
    finally:
        for frame in frames:
            frame.close()
    opaque = directory / "opaque.png"
    with Image.new("RGB", (2, 2), FRAME_COLORS[3][:3]) as image:
        image.save(opaque)
    fixtures["tiff-alpha"] = [tiff, opaque]
    for mode in ("RGB", "RGBA"):
        path = directory / f"pixels-{mode}.png"
        with Image.new(mode, (32, 32)) as image:
            for y in range(32):
                for x in range(32):
                    color = ((x * 37 + y * 11) % 256, (x * 13 + y * 53) % 256, (x * 73 + y * 7) % 256)
                    image.putpixel((x, y), (*color, (x * 17) % 256) if mode == "RGBA" else color)
            image.save(path)
        fixtures[f"webp-{mode}"] = [path]
    presentation = Presentation()
    presentation.core_properties.title = "Guide *literal*"
    slide = presentation.slides.add_slide(presentation.slide_layouts[5])
    assert slide.shapes.title is not None
    slide.shapes.title.text = LITERAL
    slide.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(1)).text_frame.text = LITERAL
    slide.shapes.add_textbox(Inches(1), Inches(3), Inches(8), Inches(2)).text_frame.text = BLOCK_TEXT
    path = directory / "literal-title.pptx"
    presentation.save(path)
    fixtures["presentation"] = [path]
    return fixtures


def verify_tiff(path: Path, *, alpha: bool) -> None:
    with Image.open(path) as saved:
        require(getattr(saved, "n_frames", 1) == 4, "tiff_missing_frames")
        for index, color in enumerate(FRAME_COLORS):
            saved.seek(index)
            require(saved.mode == ("RGBA" if alpha else "RGB"), "tiff_mode_mismatch")
            if alpha:
                require(saved.getpixel((0, 0)) == color, f"tiff_pixel_mismatch:{index}")
        if not alpha:
            saved.seek(0)
            require(saved.getpixel((0, 0)) == (255, 255, 255), "explicit_rgb_not_flattened")


def verify_webp(path: Path, source: Path) -> None:
    with Image.open(source) as original, Image.open(path) as output:
        require(output.format == "WEBP" and output.mode == original.mode, "webp_format_or_mode_mismatch")
        require(output.size == original.size and output.tobytes() == original.tobytes(), "webp_not_lossless")


def verify_presentation(path: Path) -> None:
    markdown = path.read_text(encoding="utf-8")
    body = re.sub(r"\A---\r?\n.*?\r?\n---(?:\r?\n|\Z)", "", markdown, count=1, flags=re.S)
    rendered = html.fragment_fromstring(MarkdownIt("commonmark", {"html": True}).render(body), create_parent=True)
    require(
        [node.text_content() for node in rendered.xpath("//h1")] == ["Guide *literal*"], "pptx_document_title_changed"
    )
    require([node.text_content() for node in rendered.xpath("//h2")] == [LITERAL], "pptx_slide_title_changed")
    require(rendered.text_content().count(LITERAL) == 2, "pptx_title_duplicated_or_equal_body_lost")
    require(all(line in rendered.text_content() for line in BLOCK_TEXT.splitlines()), "pptx_literal_text_changed")
    require(not rendered.xpath("//a | //b | //em | //ul | //ol | //hr"), "pptx_text_became_active_markup")


def cli_convert(binary: Path, source: Path, target: str, work: Path, env: dict[str, str], timeout: float) -> Path:
    destination = work / f"cli-{source.stem}"
    try:
        process = subprocess.run(
            [
                str(binary),
                "convert",
                str(source),
                "--to",
                target,
                "--output-dir",
                str(destination),
                "--json",
                "--quiet",
            ],
            cwd=work,
            env=env,
            capture_output=True,
            timeout=timeout,
        )
    except subprocess.TimeoutExpired as exc:
        (work / f"{source.stem}.stdout.json").write_bytes(exc.stdout or b"")
        (work / f"{source.stem}.stderr.txt").write_bytes(exc.stderr or b"")
        raise
    (work / f"{source.stem}.stdout.json").write_bytes(process.stdout)
    (work / f"{source.stem}.stderr.txt").write_bytes(process.stderr)
    require(process.returncode == 0, f"packaged_cli_failed:{source.name}:{process.returncode}")
    payload = json.loads(process.stdout.decode("utf-8"))
    require(payload.get("success") is True, f"packaged_cli_result_failed:{source.name}")
    path = Path(payload["data"]["output"])
    if not path.is_absolute():
        path = work / path
    path = path.resolve(strict=True)
    require(path.is_relative_to(destination.resolve()), "cli_output_outside_destination")
    return path


def run_smoke(binary: Path, work: Path, *, candidate_id: str = "", timeout: float = 120) -> dict[str, Any]:
    work = work.resolve()
    work.mkdir(parents=True, exist_ok=False)
    binary_before = identity(binary)
    report: dict[str, Any] = {
        "schema": "docwen.conversion_fidelity_smoke.v1",
        "candidate_id": candidate_id,
        "binary": binary_before,
        "status": "failed",
        "cases": [],
        "limitations": [
            "Candidate provenance is established by the caller's native fetch receipt, not the candidate_id label.",
            "Sparse fixture is bounded at 512x512; correct content does not prove internal allocation complexity.",
            "No GUI, OCR, Office backend, SmartArt, notes, cancellation or other platform acceptance is implied.",
        ],
    }
    machine: Machine | None = None
    try:
        env = isolated_environment(work)
        fixtures = write_fixtures(work / "inputs")
        originals = {path: identity(path) for paths in fixtures.values() for path in paths}
        report["inputs"] = list(originals.values())
        machine = Machine(binary, work, env, timeout)
        report["initialize"] = machine.request(
            "initialize",
            {
                "protocol": {"name": "docwen.machine", "major": 2, "minor": 0},
                "client": {"name": "conversion-fidelity-smoke", "version": "1.0.0"},
                "features": {"progress": True, "cancellation": False},
            },
        )
        require(report["initialize"].get("artifact_bundle_schema") == "docwen.artifact_bundle.v3", "handshake_mismatch")
        for case in ("diagnostics64", "nonfinite", "overflow", "sparse", "tiff-alpha", "tiff-rgb"):
            is_image = case.startswith("tiff-")
            staging = work / case
            terminal = machine.task(
                case,
                fixtures["tiff-alpha" if is_image else case],
                staging,
                "merge.images.to_tiff" if is_image else "merge.xlsx.tables",
                {"keep_alpha": case == "tiff-alpha"} if is_image else {"merge_mode": "cell", "offset_range": 0},
            )
            case_result: dict[str, Any] = {"case": case, "terminal": terminal["method"]}
            if case == "overflow":
                require(terminal["method"] == "task/failed", "overflow_was_published")
                require("finite numeric range" in terminal["params"]["error"]["message"], "overflow_reason_missing")
                require(not any(staging.iterdir()), "overflow_left_artifacts")
            else:
                output = output_from_terminal(terminal, staging, "image/tiff" if is_image else XLSX_MEDIA)
                case_result["output"] = identity(output)
                if case == "diagnostics64":
                    verify_sheet(output, {"A1": 2080})
                    count = len(terminal["params"]["diagnostics"])
                    require(count >= 65, "diagnostics_were_truncated")
                    case_result["diagnostics"] = count
                elif case == "nonfinite":
                    verify_sheet(
                        output,
                        {
                            "F3": "coordinate-anchor",
                            **{
                                f"{chr(65 + i)}{row}": f"{text},1" if row == 1 else f"1,{text}"
                                for i, text in enumerate(NONFINITE_TEXT)
                                for row in (1, 2)
                            },
                        },
                    )
                elif case == "sparse":
                    verify_sheet(output, {"A1": 2, "SR512": 4})
                else:
                    verify_tiff(output, alpha=case == "tiff-alpha")
            report["cases"].append(case_result)
        machine.close()
        machine = None
        for case in ("webp-RGB", "webp-RGBA", "presentation"):
            source = fixtures[case][0]
            output = cli_convert(binary, source, "md" if case == "presentation" else "webp", work, env, timeout)
            if case == "presentation":
                verify_presentation(output)
            else:
                verify_webp(output, source)
            report["cases"].append({"case": case, "output": identity(output)})
        require(all(identity(path) == before for path, before in originals.items()), "source_fixture_changed")
        require(identity(binary) == binary_before, "candidate_binary_changed")
        report["status"] = "passed"
    except BaseException as exc:
        report["error"] = f"{type(exc).__name__}: {exc}"
        raise
    finally:
        if machine is not None:
            try:
                machine.close()
            except Exception as exc:
                report["cleanup_error"] = f"{type(exc).__name__}: {exc}"
        (work / "summary.json").write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return report


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--binary-dir", type=Path, required=True)
    parser.add_argument(
        "--evidence-dir", type=Path, required=True, help="New directory within this acceptance run's owned area."
    )
    parser.add_argument(
        "--candidate-id", required=True, help="Exact candidate identifier from the native fetch receipt."
    )
    parser.add_argument("--expected-binary-sha256", help="Optional exact executable hash from the same candidate.")
    parser.add_argument(
        "--timeout", type=float, default=120, help="Seconds per task/CLI command, excluding process startup."
    )
    args = parser.parse_args(argv)
    binary = packaged_binary(args.binary_dir)
    if args.expected_binary_sha256:
        require(identity(binary)["sha256"] == args.expected_binary_sha256, "candidate_binary_hash_mismatch")
    require(args.timeout > 0, "timeout_must_be_positive")
    report = run_smoke(binary, args.evidence_dir, candidate_id=args.candidate_id, timeout=args.timeout)
    print(
        json.dumps(
            {
                "status": report["status"],
                "cases": len(report["cases"]),
                "summary": str(args.evidence_dir / "summary.json"),
            }
        )
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())

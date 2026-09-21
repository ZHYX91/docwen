"""Bounded, cancellable local Mermaid CLI rendering for document images."""

from __future__ import annotations

import contextlib
import json
import os
import signal
import subprocess
import tempfile
import time
from collections.abc import Iterator
from io import BytesIO
from pathlib import Path
from typing import Any, cast

from PIL import Image

from docwen_core.mermaid_runtime import inspect_mermaid_runtime

# Import the CLI only after Python attaches this process to its cleanup
# boundary, closing the Windows race between creation and job assignment.
_BOOTSTRAP = r"""import { once } from 'node:events';
import { pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';
import { readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';
async function main() {
await once(process.stdin, 'data');
process.stdin.destroy();
const [entry, root] = process.argv.slice(2);
const { renderMermaid } = await import(pathToFileURL(path.join(path.dirname(entry), 'index.js')).href);
const puppeteer = createRequire(entry)('puppeteer');
const config = JSON.parse(await readFile(path.join(root, 'mermaid.json'), 'utf8'));
const browserConfig = JSON.parse(await readFile(path.join(root, 'browser.json'), 'utf8'));
const source = await readFile(path.join(root, 'diagram.mmd'), 'utf8');
const browser = await puppeteer.launch(browserConfig);
let blocked = 0;
try {
  const context = { newPage: async () => {
    const page = await browser.newPage();
    const screenshot = page.screenshot.bind(page);
    page.screenshot = async options => {
      const external = await page.evaluate(() => [...document.querySelectorAll('image, img')].some(image => {
        const href = image.getAttribute('href') || image.getAttribute('xlink:href') || image.getAttribute('src') || '';
        return !/^data:image\//i.test(href);
      }));
      if (external) throw new Error('External or local-file diagram images are unsupported');
      return screenshot(options);
    };
    page.on('request', request => {
      const proceed = request.continue.bind(request);
      request.continue = (overrides, priority) => {
        if (request.isInterceptResolutionHandled()) return Promise.resolve();
        if (/^(data|blob|about):/.test(request.url())) return proceed(overrides, priority);
        blocked += 1;
        return request.abort('blockedbyclient', priority);
      };
    });
    return page;
  }};
  // CLI's interceptor fulfills only its explicitly registered package resources.
  // Every other intercepted network or local-file request must fail closed.
  const result = await renderMermaid(context, source, 'png', {
    viewport: { width: 1600, height: 900, deviceScaleFactor: 2 },
    backgroundColor: 'white', mermaidConfig: config,
  });
  if (blocked) throw new Error('External or local-file diagram resources were blocked');
  await writeFile(path.join(root, 'diagram.png'), result.data);
} finally { await browser.close(); }
}
main().catch(error => { console.error(String(error.message).slice(0, 3000)); process.exitCode = 1; });
"""


class MermaidRenderError(RuntimeError):
    """The diagram cannot be rendered; the caller should retain its source."""


def _attach_windows_job(process: subprocess.Popen) -> Any:
    import win32api
    import win32con
    import win32job

    job = cast(Any, win32job.CreateJobObject(None, ""))
    try:
        limits = win32job.QueryInformationJobObject(job, win32job.JobObjectExtendedLimitInformation)
        limits["BasicLimitInformation"]["LimitFlags"] |= win32job.JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
        win32job.SetInformationJobObject(job, win32job.JobObjectExtendedLimitInformation, limits)
        handle = win32api.OpenProcess(win32con.PROCESS_SET_QUOTA | win32con.PROCESS_TERMINATE, False, process.pid)
        try:
            win32job.AssignProcessToJobObject(job, handle)
        finally:
            win32api.CloseHandle(handle)
        return job
    except BaseException:
        job.Close()
        raise


def _stop_windows_job(job: Any) -> None:
    import win32job

    try:
        # Closing a kill-on-close job only initiates termination. Chromium can
        # still hold profile files after the Node parent has exited.
        win32job.TerminateJobObject(job, 1)
        deadline = time.monotonic() + 5
        while win32job.QueryInformationJobObject(job, win32job.JobObjectBasicAccountingInformation)["ActiveProcesses"]:
            if time.monotonic() >= deadline:
                raise MermaidRenderError("Mermaid child processes did not stop before cleanup")
            time.sleep(0.01)
    finally:
        job.Close()


def _run_renderer(command: list[str], job_dir: Path, cancellation: Any, timeout: float) -> None:
    environment = dict(os.environ)
    environment.pop("NODE_OPTIONS", None)
    environment["TMPDIR"] = environment["TEMP"] = environment["TMP"] = str(job_dir)
    log_path = job_dir / "renderer.log"
    process = None
    job = None
    try:
        with log_path.open("wb") as log:
            process = subprocess.Popen(
                command,
                cwd=job_dir,
                env=environment,
                stdin=subprocess.PIPE,
                stdout=log,
                stderr=subprocess.STDOUT,
                creationflags=subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0,
                start_new_session=os.name != "nt",
            )
            if os.name == "nt":
                job = _attach_windows_job(process)
            if cancellation is not None:
                cancellation.check()
            assert process.stdin is not None
            process.stdin.write(b"go\n")
            process.stdin.close()
            deadline = time.monotonic() + timeout
            while process.poll() is None:
                if cancellation is not None:
                    cancellation.check()
                if time.monotonic() >= deadline:
                    raise MermaidRenderError(f"Mermaid rendering exceeded {timeout:g} seconds")
                if log_path.stat().st_size > 1024 * 1024:
                    raise MermaidRenderError("Mermaid renderer diagnostic output exceeded its limit")
                time.sleep(0.05)
            if cancellation is not None:
                cancellation.check()
            if process.returncode != 0:
                with log_path.open("rb") as diagnostic:
                    diagnostic.seek(max(0, log_path.stat().st_size - 4800))
                    detail = " ".join(diagnostic.read(4800).decode("utf-8", errors="replace").split())[-1200:]
                raise MermaidRenderError(f"Mermaid CLI exited with status {process.returncode}: {detail}")
    finally:
        try:
            if job is not None:
                _stop_windows_job(job)
            elif process is not None and os.name != "nt":
                with contextlib.suppress(ProcessLookupError):
                    os.killpg(process.pid, signal.SIGKILL)
        finally:
            if process is not None:
                if process.poll() is None:
                    process.kill()
                process.wait(timeout=5)
                if process.stdin is not None:
                    process.stdin.close()


@contextlib.contextmanager
def _renderer_directory(root: Path) -> Iterator[Path]:
    temporary = tempfile.TemporaryDirectory(prefix="docwen-mermaid-", dir=root)
    try:
        yield Path(temporary.name)
    finally:
        # Windows may briefly retain a browser profile's file handles after
        # every process in the job has exited. Retry only sharing violations;
        # access-control failures and an exhausted deadline remain visible.
        deadline = time.monotonic() + 5
        while True:
            try:
                temporary.cleanup()
                break
            except PermissionError as error:
                if getattr(error, "winerror", None) not in (32, 33) or time.monotonic() >= deadline:
                    raise
                time.sleep(0.05)


def render_mermaid_png(
    source: str,
    *,
    work_dir: str | Path,
    cancellation: Any = None,
    timeout_seconds: float = 30.0,
    cli_path: str = "",
) -> bytes:
    """Render one source block without downloading tools or changing the source."""
    if not source.strip():
        raise MermaidRenderError("Mermaid source is empty")
    if len(source.encode("utf-8")) > 200_000 or len(source) > 50_000:
        raise MermaidRenderError("Mermaid source exceeds the diagram size limit")
    if timeout_seconds <= 0:
        raise ValueError("Mermaid timeout must be positive")
    if cancellation is not None:
        cancellation.check()
    runtime = inspect_mermaid_runtime(cli_path)
    if not runtime.available:
        raise MermaidRenderError(
            f"Compatible local Mermaid CLI unavailable ({runtime.reason}); install CLI 11.16+ with Mermaid 11.16+ "
            "from the 11.x family and put mmdc on PATH or set DOCWEN_MERMAID_CLI"
        )
    root = Path(work_dir).resolve()
    root.mkdir(parents=True, exist_ok=True)
    with _renderer_directory(root) as job_dir:
        (job_dir / "bootstrap.mjs").write_text(_BOOTSTRAP, encoding="utf-8")
        (job_dir / "diagram.mmd").write_text(source, encoding="utf-8", newline="")
        config = {
            "securityLevel": "strict",
            "theme": "default",
            "maxTextSize": 50_000,
            "maxEdges": 500,
            "secure": ["secure", "securityLevel", "startOnLoad", "maxTextSize", "maxEdges", "htmlLabels"],
            "htmlLabels": False,
            "flowchart": {"htmlLabels": False},
            "fontFamily": '"Microsoft YaHei", "Noto Sans CJK SC", sans-serif',
        }
        (job_dir / "mermaid.json").write_text(json.dumps(config), encoding="utf-8")
        # CLI's resource interceptor serves its packaged assets in process.
        # Other browser network requests have no direct or loopback bypass.
        browser = {
            "headless": "shell",
            "userDataDir": str(job_dir / "browser-profile"),
            "args": [
                "--proxy-server=http://127.0.0.1:9",
                "--proxy-bypass-list=<-loopback>",
                "--disable-background-networking",
                "--disable-sync",
                "--no-first-run",
            ],
        }
        (job_dir / "browser.json").write_text(json.dumps(browser), encoding="utf-8")
        output = job_dir / "diagram.png"
        command = [
            runtime.node,
            "--max-old-space-size=512",
            str(job_dir / "bootstrap.mjs"),
            runtime.cli_entry,
            str(job_dir),
        ]
        _run_renderer(command, job_dir, cancellation, timeout_seconds)
        if not output.is_file() or not 0 < output.stat().st_size <= 32 * 1024 * 1024:
            raise MermaidRenderError("Mermaid CLI produced no bounded PNG output")
        data = output.read_bytes()
        with Image.open(BytesIO(data)) as image:
            if image.format != "PNG" or image.width * image.height > 40_000_000:
                raise MermaidRenderError("Mermaid PNG dimensions exceed the image limit")
            image.verify()
        return data

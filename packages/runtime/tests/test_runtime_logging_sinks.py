"""Real file and console logging sinks isolated in an owned child process."""

from __future__ import annotations

import pytest

pytestmark = [pytest.mark.integration, pytest.mark.release_gate]


@pytest.mark.integration
def test_stdlib_namespaces_reach_real_file_and_console_once(tmp_path):
    """A child process owns and closes all logging-global state for this check."""
    import json
    import sys

    from tests.support.subprocess_runner import run_subprocess

    program = r"""import contextlib
import io
import json
import logging
import os
import sys
from pathlib import Path
from docwen_runtime.logging import get_logging_runtime_state, init_logging

os.environ.pop("DOCWEN_LOG_DIR", None)
os.environ.pop("DOCWEN_LOG_TO_TEMP", None)
root = Path(sys.argv[1])
console = io.StringIO()
names = ["docwen", "docwen.child", "docwen_runtime.ipc.single_instance",
         "docwen_runtime.config.loader", "docwen_plugin_markdown.renderer", "outside"]
config = {"logger": {"enable": True, "console_enable": True,
          "level": "WARNING", "console_level": "WARNING",
          "console_colorize": "never", "directory_mode": "custom", "directory": str(root),
          "format": "%(levelname)s:%(name)s:%(message)s"}}
try:
    with contextlib.redirect_stderr(console):
        owner = init_logging(config)
        # Reinitialization must not accumulate a second root bridge.
        owner = init_logging(config)
        for index, name in enumerate(names):
            source = logging.getLogger(name)
            source.setLevel(logging.DEBUG)
            source.warning("unique-warning-%d", index)
        source = logging.getLogger("docwen_runtime.config.loader")
        source.setLevel(logging.ERROR)
        source.warning("level-filtered")
        source.setLevel(logging.DEBUG)
        source.disabled = True
        source.warning("disabled-filtered")
        source.disabled = False
        source.info("sink-filtered")
        for handler in owner.handlers:
            handler.flush()
    state = get_logging_runtime_state()
    text = Path(state.active_log_file).read_text(encoding="utf-8")
    print(json.dumps({"file": text, "console": console.getvalue(), "names": names}))
finally:
    logging.shutdown()
"""
    completed = run_subprocess([sys.executable, "-c", program, str(tmp_path / "logs")], timeout=20)
    assert completed.returncode == 0, completed.stderr
    payload = json.loads(completed.stdout)
    expected = [f"WARNING:{name}:unique-warning-{index}" for index, name in enumerate(payload["names"])]
    assert payload["file"].splitlines() == expected
    assert payload["console"].splitlines() == expected

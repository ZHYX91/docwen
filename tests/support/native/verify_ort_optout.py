"""Explicit Linux diagnostic; never executes ORT with telemetry enabled."""

from __future__ import annotations

import argparse
import base64
import json
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[3]
NATIVE = Path(__file__).resolve().parent
SELF_TEST = r"""
import ctypes
import errno
libc = ctypes.CDLL(None, use_errno=True)
for name in ('fopen', 'fopen64'):
    function = getattr(libc, name)
    function.restype = ctypes.c_void_p
    assert not function(b'/sys/class/drm/card0/device/vendor', b'r')
    assert ctypes.get_errno() == errno.ENOENT
assert libc.socket(2, 1, 0) == -1
assert libc.connect(-1, None, 0) == -1
assert libc.send(-1, None, 0, 0) == -1
assert libc.sendto(-1, None, 0, 0, None, 0) == -1
assert libc.sendmsg(-1, None, 0) == -1
assert libc.getaddrinfo(b'probe.invalid', None, None, None) != 0
"""
ORT_TEST = r"""
import json
import os
import sys
import time
assert os.environ['ORT_DISABLE_TELEMETRY'] == '1'
import onnxruntime as ort
assert ort.__version__ == '1.31.1', ort.__version__
import numpy as np
options = ort.SessionOptions()
options.intra_op_num_threads = 1
options.inter_op_num_threads = 1
session = ort.InferenceSession(sys.argv[1], sess_options=options,
                               providers=['CPUExecutionProvider'])
outputs = session.run(None, {session.get_inputs()[0].name: np.zeros((1, 3, 32, 32), dtype=np.float32)})
assert outputs and all(np.isfinite(value).all() for value in outputs)
time.sleep(3)
print(json.dumps({'ort_version': ort.__version__, 'providers': session.get_providers(),
                  'output_shapes': [list(value.shape) for value in outputs], 'post_run_observation_seconds': 3}))
"""

FROZEN_TEST = r"""
import base64
import json
import os
import sys
from pathlib import Path
assert os.environ['ORT_DISABLE_TELEMETRY'] == '1'
sys.path.insert(0, sys.argv[1])
from scripts.release.verify_packaged_cli import _MACHINE_STDERR_FILENAME, _run_machine_protocol_smoke
work = Path(sys.argv[4])
work.mkdir()
(work / 'config_home').mkdir()
(work / 'log_home').mkdir()
# Only the frozen process gets preloaded, so a load event proves that the
# packaged executable, not just this Python verifier, loaded the fixture.
command = ['/usr/bin/env', 'LD_PRELOAD=' + sys.argv[3], sys.argv[2]]
output = _run_machine_protocol_smoke(command, work_dir=work)
assert output.is_file()
stderr_path = work / _MACHINE_STDERR_FILENAME
stderr = stderr_path.read_bytes() if stderr_path.is_file() else b''
print(json.dumps({'frozen_machine_corpus': 'passed', 'output_name': output.name,
                  'machine_stderr_base64': base64.b64encode(stderr).decode('ascii')}))
"""


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--binary-dir", type=Path, help="Also run the unchanged frozen machine corpus.")
    parser.add_argument("--binary-name", default="DocWenCLI")
    args = parser.parse_args()
    if sys.platform != "linux":
        raise SystemExit("This native diagnostic requires Linux.")
    compiler = shutil.which("cc")
    if not compiler:
        raise SystemExit("A C compiler (cc) is required.")
    model = ROOT / "models/rapidocr/ch_PP-OCRv4_det_infer.onnx"
    if not model.is_file():
        raise SystemExit(f"Missing repository OCR model: {model}")
    with tempfile.TemporaryDirectory(prefix="docwen-ort-optout-") as directory:
        temporary = Path(directory)
        library = temporary / "probe.so"
        subprocess.run(
            [
                compiler,
                "-shared",
                "-fPIC",
                "-Wall",
                "-Wextra",
                "-Werror",
                "-o",
                str(library),
                str(NATIVE / "ort_optout_probe.c"),
                "-ldl",
            ],
            check=True,
            timeout=30,
        )
        environment = dict(os.environ, ORT_DISABLE_TELEMETRY="1", LD_PRELOAD=str(library))
        self_log = temporary / "self-test.log"
        environment["DOCWEN_NATIVE_PROBE_LOG"] = str(self_log)
        subprocess.run([sys.executable, "-c", SELF_TEST], env=environment, check=True, timeout=30)
        events = self_log.read_text().splitlines()
        assert events.count("vendor-ENOENT") == 2, events
        assert "probe-loaded" in events, events
        for name in ("socket", "connect", "send", "sendto", "sendmsg", "getaddrinfo"):
            assert f"network:{name}" in events, events
        ort_log = temporary / "ort.log"
        environment["DOCWEN_NATIVE_PROBE_LOG"] = str(ort_log)
        result = subprocess.run(
            [sys.executable, "-c", ORT_TEST, str(model)],
            env=environment,
            check=True,
            capture_output=True,
            timeout=90,
        )
        events = ort_log.read_text().splitlines()
        assert "probe-loaded" in events, events
        attempts = [event for event in events if event.startswith("network:")]
        assert not attempts, attempts
        report = json.loads(result.stdout)
        report.update(
            native_probe_self_test="passed",
            observed_libc_network_attempts=attempts,
            forced_missing_vendor_reads=events.count("vendor-ENOENT"),
            stderr=result.stderr.decode("utf-8", errors="replace"),
            stderr_base64=base64.b64encode(result.stderr).decode("ascii"),
            scope="Instrumented libc APIs only; not direct-syscall or packet-level coverage.",
        )
        if args.binary_dir is not None:
            binary = args.binary_dir.resolve() / args.binary_name
            if not binary.is_file():
                raise SystemExit(f"Missing frozen CLI: {binary}")
            frozen_log = temporary / "frozen.log"
            frozen_environment = dict(environment, DOCWEN_NATIVE_PROBE_LOG=str(frozen_log))
            frozen_environment.pop("LD_PRELOAD")
            frozen = subprocess.run(
                [sys.executable, "-c", FROZEN_TEST, str(ROOT), str(binary), str(library), str(temporary / "machine")],
                env=frozen_environment,
                check=True,
                capture_output=True,
                timeout=600,
            )
            frozen_events = frozen_log.read_text().splitlines()
            assert "probe-loaded" in frozen_events, frozen_events
            frozen_attempts = [event for event in frozen_events if event.startswith("network:")]
            assert not frozen_attempts, frozen_attempts
            report["frozen"] = dict(
                json.loads(frozen.stdout),
                observed_libc_network_attempts=frozen_attempts,
                forced_missing_vendor_reads=frozen_events.count("vendor-ENOENT"),
                stderr=frozen.stderr.decode("utf-8", errors="replace"),
                verifier_stderr_base64=base64.b64encode(frozen.stderr).decode("ascii"),
            )
        print(json.dumps(report, indent=2))


if __name__ == "__main__":
    main()

"""Build analysis must inherit the native opt-out before importing PyInstaller."""

from __future__ import annotations

import os
import subprocess
import sys
import textwrap
from pathlib import Path

import pytest

pytestmark = [pytest.mark.unit, pytest.mark.skipif(sys.platform != "linux", reason="Linux ORT telemetry policy")]
BUILD = Path(__file__).resolve().parents[2] / "scripts/build/build.py"


@pytest.mark.parametrize("initial", [None, "0", "1"])
@pytest.mark.parametrize("preloaded", [False, True])
def test_build_forces_native_opt_out_before_analysis(initial: str | None, preloaded: bool) -> None:
    script = f"""
        import importlib.abc, os, runpy, subprocess, sys, types
        class AnalysisSentinel(importlib.abc.MetaPathFinder):
            def find_spec(self, fullname, path=None, target=None):
                if fullname.startswith(('onnxruntime', 'rapidocr_onnxruntime')):
                    raise AssertionError('real native import is forbidden in this test')
                if fullname == 'PyInstaller':
                    assert os.environ.get('ORT_DISABLE_TELEMETRY') == '1'
                    child = subprocess.run(
                        [sys.executable, '-I', '-c',
                         "import os; assert os.environ.get('ORT_DISABLE_TELEMETRY') == '1'"],
                        capture_output=True, text=True, timeout=10,
                    )
                    assert child.returncode == 0, child.stderr
                    print('analysis guarded')
                    raise SystemExit(0)
        sys.meta_path.insert(0, AnalysisSentinel())
        if {preloaded!r}:
            sys.modules['onnxruntime'] = types.ModuleType('onnxruntime')
        os.environ.pop('ORT_DISABLE_TELEMETRY', None)
        initial = {initial!r}
        if initial is not None:
            os.environ['ORT_DISABLE_TELEMETRY'] = initial
        sys.path.insert(0, {str(BUILD.parent)!r})
        runpy.run_path({str(BUILD)!r}, run_name='build_test')
        raise AssertionError('PyInstaller analysis boundary was not reached')
    """
    result = subprocess.run(
        [sys.executable, "-I", "-c", textwrap.dedent(script)],
        env={**os.environ, "ORT_DISABLE_TELEMETRY": "1"},
        capture_output=True,
        text=True,
        timeout=30,
        check=False,
    )
    if preloaded and initial != "1":
        assert result.returncode != 0
        assert "onnxruntime_telemetry_bootstrap_started_too_late" in result.stderr
        assert result.stdout == ""
    else:
        assert result.returncode == 0, result.stderr
        assert result.stdout == "analysis guarded\n"
        assert result.stderr == ""

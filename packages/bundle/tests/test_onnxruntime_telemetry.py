"""Native telemetry must be disabled before any DocWen-owned ORT import."""

from __future__ import annotations

import os
import subprocess
import sys
import textwrap
from pathlib import Path

import pytest

pytestmark = pytest.mark.unit
ROOT = Path(__file__).resolve().parents[3]
HOOK = ROOT / "packages/bundle/src/docwen_bundle/pyi_runtime_egress_guard.py"


def _run(script: str) -> subprocess.CompletedProcess[str]:
    env = os.environ.copy()
    # The interpreter itself always starts opted out. Negative cases only use
    # synthetic import sentinels; they must never initialize native ORT.
    env["ORT_DISABLE_TELEMETRY"] = "1"
    env["PYTHONPATH"] = os.pathsep.join(
        [str(path) for path in (ROOT / "packages").rglob("src")] + [env.get("PYTHONPATH", "")]
    )
    return subprocess.run(
        [sys.executable, "-c", textwrap.dedent(script)],
        env=env,
        capture_output=True,
        text=True,
        timeout=30,
        check=False,
    )


@pytest.mark.skipif(sys.platform != "linux", reason="Linux ORT native telemetry policy")
@pytest.mark.parametrize("initial", [None, "", "0", "false", "1"])
@pytest.mark.parametrize(
    "entry", ["docwen_bundle.cli_entry", "docwen_bundle.gui_entry", "docwen_core", "docwen_runtime.capabilities"]
)
def test_entry_forces_opt_out_without_importing_ort(entry: str, initial: str | None) -> None:
    completed = _run(f"""
        import importlib, importlib.abc, os, sys
        class NoNativeImport(importlib.abc.MetaPathFinder):
            def find_spec(self, fullname, path=None, target=None):
                if fullname.startswith(('onnxruntime', 'rapidocr_onnxruntime')):
                    raise AssertionError('unexpected native import: ' + fullname)
        sys.meta_path.insert(0, NoNativeImport())
        os.environ.pop('ORT_DISABLE_TELEMETRY', None)
        initial = {initial!r}
        if initial is not None:
            os.environ['ORT_DISABLE_TELEMETRY'] = initial
        importlib.import_module({entry!r})
        assert os.environ['ORT_DISABLE_TELEMETRY'] == '1'
    """)
    assert completed.returncode == 0, completed.stderr
    assert completed.stdout == completed.stderr == ""


@pytest.mark.skipif(sys.platform != "linux", reason="Linux ORT native telemetry policy")
@pytest.mark.parametrize(
    "statement",
    [
        "import docwen_core.text._table_layout",
        "import docwen_core.text._slanet_table",
        "import docwen_core; os.environ['ORT_DISABLE_TELEMETRY']='0'; import docwen_core.text._table_layout",
        "import docwen_core; os.environ['ORT_DISABLE_TELEMETRY']='0'; import docwen_core.text._slanet_table",
        "import docwen_core; os.environ['ORT_DISABLE_TELEMETRY']='0'; "
        "from docwen_core.text.ocr import _configure_onnxruntime_logging; _configure_onnxruntime_logging()",
        "from docwen_core.text.ocr import _get_ocr_slot; _get_ocr_slot(model_dir='/unused-models')",
        # Exercise reassertion after Core has already been imported.
        "import docwen_core; os.environ['ORT_DISABLE_TELEMETRY']='0'; "
        "from docwen_core.text.ocr import _get_ocr_slot; _get_ocr_slot(model_dir='/unused-models')",
    ],
)
def test_library_import_is_opted_out_before_native_loader(statement: str) -> None:
    completed = _run(f"""
        import importlib.abc, os, sys
        class NativeImportSentinel(importlib.abc.MetaPathFinder):
            def find_spec(self, fullname, path=None, target=None):
                if fullname.startswith(('onnxruntime', 'rapidocr_onnxruntime')):
                    assert os.environ.get('ORT_DISABLE_TELEMETRY') == '1'
                    print('guarded')
                    raise SystemExit(0)
        sys.meta_path.insert(0, NativeImportSentinel())
        os.environ.pop('ORT_DISABLE_TELEMETRY', None)
        {statement}
        raise AssertionError('native import boundary was not reached')
    """)
    assert completed.returncode == 0, completed.stderr
    assert completed.stdout == "guarded\n"
    assert completed.stderr == ""


@pytest.mark.skipif(sys.platform != "linux", reason="Linux ORT native telemetry policy")
@pytest.mark.parametrize(
    "argv",
    [
        ["DocWen"],
        ["DocWen", "--multiprocessing-fork", "parent_pid=1", "pipe_handle=2"],
        ["DocWen", "-c", "from multiprocessing.resource_tracker import main;main(1)"],
        ["DocWen", "-c", "from multiprocessing.forkserver import main;main(1)"],
    ],
)
def test_frozen_hook_forces_opt_out_in_main_and_helpers(argv: list[str]) -> None:
    completed = _run(f"""
        import os, runpy, sys
        sys.argv = {argv!r}
        os.environ['ORT_DISABLE_TELEMETRY'] = '0'
        runpy.run_path({str(HOOK)!r})
        assert os.environ['ORT_DISABLE_TELEMETRY'] == '1'
        assert not any(name.startswith(('onnxruntime', 'rapidocr_onnxruntime')) for name in sys.modules)
    """)
    assert completed.returncode == 0, completed.stderr
    assert completed.stdout == completed.stderr == ""


@pytest.mark.parametrize("module", ["onnxruntime", "onnxruntime.capi", "rapidocr_onnxruntime"])
@pytest.mark.parametrize("helper", [False, True])
def test_frozen_hook_rejects_preloaded_native_dependency(module: str, helper: bool) -> None:
    completed = _run(f"""
        import runpy, sys, types
        if {helper!r}:
            sys.argv = ['DocWen', '--multiprocessing-fork']
        sys.modules[{module!r}] = types.ModuleType({module!r})
        runpy.run_path({str(HOOK)!r})
    """)
    assert completed.returncode != 0
    assert "dependency_egress_guard_started_too_late" in completed.stderr


@pytest.mark.parametrize("platform", ["win32", "darwin"])
def test_runtime_bootstrap_leaves_other_platforms_unchanged(platform: str) -> None:
    completed = _run(f"""
        import os, runpy, sys
        sys.platform = {platform!r}
        os.environ['ORT_DISABLE_TELEMETRY'] = '0'
        runpy.run_path({str(ROOT / "packages/runtime/src/docwen_runtime/__init__.py")!r})
        assert os.environ['ORT_DISABLE_TELEMETRY'] == '0'
    """)
    assert completed.returncode == 0, completed.stderr


@pytest.mark.skipif(sys.platform != "linux", reason="Linux ORT native telemetry policy")
@pytest.mark.parametrize("entry", ["docwen_core", "docwen_runtime"])
@pytest.mark.parametrize("initial", [None, "0", "1"])
def test_source_bootstrap_detects_observably_unsafe_preloaded_ort(entry: str, initial: str | None) -> None:
    completed = _run(f"""
        import importlib, os, sys, types
        sys.modules['onnxruntime'] = types.ModuleType('onnxruntime')
        os.environ.pop('ORT_DISABLE_TELEMETRY', None)
        initial = {initial!r}
        if initial is not None:
            os.environ['ORT_DISABLE_TELEMETRY'] = initial
        importlib.import_module({entry!r})
        assert os.environ['ORT_DISABLE_TELEMETRY'] == '1'
    """)
    if initial == "1":
        assert completed.returncode == 0, completed.stderr
    else:
        assert completed.returncode != 0
        assert "onnxruntime_telemetry_bootstrap_started_too_late" in completed.stderr


@pytest.mark.skipif(sys.platform != "linux", reason="Linux ORT native telemetry policy")
@pytest.mark.parametrize(
    "statement",
    [
        "import docwen_core.text._table_layout",
        "import docwen_core.text._slanet_table",
        "from docwen_core.text.ocr import _configure_onnxruntime_logging; _configure_onnxruntime_logging()",
    ],
)
def test_native_boundary_rejects_late_environment_mutation(statement: str) -> None:
    completed = _run(f"""
        import os, sys, types
        import docwen_core
        sys.modules['onnxruntime'] = types.ModuleType('onnxruntime')
        os.environ['ORT_DISABLE_TELEMETRY'] = '0'
        {statement}
    """)
    assert completed.returncode != 0
    assert "onnxruntime_telemetry_bootstrap_started_too_late" in completed.stderr

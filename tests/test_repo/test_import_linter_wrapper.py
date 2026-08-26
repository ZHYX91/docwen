from __future__ import annotations

import json
import os
import subprocess
import sys
from pathlib import Path

import pytest

pytestmark = pytest.mark.unit

PROJECT_ROOT = Path(__file__).resolve().parents[2]
WRAPPER = PROJECT_ROOT / "tools" / "run_import_linter.py"


def _workspace_src_paths() -> list[str]:
    paths = [str(PROJECT_ROOT / "src")]
    paths.extend(str(src_dir) for src_dir in sorted((PROJECT_ROOT / "packages").glob("**/src")) if src_dir.is_dir())
    return paths


def test_wrapper_calls_console_entrypoint_and_forwards_process_contract(
    tmp_path: Path,
) -> None:
    fake_package = tmp_path / "importlinter"
    fake_package.mkdir()
    (fake_package / "__init__.py").write_text("", encoding="utf-8")
    (fake_package / "cli.py").write_text(
        "from __future__ import annotations\n"
        "\n"
        "import json\n"
        "import os\n"
        "import sys\n"
        "\n"
        "def lint_imports_command(*, prog_name: str) -> None:\n"
        "    print(json.dumps({\n"
        "        'argv': sys.argv[1:],\n"
        "        'cwd': os.getcwd(),\n"
        "        'prog_name': prog_name,\n"
        "        'pythonpath': os.environ['PYTHONPATH'].split(os.pathsep),\n"
        "    }))\n"
        "    raise SystemExit(23)\n",
        encoding="utf-8",
    )
    env = os.environ.copy()
    env["PYTHONPATH"] = str(tmp_path)

    result = subprocess.run(
        [sys.executable, str(WRAPPER), "--contract", "sample", "--verbose"],
        cwd=tmp_path,
        env=env,
        capture_output=True,
        text=True,
        encoding="utf-8",
        timeout=30,
    )

    assert result.returncode == 23, result
    assert result.stderr == ""
    payload = json.loads(result.stdout)
    assert payload == {
        "argv": ["--contract", "sample", "--verbose"],
        "cwd": str(PROJECT_ROOT),
        "prog_name": "lint-imports",
        "pythonpath": [*_workspace_src_paths(), str(tmp_path)],
    }

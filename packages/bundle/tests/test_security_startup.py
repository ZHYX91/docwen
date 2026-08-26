"""Production composition roots must own the egress-guard lifecycle."""

from __future__ import annotations

import json
import subprocess
import sys
from contextlib import contextmanager
from pathlib import Path
from types import SimpleNamespace

import pytest

pytestmark = pytest.mark.unit


@pytest.mark.parametrize("entry_module", ["docwen_bundle.cli_entry", "docwen_bundle.gui_entry"])
def test_composition_entry_imports_security_before_product_modules(entry_module: str) -> None:
    script = (
        "import importlib,json,sys; "
        f"importlib.import_module({entry_module!r}); "
        "blocked=('docwen_core','docwen_application','docwen_cli','docwen_gui','docwen_plugin_'); "
        "print(json.dumps(sorted(name for name in sys.modules if name.startswith(blocked))))"
    )
    completed = subprocess.run(
        [sys.executable, "-c", script],
        check=False,
        capture_output=True,
        text=True,
        timeout=15,
    )

    assert completed.returncode == 0, completed.stderr
    assert json.loads(completed.stdout) == []


def test_pyinstaller_runtime_hook_activates_before_product_imports() -> None:
    script = (
        "import json; "
        "import docwen_bundle.pyi_runtime_egress_guard; "
        "from docwen_runtime.security import dependency_egress_guard_status; "
        "print(json.dumps(dependency_egress_guard_status().to_dict()))"
    )
    completed = subprocess.run(
        [sys.executable, "-c", script],
        check=False,
        capture_output=True,
        text=True,
        timeout=15,
    )

    assert completed.returncode == 0, completed.stderr
    status = json.loads(completed.stdout)
    assert status["state"] == "enforced"
    assert status["bootstrap"] == "pyinstaller_runtime_hook"


def test_pyinstaller_runtime_hook_rejects_late_qt_activation() -> None:
    script = "import PySide6.QtCore; import docwen_bundle.pyi_runtime_egress_guard"
    completed = subprocess.run(
        [sys.executable, "-c", script],
        check=False,
        capture_output=True,
        text=True,
        timeout=15,
    )

    assert completed.returncode != 0
    assert "dependency_egress_guard_started_too_late" in completed.stderr


def test_pyinstaller_runtime_hook_leaves_spawned_helpers_unmanaged() -> None:
    script = (
        "import json,sys; "
        "sys.argv=['DocWen.exe','--multiprocessing-fork','parent_pid=1','pipe_handle=2']; "
        "import docwen_bundle.pyi_runtime_egress_guard; "
        "from docwen_runtime.security import dependency_egress_guard_status; "
        "print(json.dumps(dependency_egress_guard_status().to_dict()))"
    )
    completed = subprocess.run(
        [sys.executable, "-c", script],
        check=False,
        capture_output=True,
        text=True,
        timeout=15,
    )

    assert completed.returncode == 0, completed.stderr
    status = json.loads(completed.stdout)
    assert status["state"] == "not_installed"
    assert status["bootstrap"] == "none"


def test_bundle_owns_the_only_installed_cli_composition_entrypoint() -> None:
    repository = Path(__file__).resolve().parents[3]
    cli_module_entry = repository / "packages" / "apps" / "cli" / "src" / "docwen_cli" / "__main__.py"
    project_config = (repository / "pyproject.toml").read_text(encoding="utf-8")

    assert not cli_module_entry.exists()
    assert 'docwen = "docwen_bundle.cli_entry:main"' in project_config


def test_cli_composition_runs_delegate_while_guard_is_active(monkeypatch: pytest.MonkeyPatch) -> None:
    import docwen_bundle.cli_entry as cli_entry
    import docwen_cli.main as cli_main_module
    from docwen_runtime.security import dependency_egress_guard_status

    observed: dict[str, object] = {}

    def _fake_main(*args, **kwargs) -> int:
        del args, kwargs
        observed.update(dependency_egress_guard_status().to_dict())
        return 17

    monkeypatch.setattr(cli_main_module, "main", _fake_main)

    assert cli_entry.main(["info"]) == 17
    assert observed["state"] == "enforced"
    assert observed["active"] is True


def test_gui_composition_enters_guard_before_bootstrap(monkeypatch: pytest.MonkeyPatch) -> None:
    import docwen_bundle.gui_bootstrap as bootstrap_module
    import docwen_bundle.gui_entry as gui_entry
    from docwen_runtime.security import dependency_egress_guard_status

    observed: dict[str, object] = {}

    def _bootstrap(*, app_name, argv):
        del app_name, argv
        observed.update(dependency_egress_guard_status().to_dict())
        return SimpleNamespace(should_exit=True, exit_code=23, files_to_add=[], instance_lock=None)

    monkeypatch.setattr(bootstrap_module, "bootstrap_gui", _bootstrap)

    assert gui_entry.main(["docwen-gui"]) == 23
    assert observed["state"] == "enforced"
    assert observed["active"] is True


@contextmanager
def _failed_guard():
    from docwen_runtime.security import NetworkGuardInstallationError

    raise NetworkGuardInstallationError()
    yield  # pragma: no cover


def test_cli_composition_fails_closed_with_security_exit_code(monkeypatch: pytest.MonkeyPatch, capsys) -> None:
    import docwen_bundle.cli_entry as cli_entry
    from docwen_cli.exit_codes import ExitCode

    monkeypatch.setattr(cli_entry, "dependency_egress_guard", _failed_guard)

    assert cli_entry.main(["info"]) == int(ExitCode.SECURITY_CHECK_FAILED)
    assert "安全检查失败" in capsys.readouterr().err


def test_gui_composition_fails_closed(monkeypatch: pytest.MonkeyPatch, caplog) -> None:
    import docwen_bundle.gui_entry as gui_entry

    monkeypatch.setattr(gui_entry, "dependency_egress_guard", _failed_guard)

    assert gui_entry.main(["docwen-gui"]) == 1
    assert "安全检查失败" in caplog.text

"""Abrupt compensation death and independent fresh-process config recovery."""

from __future__ import annotations

import hashlib
import json
import os
import stat
from pathlib import Path

import pytest

from ._config_transaction_recovery_support import (
    _canonical_json,
    _config_tree,
    _run_child,
)

pytestmark = pytest.mark.unit


def _fresh_process_config_snapshot(base_dir: Path, user_dir: Path) -> dict[str, object]:
    """Read through a new production loader after the writer has really exited."""
    result = _run_child(
        """
        import json
        import os
        import stat
        import sys
        from pathlib import Path
        from docwen_runtime.config.loader import ConfigLoader
        base, user = map(Path, sys.argv[1:])
        loader = ConfigLoader(base_dir=base, user_dir=user)
        files = {}
        for name in ('gui.toml', 'output.toml'):
            path = user / name
            files[name] = {
                'bytes': path.read_bytes().hex(),
                'mode': stat.S_IMODE(path.stat().st_mode),
                'link': os.readlink(path) if path.is_symlink() else None,
            }
        snapshot = {
            'trusted': loader.config_state_trusted,
            'values': [loader.config.gui.window.default_mode, loader.config.output.directory.mode],
            'files': files,
            'journal': (user / '.docwen-config.transaction.json').exists(),
        }
        (base / 'fresh-config-observation.json').write_text(json.dumps(snapshot), encoding='utf-8')
        """,
        base_dir,
        user_dir,
    )
    assert result.returncode == 0, result.stderr
    return json.loads((base_dir / "fresh-config-observation.json").read_text(encoding="utf-8"))


@pytest.mark.parametrize("logical_symlink", [False, True], ids=["regular", "logical-symlink"])
@pytest.mark.parametrize("recovery_failure", [False, True], ids=["recover", "retry-recovery"])
def test_commit_replace_barrier_failure_then_compensation_death_never_trusts_mixed_generation(
    tmp_path: Path, logical_symlink: bool, recovery_failure: bool
) -> None:
    """Actual COMMITTED replacement, failed barrier, restored first file, hard exit."""
    import base64
    import tomllib

    if logical_symlink and os.name == "nt":
        pytest.skip("This source regression does not require native Windows symlink privileges")
    base_dir, user_dir, preimages = _config_tree(tmp_path)
    gui = user_dir / "gui.toml"
    if logical_symlink:
        target = tmp_path / "gui-actual.toml"
        gui.rename(target)
        gui.symlink_to(Path("..") / target.name)
    expected_files = {}
    for name in ("gui.toml", "output.toml"):
        path = user_dir / name
        path.chmod(0o640)
        expected_files[name] = {
            "bytes": preimages[name].hex(),
            "mode": stat.S_IMODE(path.stat().st_mode),
            "link": os.readlink(path) if path.is_symlink() else None,
        }
    writer = _run_child(
        """
        import json
        import os
        import sys
        from pathlib import Path
        from docwen_runtime import toml_io
        from docwen_runtime.config import loader as module
        base, user = map(Path, sys.argv[1:])
        loader = module.ConfigLoader(base_dir=base, user_dir=user)
        journal = user / '.docwen-config.transaction.json'
        original_replace = toml_io.os.replace
        original_sync = toml_io._sync_directory
        original_restore = module._restore_user_file_preimage
        phase = {'pending': None, 'committed_replaced': False, 'rollback_durable': False}
        def replace(source, destination):
            original_replace(source, destination)
            if Path(destination) == journal:
                state = json.loads(journal.read_bytes())['payload']['state']
                if state == 'COMMITTED':
                    phase['committed_replaced'] = True
                    phase['pending'] = 'commit'
                    phase['committed_bytes'] = {
                        name: (user / name).read_bytes().hex() for name in ('gui.toml', 'output.toml')
                    }
                elif phase['committed_replaced']:
                    phase['pending'] = 'rollback'
        def sync(path):
            pending = phase.pop('pending', None)
            if pending == 'commit':
                raise OSError('actual COMMITTED replace succeeded; directory barrier failed')
            original_sync(path)
            if pending == 'rollback':
                phase['rollback_durable'] = True
        def restore(before, *, operation):
            original_restore(before, operation=operation)
            phase['restored'] = before.path.name
            phase['journal'] = json.loads(journal.read_bytes())
            phase['bytes_at_death'] = {
                name: (user / name).read_bytes().hex() for name in ('gui.toml', 'output.toml')
            }
            (base / 'writer-observation.json').write_text(json.dumps(phase), encoding='utf-8')
            os._exit(97)
        toml_io.os.replace = replace
        toml_io._sync_directory = sync
        module._restore_user_file_preimage = restore
        loader.set_values({'gui.window.default_mode': 'single', 'output.directory.mode': 'source'})
        raise AssertionError('the writer did not reach first compensation')
        """,
        base_dir,
        user_dir,
    )
    assert writer.returncode == 97, writer.stderr
    observation = json.loads((base_dir / "writer-observation.json").read_text(encoding="utf-8"))
    assert observation["committed_replaced"] is True
    committed = {
        name: tomllib.loads(bytes.fromhex(raw).decode("utf-8")) for name, raw in observation["committed_bytes"].items()
    }
    assert committed["gui.toml"]["window"]["default_mode"] == "single"
    assert committed["output.toml"]["directory"]["mode"] == "source"
    assert observation["restored"] == "output.toml"
    assert observation["bytes_at_death"]["output.toml"] == preimages["output.toml"].hex()
    assert observation["bytes_at_death"]["gui.toml"] != preimages["gui.toml"].hex()
    if recovery_failure:
        recovery = _run_child(
            """
            import json
            import sys
            from pathlib import Path
            from docwen_runtime.config import loader as module
            base, user = map(Path, sys.argv[1:])
            original_restore = module._restore_user_file_preimage
            def restore(before, *, operation):
                original_restore(before, operation=operation)
                raise OSError('recovery interrupted after real first restore')
            module._restore_user_file_preimage = restore
            try:
                module.ConfigLoader(base_dir=base, user_dir=user)
            except OSError as exc:
                (base / 'recovery-failure.json').write_text(json.dumps({'error': str(exc)}), encoding='utf-8')
            else:
                raise AssertionError('partial recovery was trusted')
            """,
            base_dir,
            user_dir,
        )
        assert recovery.returncode == 0, recovery.stderr
        assert json.loads((base_dir / "recovery-failure.json").read_text(encoding="utf-8"))["error"] == (
            "recovery interrupted after real first restore"
        )
        assert json.loads((user_dir / ".docwen-config.transaction.json").read_bytes())["payload"]["state"] == (
            "PREPARED"
        )
        assert (user_dir / "gui.toml").read_bytes() != preimages["gui.toml"]
    recovered = _fresh_process_config_snapshot(base_dir, user_dir)
    assert recovered == {
        "trusted": True,
        "values": ["batch", "custom"],
        "files": expected_files,
        "journal": False,
    }
    assert observation["rollback_durable"] is True
    assert observation["journal"]["payload"]["state"] == "PREPARED"
    preimage_rows = observation["journal"]["payload"]["preimages"]
    assert [item["path"] for item in preimage_rows] == ["gui.toml", "output.toml"]
    assert preimage_rows[0]["symlink_target"] == ("../gui-actual.toml" if logical_symlink else None)
    for row in preimage_rows:
        assert base64.b64decode(row["content"], validate=True) == preimages[row["path"]]
        assert row["mode"] == expected_files[row["path"]]["mode"]
    assert (
        hashlib.sha256(_canonical_json(observation["journal"]["payload"])).hexdigest()
        == (observation["journal"]["checksum"])
    )

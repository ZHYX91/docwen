"""Regression tests for batch destination planning after path expansion."""

from __future__ import annotations

import argparse
from pathlib import Path

import pytest

from docwen_cli.commands.execution_v3 import _preflight_batch_collisions

pytestmark = pytest.mark.unit


def test_batch_collision_preflight_expands_directory_inputs(tmp_path: Path) -> None:
    source_root = tmp_path / "inputs"
    blue = source_root / "blue"
    red = source_root / "red"
    blue.mkdir(parents=True)
    red.mkdir(parents=True)
    (blue / "same.png").write_bytes(b"blue")
    (red / "same.png").write_bytes(b"red")
    output = tmp_path / "output"
    output.mkdir()

    args = argparse.Namespace(
        command_path="batch convert",
        files=[str(source_root)],
        to="jpg",
        overwrite=True,
    )

    rejection = _preflight_batch_collisions(args, output)

    assert rejection is not None
    code, _message, details = rejection
    assert code == "output_collision"
    assert len(details["collisions"]) == 1
    assert {Path(path).parent.name for path in details["collisions"][0]} == {"blue", "red"}

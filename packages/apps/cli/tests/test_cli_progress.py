"""CLI progress rendering contracts."""

from __future__ import annotations

import pytest
from _pytest.capture import CaptureFixture

from docwen_cli.utils import create_progress_callback

pytestmark = pytest.mark.unit


def test_progress_callback_writes_normal_message_to_stderr(capsys: CaptureFixture[str]) -> None:
    callback = create_progress_callback()

    callback("Converting page 1")

    captured = capsys.readouterr()
    assert captured.out == ""
    assert captured.err == "Converting page 1\n"


def test_progress_callback_adds_localized_prefix_in_verbose_mode(capsys: CaptureFixture[str]) -> None:
    callback = create_progress_callback(verbose=True)

    callback("Converting page 1")

    captured = capsys.readouterr()
    assert captured.out == ""
    assert captured.err.endswith(" Converting page 1\n")
    assert captured.err != "Converting page 1\n"


@pytest.mark.parametrize("options", [{"quiet": True}, {"json_mode": True}])
def test_progress_callback_is_silent_for_machine_readable_modes(
    capsys: CaptureFixture[str],
    options: dict[str, bool],
) -> None:
    callback = create_progress_callback(**options)

    callback("must stay hidden")

    captured = capsys.readouterr()
    assert captured.out == ""
    assert captured.err == ""

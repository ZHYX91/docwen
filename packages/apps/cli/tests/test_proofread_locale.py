"""Explicit CLI language becomes a request fact only on locale-aware routes."""

from __future__ import annotations

from argparse import Namespace

import pytest

from docwen_cli.commands.execution_options import build_execution_options

pytestmark = pytest.mark.unit


@pytest.mark.parametrize("command", ["validate", "convert"])
def test_explicit_cli_language_reaches_locale_aware_route(command):
    args = Namespace(command=command, lang="fr_FR")
    assert build_execution_options(args, route_options={"locale"})["locale"] == "fr_FR"
    assert "locale" not in build_execution_options(args, route_options=set())

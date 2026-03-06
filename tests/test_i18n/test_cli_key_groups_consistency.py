"""i18n 单元测试。"""

from __future__ import annotations

from pathlib import Path

import pytest
import tomlkit

from docwen.i18n.locale_validator import get_all_locale_paths, get_reference_locale_path

pytestmark = pytest.mark.unit


def _load(path: Path):
    return tomlkit.parse(path.read_text(encoding="utf-8"))


def _get_table(doc: tomlkit.TOMLDocument, path: str):
    cur = doc
    for part in path.split("."):
        cur = cur.get(part, {})
    return cur


@pytest.mark.parametrize(
    "table_path",
    [
        "cli.groups",
        "cli.help",
        "cli.messages",
        "cli.batch",
        "cli.categories",
        "cli.prompts",
        "cli.validation",
        "cli.interactive",
        "cli.interactive.menus",
        "cli.interactive.formats",
        "cli.interactive.dpi",
        "cli.interactive.compress",
        "cli.interactive.pdf_sizes",
        "cli.interactive.numbering_schemes",
        "cli.interactive.optimization_types",
        "cli.interactive.proofread",
        "cli.interactive.prompts",
    ],
)
def test_cli_table_keys_are_consistent_across_locales(table_path: str) -> None:
    ref_path = get_reference_locale_path()
    ref_doc = _load(ref_path)
    ref_table = _get_table(ref_doc, table_path)
    ref_keys = {str(k) for k in ref_table}
    assert ref_keys, f"Reference locale missing [{table_path}] keys"

    for path in get_all_locale_paths():
        doc = _load(path)
        table = _get_table(doc, table_path)
        keys = {str(k) for k in table}
        missing = sorted(ref_keys - keys)
        assert not missing, f"Missing [{table_path}] keys in {path.name}: {missing}"

"""i18n 单元测试。"""

from __future__ import annotations

import tomllib
from functools import cache
from pathlib import Path

import pytest

pytestmark = pytest.mark.unit


@cache
def _load(path: Path):
    """Parse each immutable shipped locale once per pytest worker."""
    return tomllib.loads(path.read_text(encoding="utf-8"))


def _get_table(doc: dict, path: str):
    cur = doc
    for part in path.split("."):
        cur = cur.get(part, {})
    return cur


PROJECT_ROOT = Path(__file__).resolve().parents[2]
LOCALES_DIR = PROJECT_ROOT / "i18n" / "locales"


def get_reference_locale_path() -> Path:
    return LOCALES_DIR / "zh_CN.toml"


def get_all_locale_paths() -> list[Path]:
    return sorted(LOCALES_DIR.glob("*.toml"))


@pytest.mark.parametrize(
    "table_path",
    [
        "cli.help",
        "cli.messages",
        "cli.interactive",
        "cli.interactive.optimization_types",
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


@pytest.mark.parametrize(
    "table_path",
    [
        "cli.groups",
        "cli.batch",
        "cli.prompts",
        "cli.validation",
        "cli.interactive.menus",
        "cli.interactive.formats",
        "cli.interactive.dpi",
        "cli.interactive.compress",
        "cli.interactive.pdf_sizes",
        "cli.interactive.numbering_schemes",
        "cli.interactive.proofread",
        "cli.interactive.prompts",
    ],
)
def test_retired_cli_tables_have_no_locale_keys(table_path: str) -> None:
    """Retired interactive UI catalogs cannot become their own usage proof."""
    for path in get_all_locale_paths():
        doc = _load(path)
        assert not _get_table(doc, table_path), f"Retired [{table_path}] keys remain in {path.name}"


def test_obsolete_cli_categories_table_is_retired_from_every_locale() -> None:
    """Category identity comes from Core, not an incomplete locale-only catalog."""
    for path in get_all_locale_paths():
        doc = _load(path)
        cli = doc.get("cli", {})
        assert "categories" not in cli, f"Obsolete [cli.categories] remains in {path.name}"

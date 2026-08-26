"""Fail-closed i18n catalogue reference governance."""

from __future__ import annotations

from functools import cache
from pathlib import Path
from textwrap import dedent

import pytest
from tests.support.i18n_reference_audit import (
    STRUCTURAL_SECTIONS,
    DynamicCallContract,
    LiteralFallbackContract,
    audit_locale_references,
    defined_locale_keys,
    locale_leaf_keys,
    semantic_numbering_keys,
)

pytestmark = pytest.mark.unit


PROJECT_ROOT = Path(__file__).resolve().parents[2]
LOCALES_DIR = PROJECT_ROOT / "i18n" / "locales"


@cache
def _production_audit():
    """Scan the immutable shipped source tree once per pytest worker."""
    return audit_locale_references(PROJECT_ROOT, locale_dir=LOCALES_DIR)


def _write_fixture_locale(root: Path, body: str) -> Path:
    locale_dir = root / "i18n" / "locales"
    locale_dir.mkdir(parents=True)
    (locale_dir / "zh_CN.toml").write_text(dedent(body).strip() + "\n", encoding="utf-8")
    return locale_dir


def _write_gui_source(root: Path, source: str, *, name: str = "sample.py") -> Path:
    src_dir = root / "packages" / "apps" / "gui" / "src"
    src_dir.mkdir(parents=True)
    path = src_dir / name
    path.write_text(dedent(source).strip() + "\n", encoding="utf-8")
    return path


def test_all_packaged_locale_leaf_key_sets_match_reference() -> None:
    reference = locale_leaf_keys(LOCALES_DIR / "zh_CN.toml")
    for path in sorted(LOCALES_DIR.glob("*.toml")):
        assert locale_leaf_keys(path) == reference, f"{path.name} leaf keys differ from zh_CN.toml"


def test_defined_key_extraction_excludes_structural_tables() -> None:
    all_leaves = locale_leaf_keys(LOCALES_DIR / "zh_CN.toml")
    expected = {key for key in all_leaves if key.split(".", 1)[0] not in STRUCTURAL_SECTIONS}
    defined = defined_locale_keys(LOCALES_DIR / "zh_CN.toml")

    assert defined == expected
    assert "action_area.document.export_markdown" in defined
    assert "placeholders.title" in all_leaves
    assert "placeholders.title" not in defined


def test_production_catalogue_has_only_live_or_finite_declarative_keys() -> None:
    audit = _production_audit()

    assert audit.unresolved == ()
    assert audit.undefined_literal_keys == ()
    assert audit.contract_mismatches == ()
    assert audit.undefined_contract_keys == frozenset()
    assert audit.unused == frozenset()


def test_production_scan_is_limited_to_shipped_package_source() -> None:
    audit = _production_audit()

    assert audit.source_files
    assert audit.source_roots
    assert all("/src/" in f"/{path}" for path in audit.source_files)
    assert all(path.startswith("packages/") and path.endswith("/src") for path in audit.source_roots)
    banned_parts = {".tmp", ".pytest_cache", "build", "dist", "tmp", "tests", "docs", "scripts"}
    assert all(not (set(Path(path).parts) & banned_parts) for path in audit.source_files)


def test_import_alias_is_recognized_but_arbitrary_constants_and_method_calls_are_not(tmp_path: Path) -> None:
    locale_dir = _write_fixture_locale(
        tmp_path,
        """
        [ui]
        live = "Live"
        orphan = "Orphan"
        """,
    )
    _write_gui_source(
        tmp_path,
        """
        from docwen_gui.i18n import t as translate

        MESSAGE_TABLE = ("ui.orphan",)
        translate("ui.live")
        object().t("ui.orphan")
        """,
    )
    (tmp_path / "tests").mkdir()
    (tmp_path / "tests" / "fixture.py").write_text('QUOTE = "ui.orphan"\n', encoding="utf-8")
    (tmp_path / "dist").mkdir()
    (tmp_path / "dist" / "generated.json").write_text('{"label": "ui.orphan"}\n', encoding="utf-8")

    audit = audit_locale_references(tmp_path, locale_dir=locale_dir, contracts={}, literal_contracts={})

    assert audit.used == {"ui.live"}
    assert audit.unused == {"ui.orphan"}
    assert audit.unresolved == ()


def test_relative_import_alias_is_recognized(tmp_path: Path) -> None:
    locale_dir = _write_fixture_locale(tmp_path, '[ui]\nlive = "Live"')
    _write_gui_source(tmp_path, 'from .i18n import t as tr\ntr("ui.live")')

    audit = audit_locale_references(tmp_path, locale_dir=locale_dir, contracts={}, literal_contracts={})

    assert audit.used == {"ui.live"}
    assert audit.unused == frozenset()


def test_missing_literal_translation_key_fails_closed(tmp_path: Path) -> None:
    locale_dir = _write_fixture_locale(tmp_path, '[ui]\none = "One"')
    _write_gui_source(
        tmp_path,
        """
        from docwen_gui.i18n import t

        t("ui.one")
        t("ui.missing", "Fallback")
        """,
    )

    audit = audit_locale_references(tmp_path, locale_dir=locale_dir, contracts={}, literal_contracts={})

    assert audit.used == {"ui.one"}
    assert audit.unused == frozenset()
    assert audit.undefined_literal_keys == ("packages/apps/gui/src/sample.py:4: ui.missing",)
    assert audit.unresolved == ()


def test_keyword_key_argument_cannot_bypass_literal_key_audit(tmp_path: Path) -> None:
    locale_dir = _write_fixture_locale(tmp_path, '[ui]\none = "One"')
    _write_gui_source(
        tmp_path,
        """
        from docwen_gui.i18n import t

        t(key="ui.one")
        t(key="ui.missing", default="Fallback")
        """,
    )

    audit = audit_locale_references(tmp_path, locale_dir=locale_dir, contracts={}, literal_contracts={})

    assert audit.used == {"ui.one"}
    assert audit.unused == frozenset()
    assert audit.undefined_literal_keys == ("packages/apps/gui/src/sample.py:4: ui.missing",)
    assert audit.unresolved == ()


@pytest.mark.parametrize(
    ("call", "reason"),
    [
        ('t("ui.one", "Fallback", "unexpected")', "more than two positional arguments"),
        ('t("ui.one", *extra)', "starred positional arguments"),
        ('t("ui.one", key="ui.one")', "multiple values for key"),
        ('t("ui.one", "Fallback", default="Other")', "multiple values for default"),
    ],
)
def test_invalid_translation_argument_binding_fails_closed(tmp_path: Path, call: str, reason: str) -> None:
    locale_dir = _write_fixture_locale(tmp_path, '[ui]\none = "One"')
    _write_gui_source(
        tmp_path,
        f"from docwen_gui.i18n import t\n\nextra = ('unexpected',)\n{call}",
    )

    audit = audit_locale_references(tmp_path, locale_dir=locale_dir, contracts={}, literal_contracts={})

    assert audit.used == frozenset()
    assert audit.unused == {"ui.one"}
    assert len(audit.unresolved) == 1
    assert reason in audit.unresolved[0]


def test_literal_fallback_contract_rejects_extra_positional_argument(tmp_path: Path) -> None:
    locale_dir = _write_fixture_locale(tmp_path, '[ui]\none = "One"')
    source = _write_gui_source(
        tmp_path,
        """
        from docwen_gui.i18n import t

        t("ui.fallback_only", "Reviewed fallback", "unexpected")
        """,
    )
    signature = (source.relative_to(tmp_path).as_posix(), "ui.fallback_only")
    literal_contracts = {
        signature: LiteralFallbackContract(
            expected_count=1,
            default="Reviewed fallback",
            rationale="fixture fallback",
        )
    }

    audit = audit_locale_references(
        tmp_path,
        locale_dir=locale_dir,
        contracts={},
        literal_contracts=literal_contracts,
    )

    assert len(audit.unresolved) == 1
    assert "more than two positional arguments" in audit.unresolved[0]
    assert audit.contract_mismatches == (
        "packages/apps/gui/src/sample.py :: literal ui.fallback_only expected 1, observed 0",
    )


@pytest.mark.parametrize(
    "call",
    [
        't("ui.fallback_only", "Reviewed fallback")',
        't(key="ui.fallback_only", default="Reviewed fallback")',
    ],
)
def test_exact_literal_fallback_contract_is_counted_and_default_bound(tmp_path: Path, call: str) -> None:
    locale_dir = _write_fixture_locale(tmp_path, '[ui]\none = "One"')
    source = _write_gui_source(
        tmp_path,
        f"from docwen_gui.i18n import t\n\n{call}",
    )
    signature = (source.relative_to(tmp_path).as_posix(), "ui.fallback_only")
    literal_contracts = {
        signature: LiteralFallbackContract(
            expected_count=1,
            default="Reviewed fallback",
            rationale="fixture fallback",
        )
    }

    audit = audit_locale_references(
        tmp_path,
        locale_dir=locale_dir,
        contracts={},
        literal_contracts=literal_contracts,
    )

    assert audit.undefined_literal_keys == ()
    assert audit.contract_mismatches == ()


def test_unknown_dynamic_translation_expression_fails_closed(tmp_path: Path) -> None:
    locale_dir = _write_fixture_locale(tmp_path, '[ui]\none = "One"\ntwo = "Two"')
    _write_gui_source(
        tmp_path,
        """
        from docwen_gui.i18n import t

        def label(name: str) -> str:
            return t(f"ui.{name}")
        """,
    )

    audit = audit_locale_references(tmp_path, locale_dir=locale_dir, contracts={}, literal_contracts={})

    assert audit.used == frozenset()
    assert audit.unused == {"ui.one", "ui.two"}
    assert len(audit.unresolved) == 1
    assert "t(f'ui.{name}')" in audit.unresolved[0]


def test_exact_dynamic_contract_has_no_prefix_wildcard_or_silent_growth(tmp_path: Path) -> None:
    locale_dir = _write_fixture_locale(tmp_path, '[ui]\none = "One"\ntwo = "Two"')
    source = _write_gui_source(
        tmp_path,
        """
        from docwen_gui.i18n import t

        def label(name: str) -> str:
            return t(f"ui.{name}")
        """,
    )
    signature = (source.relative_to(tmp_path).as_posix(), "f'ui.{name}'")
    contracts = {
        signature: DynamicCallContract(expected_count=1, keys=frozenset({"ui.one"}), rationale="fixture domain")
    }

    audit = audit_locale_references(
        tmp_path,
        locale_dir=locale_dir,
        contracts=contracts,
        literal_contracts={},
    )

    assert audit.used == {"ui.one"}
    assert audit.unused == {"ui.two"}
    assert audit.contract_mismatches == ()


def test_stale_dynamic_contract_fails_when_source_call_disappears(tmp_path: Path) -> None:
    locale_dir = _write_fixture_locale(tmp_path, '[ui]\none = "One"')
    source = _write_gui_source(tmp_path, "from docwen_gui.i18n import t")
    signature = (source.relative_to(tmp_path).as_posix(), "f'ui.{name}'")
    contracts = {
        signature: DynamicCallContract(expected_count=1, keys=frozenset({"ui.one"}), rationale="fixture domain")
    }

    audit = audit_locale_references(
        tmp_path,
        locale_dir=locale_dir,
        contracts=contracts,
        literal_contracts={},
    )

    assert len(audit.contract_mismatches) == 1
    assert "expected 1, observed 0" in audit.contract_mismatches[0]


def test_numbering_toml_reads_only_reviewed_semantic_fields(tmp_path: Path) -> None:
    locale_dir = _write_fixture_locale(
        tmp_path,
        """
        [editors.numbering_add.names]
        scheme = "Scheme"
        decoy = "Decoy"
        [editors.numbering_add.descriptions]
        scheme_desc = "Description"
        [editors.numbering_clean.names]
        clean = "Clean"
        [editors.numbering_clean.descriptions]
        clean_desc = "Clean description"
        [ui]
        decoy = "Decoy"
        """,
    )
    add_path = tmp_path / "configs" / "numbering" / "add.toml"
    add_path.parent.mkdir(parents=True)
    add_path.write_text(
        """
        [schemes.demo]
        name_key = "scheme"
        description_key = "scheme_desc"
        description = "ui.decoy"
        """.strip()
        + "\n",
        encoding="utf-8",
    )
    cleanup_path = tmp_path / "configs" / "numbering" / "cleanup.toml"
    cleanup_path.write_text(
        """
        [[rules]]
        name_key = "clean"
        description_key = "clean_desc"
        pattern = "ui.decoy"
        """.strip()
        + "\n",
        encoding="utf-8",
    )
    (tmp_path / "configs" / "labels.toml").write_text('label_key = "ui.decoy"\n', encoding="utf-8")

    used = semantic_numbering_keys(tmp_path)
    audit = audit_locale_references(tmp_path, locale_dir=locale_dir, contracts={}, literal_contracts={})

    assert used == {
        "editors.numbering_add.names.scheme",
        "editors.numbering_add.descriptions.scheme_desc",
        "editors.numbering_clean.names.clean",
        "editors.numbering_clean.descriptions.clean_desc",
    }
    assert audit.unused == {"editors.numbering_add.names.decoy", "ui.decoy"}

from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest
from PySide6.QtWidgets import QTableWidgetItem

from docwen_core.toml_tools import toml_table, toml_value

pytestmark = pytest.mark.gui


def _write_minimal_base_config_tree(base_dir: Path) -> None:
    """Create an empty TOML file for every spec in the registry."""
    from docwen_runtime.config.registry import CONFIG_FILES

    for spec in CONFIG_FILES:
        path = base_dir / spec.rel_path
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text("\n", encoding="utf-8")


def _visible_rows(table) -> list[int]:
    return [row for row in range(table.rowCount()) if not table.isRowHidden(row)]


# ── Symbol Mapping Editor ───────────────────────────────────────────────────


def test_symbol_mapping_editor_preserves_pairs(
    qapp,
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """_SymbolPairingEditor reads/writes proofread/pairs.toml with key 'items'."""
    from docwen_gui.widgets.settings.proofread_tab import _SymbolPairingEditor
    from docwen_runtime.config.loader import ConfigLoader

    config_dir = tmp_path / "configs"
    config_dir.mkdir(parents=True, exist_ok=True)
    _write_minimal_base_config_tree(config_dir)
    loader = ConfigLoader(base_dir=config_dir, user_dir=config_dir)

    def _seed_pairs(doc) -> None:
        doc["items"] = [["<", ">"], ["【", "】"]]

    assert loader.update_file_sections(
        "proofread/pairs.toml",
        {"items": [["<", ">"], ["【", "】"]]},
    )

    dialog = _SymbolPairingEditor(config_dir / "proofread" / "pairs.toml")

    assert dialog._table.columnCount() == 2
    item00 = dialog._table.item(0, 0)
    assert item00 is not None
    assert item00.text() == "<"
    item01 = dialog._table.item(0, 1)
    assert item01 is not None
    assert item01.text() == ">"
    item10 = dialog._table.item(1, 0)
    assert item10 is not None
    assert item10.text() == "【"

    dialog._on_accept()

    saved = loader.get_file_dict("proofread/pairs.toml")
    items = saved.get("items", [])
    assert ["<", ">"] in items
    assert ["【", "】"] in items

    dialog.close()


# ── Symbol Error Editor ─────────────────────────────────────────────────────


def test_symbol_error_editor_preserves_entries_and_remark(
    qapp,
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """_SymbolErrorEditor reads/writes proofread/symbol_map.toml with key 'entries'."""
    from docwen_gui.widgets.settings.proofread_tab import _SymbolErrorEditor
    from docwen_runtime.config.loader import ConfigLoader

    config_dir = tmp_path / "configs"
    config_dir.mkdir(parents=True, exist_ok=True)
    _write_minimal_base_config_tree(config_dir)
    loader = ConfigLoader(base_dir=config_dir, user_dir=config_dir)

    def _seed_symbol_errors(doc) -> None:
        table = toml_table()
        table["0"] = toml_value(["０", "○"], "全角数字和圆圈")
        doc["entries"] = table

    assert loader.update_file_document(
        "proofread/symbol_map.toml",
        _seed_symbol_errors,
    )

    dialog = _SymbolErrorEditor(config_dir / "proofread" / "symbol_map.toml")

    assert dialog._table.columnCount() == 3
    item00 = dialog._table.item(0, 0)
    assert item00 is not None
    assert item00.text() == "0"
    item01 = dialog._table.item(0, 1)
    assert item01 is not None
    assert item01.text() == "０|○"
    item02 = dialog._table.item(0, 2)
    assert item02 is not None
    assert item02.text() == "全角数字和圆圈"

    dialog._on_accept()

    saved = loader.get_file_dict("proofread/symbol_map.toml")
    assert saved.get("entries", {}).get("0") == ["０", "○"]
    saved_text = (config_dir / "proofread" / "symbol_map.toml").read_text(encoding="utf-8")
    assert "全角数字和圆圈" in saved_text

    dialog.close()


# ── Typos Dictionary Editor ──────────────────────────────────────────────────


def test_typos_editor_preserves_entries_and_remark(
    qapp,
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """_TyposDictionaryEditor reads/writes proofread/typos.toml with key 'entries'."""
    from docwen_gui.widgets.settings.proofread_tab import _TyposDictionaryEditor
    from docwen_runtime.config.loader import ConfigLoader

    config_dir = tmp_path / "configs"
    config_dir.mkdir(parents=True, exist_ok=True)
    _write_minimal_base_config_tree(config_dir)
    loader = ConfigLoader(base_dir=config_dir, user_dir=config_dir)

    def _seed_typos(doc) -> None:
        table = toml_table()
        table["己"] = toml_value(["已"], "常见形近错别字")
        doc["entries"] = table

    assert loader.update_file_document(
        "proofread/typos.toml",
        _seed_typos,
    )

    dialog = _TyposDictionaryEditor(config_dir / "proofread" / "typos.toml")

    assert dialog._table.columnCount() == 3
    from docwen_gui.i18n import t

    correct_header = dialog._table.horizontalHeaderItem(0)
    misspellings_header = dialog._table.horizontalHeaderItem(1)
    assert correct_header is not None and correct_header.text() == t("editors.mapping.correct_word")
    assert misspellings_header is not None and misspellings_header.text() == t("editors.mapping.misspellings")
    item00 = dialog._table.item(0, 0)
    assert item00 is not None
    assert item00.text() == "己"
    item01 = dialog._table.item(0, 1)
    assert item01 is not None
    assert item01.text() == "已"
    item02 = dialog._table.item(0, 2)
    assert item02 is not None
    assert item02.text() == "常见形近错别字"

    dialog._on_accept()

    saved = loader.get_file_dict("proofread/typos.toml")
    assert saved.get("entries", {}).get("己") == ["已"]
    saved_text = (config_dir / "proofread" / "typos.toml").read_text(encoding="utf-8")
    assert "常见形近错别字" in saved_text

    dialog.close()


def test_typos_editor_search_filters_rows(
    qapp,
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_gui.widgets.settings.proofread_tab import _TyposDictionaryEditor
    from docwen_runtime.config.loader import ConfigLoader

    config_dir = tmp_path / "configs"
    config_dir.mkdir(parents=True, exist_ok=True)
    _write_minimal_base_config_tree(config_dir)
    loader = ConfigLoader(base_dir=config_dir, user_dir=config_dir)

    def _seed_typos(doc) -> None:
        table = toml_table()
        table["己"] = toml_value(["已"], "常见形近错别字")
        table["登陆"] = toml_value(["登录"], "术语统一")
        doc["entries"] = table

    assert loader.update_file_document(
        "proofread/typos.toml",
        _seed_typos,
    )

    dialog = _TyposDictionaryEditor(config_dir / "proofread" / "typos.toml")

    assert dialog._table.rowCount() == 2
    dialog._search_box.setText("术语")  # pyright: ignore[reportPrivateUsage]
    qapp.processEvents()

    assert _visible_rows(dialog._table) == [1]

    dialog.close()


def test_typos_editor_paginates_and_search_resets_to_first_page(
    qapp,
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_gui.widgets.settings.proofread_tab import _TyposDictionaryEditor
    from docwen_runtime.config.loader import ConfigLoader

    config_dir = tmp_path / "configs"
    config_dir.mkdir(parents=True, exist_ok=True)
    _write_minimal_base_config_tree(config_dir)
    loader = ConfigLoader(base_dir=config_dir, user_dir=config_dir)

    def _seed_typos(doc) -> None:
        table = toml_table()
        for index in range(25):
            key = f"K{index:02d}"
            table[key] = toml_value([f"V{index:02d}"], f"comment {index:02d}")
        doc["entries"] = table

    assert loader.update_file_document(
        "proofread/typos.toml",
        _seed_typos,
    )

    dialog = _TyposDictionaryEditor(config_dir / "proofread" / "typos.toml")

    assert dialog._table.rowCount() == 25
    assert _visible_rows(dialog._table) == list(range(20))
    assert dialog._page_label.text() == "1 / 2"  # pyright: ignore[reportPrivateUsage]

    dialog._go_to_next_page()  # pyright: ignore[reportPrivateUsage]

    assert _visible_rows(dialog._table) == list(range(20, 25))
    assert dialog._page_label.text() == "2 / 2"  # pyright: ignore[reportPrivateUsage]

    dialog._add_row()  # pyright: ignore[reportPrivateUsage]

    assert dialog._table.rowCount() == 26
    assert _visible_rows(dialog._table) == list(range(20, 26))
    assert dialog._page_label.text() == "2 / 2"  # pyright: ignore[reportPrivateUsage]

    dialog._search_box.setText("K03")  # pyright: ignore[reportPrivateUsage]
    qapp.processEvents()

    assert dialog._page_label.text() == "1 / 1"  # pyright: ignore[reportPrivateUsage]
    assert _visible_rows(dialog._table) == [3]

    dialog.close()


def test_typos_editor_duplicate_key_merge_strategy_preserves_old_project_behavior(
    qapp,
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_gui.widgets.settings import proofread_tab
    from docwen_gui.widgets.settings.proofread_tab import _TyposDictionaryEditor
    from docwen_runtime.config.loader import ConfigLoader

    config_dir = tmp_path / "configs"
    config_dir.mkdir(parents=True, exist_ok=True)
    _write_minimal_base_config_tree(config_dir)
    loader = ConfigLoader(base_dir=config_dir, user_dir=config_dir)

    def _seed_typos(doc) -> None:
        table = toml_table()
        table["己"] = toml_value(["已"], "old comment")
        doc["entries"] = table

    assert loader.update_file_document(
        "proofread/typos.toml",
        _seed_typos,
    )

    monkeypatch.setattr(proofread_tab, "_choose_duplicate_strategy", lambda *_args, **_kwargs: "merge")

    dialog = _TyposDictionaryEditor(config_dir / "proofread" / "typos.toml")
    dialog._add_row()  # pyright: ignore[reportPrivateUsage]
    row = dialog._table.rowCount() - 1
    dialog._table.setItem(row, 0, QTableWidgetItem("己"))
    dialog._table.setItem(row, 1, QTableWidgetItem("己经|以"))
    dialog._table.setItem(row, 2, QTableWidgetItem("new comment"))

    dialog._on_accept()  # pyright: ignore[reportPrivateUsage]

    saved = loader.get_file_dict("proofread/typos.toml")
    assert saved.get("entries", {}).get("己") == ["已", "己经", "以"]
    saved_text = (config_dir / "proofread" / "typos.toml").read_text(encoding="utf-8")
    assert "new comment" in saved_text

    dialog.close()


def test_typos_editor_duplicate_key_overwrite_strategy_preserves_old_project_behavior(
    qapp,
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_gui.widgets.settings import proofread_tab
    from docwen_gui.widgets.settings.proofread_tab import _TyposDictionaryEditor
    from docwen_runtime.config.loader import ConfigLoader

    config_dir = tmp_path / "configs"
    config_dir.mkdir(parents=True, exist_ok=True)
    _write_minimal_base_config_tree(config_dir)
    loader = ConfigLoader(base_dir=config_dir, user_dir=config_dir)

    def _seed_typos(doc) -> None:
        table = toml_table()
        table["己"] = toml_value(["已"], "old comment")
        doc["entries"] = table

    assert loader.update_file_document(
        "proofread/typos.toml",
        _seed_typos,
    )

    monkeypatch.setattr(proofread_tab, "_choose_duplicate_strategy", lambda *_args, **_kwargs: "overwrite")

    dialog = _TyposDictionaryEditor(config_dir / "proofread" / "typos.toml")
    dialog._add_row()  # pyright: ignore[reportPrivateUsage]
    row = dialog._table.rowCount() - 1
    dialog._table.setItem(row, 0, QTableWidgetItem("己"))
    dialog._table.setItem(row, 1, QTableWidgetItem("己经"))
    dialog._table.setItem(row, 2, QTableWidgetItem("new comment"))

    dialog._on_accept()  # pyright: ignore[reportPrivateUsage]

    saved = loader.get_file_dict("proofread/typos.toml")
    assert saved.get("entries", {}).get("己") == ["己经"]
    saved_text = (config_dir / "proofread" / "typos.toml").read_text(encoding="utf-8")
    assert "new comment" in saved_text
    assert "old comment" not in saved_text

    dialog.close()


def test_typos_editor_duplicate_key_cancel_strategy_aborts_save(
    qapp,
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_gui.widgets.settings import proofread_tab
    from docwen_gui.widgets.settings.proofread_tab import _TyposDictionaryEditor
    from docwen_runtime.config.loader import ConfigLoader

    config_dir = tmp_path / "configs"
    config_dir.mkdir(parents=True, exist_ok=True)
    _write_minimal_base_config_tree(config_dir)
    loader = ConfigLoader(base_dir=config_dir, user_dir=config_dir)

    def _seed_typos(doc) -> None:
        table = toml_table()
        table["己"] = toml_value(["已"], "old comment")
        doc["entries"] = table

    assert loader.update_file_document(
        "proofread/typos.toml",
        _seed_typos,
    )

    monkeypatch.setattr(proofread_tab, "_choose_duplicate_strategy", lambda *_args, **_kwargs: "cancel")

    dialog = _TyposDictionaryEditor(config_dir / "proofread" / "typos.toml")
    dialog._add_row()  # pyright: ignore[reportPrivateUsage]
    row = dialog._table.rowCount() - 1
    dialog._table.setItem(row, 0, QTableWidgetItem("己"))
    dialog._table.setItem(row, 1, QTableWidgetItem("己经"))
    dialog._table.setItem(row, 2, QTableWidgetItem("new comment"))

    dialog._on_accept()  # pyright: ignore[reportPrivateUsage]

    saved = loader.get_file_dict("proofread/typos.toml")
    assert saved.get("entries", {}).get("己") == ["已"]
    saved_text = (config_dir / "proofread" / "typos.toml").read_text(encoding="utf-8")
    assert "old comment" in saved_text
    assert "new comment" not in saved_text

    dialog.close()


# ── Sensitive Word Editor ────────────────────────────────────────────────────


def test_sensitive_word_editor_preserves_exception_items_and_remark(
    qapp,
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_gui.widgets.settings.proofread_tab import _SensitiveWordEditor
    from docwen_runtime.config.loader import ConfigLoader

    config_dir = tmp_path / "configs"
    config_dir.mkdir(parents=True, exist_ok=True)
    _write_minimal_base_config_tree(config_dir)
    loader = ConfigLoader(base_dir=config_dir, user_dir=config_dir)

    def _seed_sensitive_words(doc) -> None:
        table = toml_table()
        table["机密"] = toml_value(["公开稿", "审批稿"], "允许上下文")
        doc["entries"] = table

    assert loader.update_file_document(
        "proofread/sensitive_words.toml",
        _seed_sensitive_words,
    )

    dialog = _SensitiveWordEditor(config_dir / "proofread" / "sensitive_words.toml")

    assert dialog._table.columnCount() == 3
    item00 = dialog._table.item(0, 0)
    assert item00 is not None
    assert item00.text() == "机密"
    item01 = dialog._table.item(0, 1)
    assert item01 is not None
    assert item01.text() == "公开稿|审批稿"
    item02 = dialog._table.item(0, 2)
    assert item02 is not None
    assert item02.text() == "允许上下文"

    dialog._on_accept()

    saved = loader.get_file_dict("proofread/sensitive_words.toml")
    assert saved.get("entries", {}).get("机密") == ["公开稿", "审批稿"]
    saved_text = (config_dir / "proofread" / "sensitive_words.toml").read_text(encoding="utf-8")
    assert "允许上下文" in saved_text

    dialog.close()


@pytest.mark.parametrize(
    "editor_name,relative,key,values,validator_argument",
    [
        (
            "_TyposDictionaryEditor",
            "proofread/typos.toml",
            "correct",
            ["foo|bar", "wrng", r"a\b", " spaced "],
            "typos_map",
        ),
        ("_SymbolErrorEditor", "proofread/symbol_map.toml", "｜", ["|", "丨", "\\", " spaced "], "symbol_map"),
        (
            "_SensitiveWordEditor",
            "proofread/sensitive_words.toml",
            "foo",
            ["foo|bar", r"foo\bar", " foo "],
            "sensitive_words",
        ),
    ],
)
@pytest.mark.parametrize("edit_unrelated", [False, True])
def test_array_editors_save_literal_values_through_config_port_without_semantic_drift(
    qapp, tmp_path, editor_name, relative, key, values, validator_argument, edit_unrelated
):
    import tomllib

    from docwen_bundle.config_port import ConfigPortAdapter
    from docwen_gui.widgets.settings import proofread_tab
    from docwen_plugin_proofread.text_validator import TextValidator
    from docwen_runtime.config.loader import ConfigLoader

    config_dir = tmp_path / "configs"
    _write_minimal_base_config_tree(config_dir)
    loader = ConfigLoader(base_dir=config_dir, user_dir=config_dir)
    expected = {key: values, "unrelated": ["other"]}
    assert loader.update_file_sections(relative, {"entries": expected})
    port = ConfigPortAdapter(loader)
    initial_text = port.get_file_text(relative)
    assert initial_text is not None
    dialog = getattr(proofread_tab, editor_name)(relative, source_text=initial_text, save_callback=port.save_file_text)
    try:
        if edit_unrelated:
            row = next(
                row for row in range(dialog._table.rowCount()) if dialog._table.item(row, 0).text() == "unrelated"
            )
            dialog._table.item(row, 2).setText("changed comment")
        dialog._on_accept()
        fresh = ConfigPortAdapter(base_dir=config_dir, user_dir=config_dir)
        saved = fresh.get_file_text(relative)
        assert saved is not None
        actual = tomllib.loads(saved)["entries"]
        assert actual == expected
        switches = {"typos_rule": False, "symbol_correction": False, "sensitive_word": False, "symbol_pairing": False}
        switches[
            {"typos_map": "typos_rule", "symbol_map": "symbol_correction", "sensitive_words": "sensitive_word"}[
                validator_argument
            ]
        ] = True
        original_options: dict[str, Any] = {validator_argument: expected}
        reloaded_options: dict[str, Any] = {validator_argument: actual}
        before = TextValidator(**original_options, enabled=switches)
        after = TextValidator(**reloaded_options, enabled=switches)

        def facts(validator, text):
            return [
                (e.start_pos, e.end_pos, e.error_text, e.replacement, e.source) for e in validator.validate_text(text)
            ]

        for text in ("foo and bar", "|", "foo", "foo|bar", "wrng", r"a\b", " spaced ", " foo "):
            assert facts(after, text) == facts(before, text)
        if validator_argument == "typos_map":
            assert not before.validate_text("foo and bar")
            assert [error.error_text for error in after.validate_text("foo|bar")] == ["foo|bar"]
        if validator_argument == "sensitive_words":
            assert [error.error_text for error in after.validate_text("foo")] == ["foo"]
            assert not after.validate_text("foo|bar")
        if validator_argument == "symbol_map":
            assert [(error.error_text, error.replacement) for error in after.validate_text("|")] == [("|", "｜")]
    finally:
        dialog.deleteLater()


@pytest.mark.parametrize(
    "editor_name,relative",
    [
        ("_TyposDictionaryEditor", "proofread/typos.toml"),
        ("_SymbolErrorEditor", "proofread/symbol_map.toml"),
        ("_SensitiveWordEditor", "proofread/sensitive_words.toml"),
    ],
)
def test_array_editor_does_not_silently_drop_invalid_empty_elements(qapp, tmp_path, monkeypatch, editor_name, relative):
    from docwen_bundle.config_port import ConfigPortAdapter
    from docwen_gui.widgets.settings import proofread_tab
    from docwen_runtime.config.loader import ConfigLoader

    config_dir = tmp_path / "configs"
    _write_minimal_base_config_tree(config_dir)
    loader = ConfigLoader(base_dir=config_dir, user_dir=config_dir)
    assert loader.update_file_sections(relative, {"entries": {"key": ["one"]}})
    port = ConfigPortAdapter(loader)
    original = port.get_file_text(relative)
    dialog = getattr(proofread_tab, editor_name)(relative, source_text=original, save_callback=port.save_file_text)
    notices = []
    monkeypatch.setattr(proofread_tab.QMessageBox, "warning", lambda *args: notices.append(args))
    try:
        dialog._table.item(0, 1).setText("one||two")
        dialog._on_accept()
        assert notices
        assert port.get_file_text(relative) == original
        assert dialog.result() != dialog.DialogCode.Accepted
    finally:
        dialog.deleteLater()

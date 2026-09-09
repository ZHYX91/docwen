"""Proofread settings tab — engine toggles, skip rules, dictionary editors.

Matches old ProofreadTab behavior:
- 6 QCheckBox: symbol_pairing, symbol_correction, typos_rule, sensitive_word,
  skip.code_blocks, skip.quote_blocks (all default True)
- 3 MappingEditor entry buttons (symbol mapping, typos, sensitive words)
"""

from __future__ import annotations

from collections.abc import Callable
from pathlib import Path
from typing import Any
from typing import cast as _cast

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QCheckBox,
    QDialog,
    QDialogButtonBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from docwen_core.toml_tools import toml_table, toml_value
from docwen_runtime.config import atomic_write_text

from ...i18n import t
from ...view_models.settings_vm import SECTION_PROOFREAD, SettingsViewModel
from ..panel_card import ChoiceGroup, WrappingLabel
from .base_tab import BaseSettingsTab
from .proofread_transfer import ProofreadRuleTransfer

# ── Path resolution ─────────────────────────────────────────────────────────


ConfigTextSaveCallback = Callable[[str, str], bool]


def _load_toml_document(path: Path, source_text: str | None = None):
    """Read a TOML file as a mutable tomlkit document via core primitives."""
    from docwen_core.toml_tools import read_toml_text

    if source_text is not None:
        return read_toml_text(source_text)
    if path.exists():
        return read_toml_text(path.read_text(encoding="utf-8"))
    return read_toml_text("")


def _save_toml_document(
    doc: Any,
    *,
    config_name: str,
    fallback_path: Path,
    save_callback: ConfigTextSaveCallback | None,
) -> bool:
    """Persist a mutable document through its injected production backend."""
    content = str(doc.as_string())
    if save_callback is not None:
        try:
            return bool(save_callback(config_name, content))
        except Exception:
            return False
    try:
        atomic_write_text(fallback_path, content)
    except Exception:
        return False
    return True


def _clean_inline_comment(comment: Any) -> str:
    if not isinstance(comment, str):
        return ""
    return comment.lstrip("#").strip()


def _split_multi_value(value: str) -> list[str]:
    return [part.strip() for part in str(value).split("|") if part.strip()]


def _choose_duplicate_strategy(
    parent: QWidget,
    *,
    entry_key: str,
    current_values: list[str],
    current_comment: str,
) -> str:
    """Ask how to handle a duplicate dictionary key."""
    current_values_text = ", ".join(current_values) if current_values else t("editors.mapping.duplicate_key_no_comment")
    current_comment_text = current_comment or t("editors.mapping.duplicate_key_no_comment")
    message = t(
        "editors.mapping.duplicate_key_message",
        entry_key=entry_key,
        current_values=current_values_text,
        current_comment=current_comment_text,
    )
    box = QMessageBox(parent)
    box.setWindowTitle(t("editors.mapping.duplicate_key_title"))
    box.setText(message)
    box.setIcon(QMessageBox.Icon.Warning)
    overwrite_btn = box.addButton(t("editors.mapping.overwrite"), QMessageBox.ButtonRole.AcceptRole)
    merge_btn = box.addButton(t("editors.mapping.merge"), QMessageBox.ButtonRole.ActionRole)
    cancel_btn = box.addButton(t("common.cancel"), QMessageBox.ButtonRole.RejectRole)
    box.setDefaultButton(cancel_btn)
    box.exec()
    clicked = box.clickedButton()
    if clicked is overwrite_btn:
        return "overwrite"
    if clicked is merge_btn:
        return "merge"
    return "cancel"


# ── Symbol Mapping Editor Dialog ────────────────────────────────────────────


class _BaseEditorDialog(QDialog):
    """Base class for mapping/typo/sensitive word editor dialogs.

    Provides a search/filter bar and a helper for multi-value (pipe-separated)
    values in the value column.
    """

    def _add_dialog_buttons(self, layout: QVBoxLayout, on_accept: Callable[[], None]) -> None:
        button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel,
            self,
        )
        button_box.button(QDialogButtonBox.StandardButton.Ok).setText(t("common.ok", "OK"))
        button_box.button(QDialogButtonBox.StandardButton.Cancel).setText(t("common.cancel", "Cancel"))
        button_box.accepted.connect(on_accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)

    def _configure_columns(self, table: QTableWidget) -> None:
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        table.horizontalHeader().setDefaultAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)

    def _add_search_bar(self, layout: QVBoxLayout) -> QLineEdit:
        """Add a search/filter QLineEdit above the table and return it."""
        self._current_page = 1  # type: ignore[attr-defined]
        self._page_size = 20  # type: ignore[attr-defined]
        search_box = QLineEdit(self)
        search_box.setPlaceholderText(t("editors.mapping.search_placeholder", "Filter..."))
        search_box.textChanged.connect(lambda _: self._on_search_changed())
        self._search_box = search_box  # type: ignore[attr-defined]
        layout.addWidget(search_box)
        return search_box

    def _add_multi_value_hint(self, layout: QVBoxLayout) -> None:
        """Add a small hint label about pipe-separated multi-values."""
        hint = WrappingLabel(
            t(
                "editors.mapping.multi_value_hint",
                "Tip: use | to separate multiple values (for example: value1|value2|value3)",
            ),
            self,
        )
        hint.setObjectName("proofreadHintLabel")
        layout.addWidget(hint)

    def _add_pagination_bar(self, layout: QVBoxLayout) -> None:
        """Add previous/next page controls below the table."""
        page_row = QWidget(self)
        page_layout = QHBoxLayout(page_row)
        page_layout.setContentsMargins(0, 0, 0, 0)
        page_layout.addStretch(1)

        self._prev_page_button = QPushButton("<", page_row)  # type: ignore[attr-defined]
        self._prev_page_button.clicked.connect(self._go_to_previous_page)  # type: ignore[attr-defined]
        page_layout.addWidget(self._prev_page_button)  # type: ignore[attr-defined]

        self._page_label = QLabel("1 / 1", page_row)  # type: ignore[attr-defined]
        page_layout.addWidget(self._page_label)  # type: ignore[attr-defined]

        self._next_page_button = QPushButton(">", page_row)  # type: ignore[attr-defined]
        self._next_page_button.clicked.connect(self._go_to_next_page)  # type: ignore[attr-defined]
        page_layout.addWidget(self._next_page_button)  # type: ignore[attr-defined]
        layout.addWidget(page_row)

    def _on_search_changed(self) -> None:
        self._current_page = 1  # type: ignore[attr-defined]
        self._refresh_table_view()

    def _go_to_previous_page(self) -> None:
        if self._current_page <= 1:  # type: ignore[attr-defined]
            return
        self._current_page -= 1  # type: ignore[attr-defined]
        self._refresh_table_view()

    def _go_to_next_page(self) -> None:
        if self._current_page >= self._max_page():  # type: ignore[attr-defined]
            return
        self._current_page += 1  # type: ignore[attr-defined]
        self._refresh_table_view()

    def _matching_rows(self) -> list[int]:
        search_text = self._search_box.text().strip().lower() if hasattr(self, "_search_box") else ""
        table = getattr(self, "_table", None)
        if table is None:
            return []
        matches: list[int] = []
        for row in range(table.rowCount()):
            if not search_text:
                matches.append(row)
                continue
            for col in range(table.columnCount()):
                item = table.item(row, col)
                if item and search_text in item.text().strip().lower():
                    matches.append(row)
                    break
        return matches

    def _max_page(self) -> int:
        count = len(self._matching_rows())
        if count <= 0:
            return 1
        return (count - 1) // self._page_size + 1  # type: ignore[attr-defined]

    def _refresh_table_view(self) -> None:
        """Filter and paginate visible table rows."""
        table = getattr(self, "_table", None)
        if table is None:
            return
        matches = self._matching_rows()
        max_page = 1 if not matches else (len(matches) - 1) // self._page_size + 1  # type: ignore[attr-defined]
        self._current_page = min(max(self._current_page, 1), max_page)  # type: ignore[attr-defined]
        start = (self._current_page - 1) * self._page_size  # type: ignore[attr-defined]
        visible_page_rows = set(matches[start : start + self._page_size])  # type: ignore[attr-defined]
        for row in range(table.rowCount()):
            table.setRowHidden(row, row not in visible_page_rows)
        if hasattr(self, "_page_label"):
            self._page_label.setText(f"{self._current_page} / {max_page}")  # type: ignore[attr-defined]
        if hasattr(self, "_prev_page_button"):
            self._prev_page_button.setEnabled(self._current_page > 1)  # type: ignore[attr-defined]
        if hasattr(self, "_next_page_button"):
            self._next_page_button.setEnabled(self._current_page < max_page)  # type: ignore[attr-defined]

    def _refresh_after_row_change(self, select_row: int | None = None) -> None:
        self._current_page = self._max_page()  # type: ignore[attr-defined]
        self._refresh_table_view()
        table = getattr(self, "_table", None)
        if table is not None and select_row is not None and not table.isRowHidden(select_row):
            table.setCurrentCell(select_row, 0)

    def _coalesce_duplicate_entries(
        self,
        entries: list[dict[str, str]],
        *,
        key_field: str,
        values_field: str,
        comment_field: str,
    ) -> list[dict[str, str]] | None:
        """Apply old-project duplicate-key overwrite/merge/cancel semantics."""
        by_key: dict[str, dict[str, str]] = {}
        ordered_keys: list[str] = []
        for entry in entries:
            key = entry.get(key_field, "").strip()
            if not key:
                continue
            if key not in by_key:
                by_key[key] = dict(entry)
                ordered_keys.append(key)
                continue

            current = by_key[key]
            current_values = _split_multi_value(current.get(values_field, ""))
            new_values = _split_multi_value(entry.get(values_field, ""))
            strategy = _choose_duplicate_strategy(
                self,
                entry_key=key,
                current_values=current_values,
                current_comment=current.get(comment_field, ""),
            )
            if strategy == "cancel":
                return None
            if strategy == "overwrite":
                by_key[key] = dict(entry)
            else:
                merged_values = list(dict.fromkeys([*current_values, *new_values]))
                current[values_field] = "|".join(merged_values)
                if entry.get(comment_field, "").strip():
                    current[comment_field] = entry[comment_field].strip()
        return [by_key[key] for key in ordered_keys if key in by_key]


class _SymbolPairingEditor(_BaseEditorDialog):
    """Dialog for editing opening and closing punctuation pairs.

    Uses a supplied effective TOML source and production save callback for
    format-preserving I/O.
    """

    def __init__(
        self,
        toml_path: str | Path,
        parent: QWidget | None = None,
        *,
        source_text: str | None = None,
        save_callback: ConfigTextSaveCallback | None = None,
    ) -> None:
        super().__init__(parent)
        self.setObjectName("symbolPairingEditor")
        self._toml_path = Path(toml_path)
        self._config_name = self._toml_path.as_posix()
        self._source_text = source_text
        self._save_callback = save_callback
        self._doc = self._load_toml()
        self._entries: list[dict[str, str]] = []
        self._extract_entries()
        self._setup_ui()
        self._populate_table()

    # ── TOML I/O ─────────────────────────────────────────────────────────

    def _load_toml(self):
        return _load_toml_document(self._toml_path, self._source_text)

    def _extract_entries(self) -> None:
        pairs = self._doc.get("items", [])
        if isinstance(pairs, list):
            for item in pairs:
                if isinstance(item, list) and len(item) >= 2:
                    self._entries.append(
                        {
                            "source": str(item[0]),
                            "target": str(item[1]),
                        }
                    )
                elif isinstance(item, dict):
                    src = item.get("source") or item.get("open", "")
                    tgt = item.get("target") or item.get("close", "")
                    self._entries.append(
                        {
                            "source": str(src),
                            "target": str(tgt),
                        }
                    )

    def _write_toml(self) -> bool:
        new_pairs: list[list[str]] = []
        for entry in self._entries:
            src = entry.get("source", "").strip()
            tgt = entry.get("target", "").strip()
            if src:
                new_pairs.append([src, tgt])
        self._doc["items"] = new_pairs
        ok = _save_toml_document(
            self._doc,
            config_name=self._config_name,
            fallback_path=self._toml_path,
            save_callback=self._save_callback,
        )
        if not ok:
            QMessageBox.warning(
                self,
                t("common.save_failed", "Save Failed"),
                t("editors.mapping.save_symbol_mapping_failed", "Could not save symbol pairing entries."),
            )
        return ok

    # ── UI ───────────────────────────────────────────────────────────────

    def _setup_ui(self) -> None:
        self.setWindowTitle(t("editors.mapping.symbol_editor_title", "Punctuation Pairing Rules"))
        self.setMinimumSize(550, 400)

        layout = QVBoxLayout(self)

        self._add_search_bar(layout)

        self._table = QTableWidget(0, 2, self)
        self._table.setHorizontalHeaderLabels(
            [
                t("editors.mapping.opening_symbol", "Opening Symbol"),
                t("editors.mapping.closing_symbol", "Closing Symbol"),
            ]
        )
        self._configure_columns(self._table)
        layout.addWidget(self._table)
        self._add_pagination_bar(layout)

        # Add / Delete row buttons
        btn_row = QWidget(self)
        btn_layout = QHBoxLayout(btn_row)
        btn_layout.setContentsMargins(0, 0, 0, 0)
        btn_layout.setSpacing(8)

        add_btn = QPushButton(t("editors.mapping.add_entry", "Add Row"), btn_row)
        add_btn.clicked.connect(self._add_row)
        btn_layout.addWidget(add_btn)

        del_btn = QPushButton(t("editors.mapping.delete_entry", "Delete Row"), btn_row)
        del_btn.clicked.connect(self._delete_row)
        btn_layout.addWidget(del_btn)

        btn_layout.addStretch(1)
        layout.addWidget(btn_row)

        # Dialog buttons
        self._add_dialog_buttons(layout, self._on_accept)

    def _populate_table(self) -> None:
        self._table.setRowCount(len(self._entries))
        for row, entry in enumerate(self._entries):
            self._table.setItem(row, 0, QTableWidgetItem(entry.get("source", "")))
            self._table.setItem(row, 1, QTableWidgetItem(entry.get("target", "")))
        self._refresh_table_view()

    def _add_row(self) -> None:
        row = self._table.rowCount()
        self._table.insertRow(row)
        for c in range(2):
            self._table.setItem(row, c, QTableWidgetItem(""))
        self._refresh_after_row_change(row)

    def _delete_row(self) -> None:
        row = self._table.currentRow()
        if row >= 0:
            self._table.removeRow(row)
            self._refresh_after_row_change()

    def _on_accept(self) -> None:
        self._entries.clear()
        for row in range(self._table.rowCount()):
            src = self._table.item(row, 0)
            tgt = self._table.item(row, 1)
            source = src.text().strip() if src else ""
            target = tgt.text().strip() if tgt else ""
            if source:
                self._entries.append({"source": source, "target": target})
        if self._write_toml():
            self.accept()


# ── Typos Dictionary Editor Dialog ──────────────────────────────────────────


class _TyposDictionaryEditor(_BaseEditorDialog):
    """Dialog for editing correct words and their misspellings.

    Uses a supplied effective TOML source and production save callback for
    format-preserving I/O. Inline comments store the Remark field.
    """

    def __init__(
        self,
        toml_path: str | Path,
        parent: QWidget | None = None,
        *,
        source_text: str | None = None,
        save_callback: ConfigTextSaveCallback | None = None,
    ) -> None:
        super().__init__(parent)
        self.setObjectName("typosDictionaryEditor")
        self._toml_path = Path(toml_path)
        self._config_name = self._toml_path.as_posix()
        self._source_text = source_text
        self._save_callback = save_callback
        self._doc = self._load_toml()
        self._entries: list[dict[str, str]] = []
        self._extract_entries()
        self._setup_ui()
        self._populate_table()

    # ── TOML I/O ─────────────────────────────────────────────────────────

    def _load_toml(self):
        return _load_toml_document(self._toml_path, self._source_text)

    def _extract_entries(self) -> None:
        raw_typos: Any = self._doc.get("entries", {})
        if not isinstance(raw_typos, dict):
            return

        # Determine if this is a tomlkit Table (has .value with items) or plain dict
        try:
            inner_items = raw_typos.value.items()  # type: ignore[union-attr]
        except (AttributeError, TypeError):
            inner_items = None

        if inner_items is not None:
            for key_item, val_item in inner_items:
                key = str(key_item)
                raw_val: Any = val_item.value if hasattr(val_item, "value") else val_item
                if isinstance(raw_val, list):
                    misspellings = "|".join(str(v) for v in raw_val)
                else:
                    misspellings = str(raw_val) if raw_val else ""
                remark = ""
                trivia = getattr(val_item, "trivia", None)
                if trivia is not None and hasattr(trivia, "comment"):
                    remark = _clean_inline_comment(getattr(trivia, "comment", ""))
                self._entries.append(
                    {
                        "correct_word": key,
                        "misspellings": misspellings,
                        "remark": remark,
                    }
                )
        else:
            # Plain dict fallback
            for key, val in raw_typos.items():
                misspellings = "|".join(str(v) for v in val) if isinstance(val, list) else (str(val) if val else "")
                self._entries.append(
                    {
                        "correct_word": str(key),
                        "misspellings": misspellings,
                        "remark": "",
                    }
                )

    def _write_toml(self) -> bool:
        typos_tbl = toml_table()
        for entry in self._entries:
            correct_word = entry.get("correct_word", "").strip()
            misspellings_text = entry.get("misspellings", "").strip()
            remark = entry.get("remark", "").strip()
            if not correct_word:
                continue
            misspellings = [c.strip() for c in misspellings_text.split("|") if c.strip()]
            typos_tbl[correct_word] = toml_value(misspellings, remark if remark else "")
        self._doc["entries"] = typos_tbl
        ok = _save_toml_document(
            self._doc,
            config_name=self._config_name,
            fallback_path=self._toml_path,
            save_callback=self._save_callback,
        )
        if not ok:
            QMessageBox.warning(
                self,
                t("common.save_failed", "Save Failed"),
                t("editors.mapping.save_typos_failed", "Could not save typos dictionary."),
            )
        return ok

    # ── UI ───────────────────────────────────────────────────────────────

    def _setup_ui(self) -> None:
        self.setWindowTitle(t("editors.mapping.typo_editor_title", "Edit Typos Dictionary"))
        self.setMinimumSize(550, 400)

        layout = QVBoxLayout(self)

        self._add_search_bar(layout)
        self._add_multi_value_hint(layout)

        self._table = QTableWidget(0, 3, self)
        self._table.setHorizontalHeaderLabels(
            [
                t("editors.mapping.correct_word", "Correct Word"),
                t("editors.mapping.misspellings", "Misspellings"),
                t("editors.mapping.comment", "Remark"),
            ]
        )
        self._configure_columns(self._table)
        layout.addWidget(self._table)
        self._add_pagination_bar(layout)

        # Add / Delete row buttons
        btn_row = QWidget(self)
        btn_layout = QHBoxLayout(btn_row)
        btn_layout.setContentsMargins(0, 0, 0, 0)
        btn_layout.setSpacing(8)

        add_btn = QPushButton(t("editors.mapping.add_entry", "Add Row"), btn_row)
        add_btn.clicked.connect(self._add_row)
        btn_layout.addWidget(add_btn)

        del_btn = QPushButton(t("editors.mapping.delete_entry", "Delete Row"), btn_row)
        del_btn.clicked.connect(self._delete_row)
        btn_layout.addWidget(del_btn)

        btn_layout.addStretch(1)
        layout.addWidget(btn_row)

        # Dialog buttons
        self._add_dialog_buttons(layout, self._on_accept)

    def _populate_table(self) -> None:
        self._table.setRowCount(len(self._entries))
        for row, entry in enumerate(self._entries):
            self._table.setItem(row, 0, QTableWidgetItem(entry.get("correct_word", "")))
            self._table.setItem(row, 1, QTableWidgetItem(entry.get("misspellings", "")))
            self._table.setItem(row, 2, QTableWidgetItem(entry.get("remark", "")))
        self._refresh_table_view()

    def _add_row(self) -> None:
        row = self._table.rowCount()
        self._table.insertRow(row)
        for c in range(3):
            self._table.setItem(row, c, QTableWidgetItem(""))
        self._refresh_after_row_change(row)

    def _delete_row(self) -> None:
        row = self._table.currentRow()
        if row >= 0:
            self._table.removeRow(row)
            self._refresh_after_row_change()

    def _on_accept(self) -> None:
        self._entries.clear()
        for row in range(self._table.rowCount()):
            correct_item = self._table.item(row, 0)
            misspellings_item = self._table.item(row, 1)
            rmk_item = self._table.item(row, 2)
            correct_word = correct_item.text().strip() if correct_item else ""
            misspellings_text = misspellings_item.text().strip() if misspellings_item else ""
            remark = rmk_item.text().strip() if rmk_item else ""
            if correct_word:
                self._entries.append(
                    {"correct_word": correct_word, "misspellings": misspellings_text, "remark": remark}
                )
        coalesced = self._coalesce_duplicate_entries(
            self._entries,
            key_field="correct_word",
            values_field="misspellings",
            comment_field="remark",
        )
        if coalesced is None:
            return
        self._entries = coalesced
        if self._write_toml():
            self.accept()


# ── Symbol Error Editor Dialog ──────────────────────────────────────────────


class _SymbolErrorEditor(_BaseEditorDialog):
    """Dialog for editing symbol correction entries (全半角映射).

    Uses a supplied effective TOML source and production save callback.
    Inline comments store the Remark field.
    """

    def __init__(
        self,
        toml_path: str | Path,
        parent: QWidget | None = None,
        *,
        source_text: str | None = None,
        save_callback: ConfigTextSaveCallback | None = None,
    ) -> None:
        super().__init__(parent)
        self.setObjectName("symbolErrorEditor")
        self._toml_path = Path(toml_path)
        self._config_name = self._toml_path.as_posix()
        self._source_text = source_text
        self._save_callback = save_callback
        self._doc = self._load_toml()
        self._entries: list[dict[str, str]] = []
        self._extract_entries()
        self._setup_ui()
        self._populate_table()

    # ── TOML I/O ─────────────────────────────────────────────────────────

    def _load_toml(self):
        return _load_toml_document(self._toml_path, self._source_text)

    def _extract_entries(self) -> None:
        raw_symbols: Any = self._doc.get("entries", {})
        if not isinstance(raw_symbols, dict):
            return

        # Determine if this is a tomlkit Table (has .value with items) or plain dict
        try:
            inner_items = raw_symbols.value.items()  # type: ignore[union-attr]
        except (AttributeError, TypeError):
            inner_items = None

        if inner_items is not None:
            for key_item, val_item in inner_items:
                correct = str(key_item)
                raw_val: Any = val_item.value if hasattr(val_item, "value") else val_item
                if isinstance(raw_val, list):
                    errors = "|".join(str(v) for v in raw_val)
                else:
                    errors = str(raw_val) if raw_val else ""
                remark = ""
                trivia = getattr(val_item, "trivia", None)
                if trivia is not None and hasattr(trivia, "comment"):
                    remark = _clean_inline_comment(getattr(trivia, "comment", ""))
                self._entries.append(
                    {
                        "correct": correct,
                        "errors": errors,
                        "remark": remark,
                    }
                )
        else:
            for correct, raw_val in raw_symbols.items():
                errors = ""
                if isinstance(raw_val, list):
                    errors = "|".join(str(v) for v in raw_val)
                elif raw_val:
                    errors = str(raw_val)
                self._entries.append(
                    {
                        "correct": str(correct),
                        "errors": errors,
                        "remark": "",
                    }
                )

    def _write_toml(self) -> bool:
        tbl = toml_table()
        for entry in self._entries:
            correct = entry.get("correct", "").strip()
            error_str = entry.get("errors", "").strip()
            remark = entry.get("remark", "").strip()
            if not correct:
                continue
            error_list = [item.strip() for item in error_str.split("|") if item.strip()]
            tbl[correct] = toml_value(error_list, remark if remark else "")
        self._doc["entries"] = tbl
        ok = _save_toml_document(
            self._doc,
            config_name=self._config_name,
            fallback_path=self._toml_path,
            save_callback=self._save_callback,
        )
        if not ok:
            QMessageBox.warning(
                self,
                t("common.save_failed", "Save Failed"),
                t("editors.mapping.save_symbol_correction_failed", "Could not save symbol correction entries."),
            )
        return ok

    # ── UI ───────────────────────────────────────────────────────────────

    def _setup_ui(self) -> None:
        self.setWindowTitle(t("settings.proofread.symbol_correction_section", "Edit Symbol Correction"))
        self.setMinimumSize(550, 400)

        layout = QVBoxLayout(self)

        self._add_search_bar(layout)
        self._add_multi_value_hint(layout)

        self._table = QTableWidget(0, 3, self)
        self._table.setHorizontalHeaderLabels(
            [
                t("editors.mapping.symbol_correct_symbol", "Correct Symbol"),
                t("editors.mapping.symbol_error_symbols", "Error Variant(s)"),
                t("editors.mapping.comment", "Remark"),
            ]
        )
        self._configure_columns(self._table)
        layout.addWidget(self._table)
        self._add_pagination_bar(layout)

        # Add / Delete row buttons
        btn_row = QWidget(self)
        btn_layout = QHBoxLayout(btn_row)
        btn_layout.setContentsMargins(0, 0, 0, 0)
        btn_layout.setSpacing(8)

        add_btn = QPushButton(t("editors.mapping.add_entry", "Add Row"), btn_row)
        add_btn.clicked.connect(self._add_row)
        btn_layout.addWidget(add_btn)

        del_btn = QPushButton(t("editors.mapping.delete_entry", "Delete Row"), btn_row)
        del_btn.clicked.connect(self._delete_row)
        btn_layout.addWidget(del_btn)

        btn_layout.addStretch(1)
        layout.addWidget(btn_row)

        # Dialog buttons
        self._add_dialog_buttons(layout, self._on_accept)

    def _populate_table(self) -> None:
        self._table.setRowCount(len(self._entries))
        for row, entry in enumerate(self._entries):
            self._table.setItem(row, 0, QTableWidgetItem(entry.get("correct", "")))
            self._table.setItem(row, 1, QTableWidgetItem(entry.get("errors", "")))
            self._table.setItem(row, 2, QTableWidgetItem(entry.get("remark", "")))
        self._refresh_table_view()

    def _add_row(self) -> None:
        row = self._table.rowCount()
        self._table.insertRow(row)
        for col in range(3):
            self._table.setItem(row, col, QTableWidgetItem(""))
        self._refresh_after_row_change(row)

    def _delete_row(self) -> None:
        row = self._table.currentRow()
        if row >= 0:
            self._table.removeRow(row)
            self._refresh_after_row_change()

    def _on_accept(self) -> None:
        self._entries.clear()
        for row in range(self._table.rowCount()):
            corr_item = self._table.item(row, 0)
            err_item = self._table.item(row, 1)
            rmk_item = self._table.item(row, 2)
            correct = corr_item.text().strip() if corr_item else ""
            errors = err_item.text().strip() if err_item else ""
            remark = rmk_item.text().strip() if rmk_item else ""
            if correct:
                self._entries.append(
                    {
                        "correct": correct,
                        "errors": errors,
                        "remark": remark,
                    }
                )
        coalesced = self._coalesce_duplicate_entries(
            self._entries,
            key_field="correct",
            values_field="errors",
            comment_field="remark",
        )
        if coalesced is None:
            return
        self._entries = coalesced
        if self._write_toml():
            self.accept()


# ── Sensitive Word Editor Dialog ────────────────────────────────────────────


class _SensitiveWordEditor(_BaseEditorDialog):
    """Dialog for editing sensitive word entries.

    Uses a supplied effective TOML source and production save callback for
    format-preserving I/O. Inline comments store the Remark field.
    """

    def __init__(
        self,
        toml_path: str | Path,
        parent: QWidget | None = None,
        *,
        source_text: str | None = None,
        save_callback: ConfigTextSaveCallback | None = None,
    ) -> None:
        super().__init__(parent)
        self.setObjectName("sensitiveWordEditor")
        self._toml_path = Path(toml_path)
        self._config_name = self._toml_path.as_posix()
        self._source_text = source_text
        self._save_callback = save_callback
        self._doc = self._load_toml()
        self._entries: list[dict[str, str]] = []
        self._extract_entries()
        self._setup_ui()
        self._populate_table()

    # ── TOML I/O ─────────────────────────────────────────────────────────

    def _load_toml(self):
        return _load_toml_document(self._toml_path, self._source_text)

    def _extract_entries(self) -> None:
        raw_words: Any = self._doc.get("entries", {})
        if not isinstance(raw_words, dict):
            return

        # Determine if this is a tomlkit Table (has .value with items) or plain dict
        try:
            inner_items = raw_words.value.items()  # type: ignore[union-attr]
        except (AttributeError, TypeError):
            inner_items = None

        if inner_items is not None:
            for key_item, val_item in inner_items:
                word = str(key_item)
                raw_val: Any = val_item.value if hasattr(val_item, "value") else val_item
                if isinstance(raw_val, list):
                    exception_items = "|".join(str(v) for v in raw_val)
                else:
                    exception_items = str(raw_val) if raw_val else ""
                remark = ""
                trivia = getattr(val_item, "trivia", None)
                if trivia is not None and hasattr(trivia, "comment"):
                    remark = _clean_inline_comment(getattr(trivia, "comment", ""))
                self._entries.append(
                    {
                        "word": word,
                        "exception_items": exception_items,
                        "remark": remark,
                    }
                )
        else:
            for word, raw_val in raw_words.items():
                exception_items = ""
                if isinstance(raw_val, list):
                    exception_items = "|".join(str(v) for v in raw_val)
                elif raw_val:
                    exception_items = str(raw_val)
                self._entries.append(
                    {
                        "word": str(word),
                        "exception_items": exception_items,
                        "remark": "",
                    }
                )

    def _write_toml(self) -> bool:
        tbl = toml_table()
        for entry in self._entries:
            word = entry.get("word", "").strip()
            exception_items = entry.get("exception_items", "").strip()
            remark = entry.get("remark", "").strip()
            if not word:
                continue
            exceptions = [item.strip() for item in exception_items.split("|") if item.strip()]
            tbl[word] = toml_value(exceptions, remark if remark else "")
        self._doc["entries"] = tbl
        ok = _save_toml_document(
            self._doc,
            config_name=self._config_name,
            fallback_path=self._toml_path,
            save_callback=self._save_callback,
        )
        if not ok:
            QMessageBox.warning(
                self,
                t("common.save_failed", "Save Failed"),
                t("editors.mapping.save_sensitive_words_failed", "Could not save sensitive words dictionary."),
            )
        return ok

    # ── UI ───────────────────────────────────────────────────────────────

    def _setup_ui(self) -> None:
        self.setWindowTitle(t("editors.mapping.sensitive_editor_title", "Edit Sensitive Words"))
        self.setMinimumSize(550, 400)

        layout = QVBoxLayout(self)

        self._add_search_bar(layout)
        self._add_multi_value_hint(layout)

        self._table = QTableWidget(0, 3, self)
        self._table.setHorizontalHeaderLabels(
            [
                t("editors.mapping.sensitive_sensitive_word", "Word"),
                t("editors.mapping.sensitive_exceptions", "Exception Items"),
                t("editors.mapping.comment", "Remark"),
            ]
        )
        self._configure_columns(self._table)
        layout.addWidget(self._table)
        self._add_pagination_bar(layout)

        # Add / Delete row buttons
        btn_row = QWidget(self)
        btn_layout = QHBoxLayout(btn_row)
        btn_layout.setContentsMargins(0, 0, 0, 0)
        btn_layout.setSpacing(8)

        add_btn = QPushButton(t("editors.mapping.add_entry", "Add Row"), btn_row)
        add_btn.clicked.connect(self._add_row)
        btn_layout.addWidget(add_btn)

        del_btn = QPushButton(t("editors.mapping.delete_entry", "Delete Row"), btn_row)
        del_btn.clicked.connect(self._delete_row)
        btn_layout.addWidget(del_btn)

        btn_layout.addStretch(1)
        layout.addWidget(btn_row)

        # Dialog buttons
        self._add_dialog_buttons(layout, self._on_accept)

    def _populate_table(self) -> None:
        self._table.setRowCount(len(self._entries))
        for row, entry in enumerate(self._entries):
            self._table.setItem(row, 0, QTableWidgetItem(entry.get("word", "")))
            self._table.setItem(row, 1, QTableWidgetItem(entry.get("exception_items", "")))
            self._table.setItem(row, 2, QTableWidgetItem(entry.get("remark", "")))
        self._refresh_table_view()

    def _add_row(self) -> None:
        row = self._table.rowCount()
        self._table.insertRow(row)
        for col in range(3):
            self._table.setItem(row, col, QTableWidgetItem(""))
        self._refresh_after_row_change(row)

    def _delete_row(self) -> None:
        row = self._table.currentRow()
        if row >= 0:
            self._table.removeRow(row)
            self._refresh_after_row_change()

    def _on_accept(self) -> None:
        self._entries.clear()
        for row in range(self._table.rowCount()):
            word_item = self._table.item(row, 0)
            exception_item = self._table.item(row, 1)
            rmk_item = self._table.item(row, 2)
            word = word_item.text().strip() if word_item else ""
            exception_items = exception_item.text().strip() if exception_item else ""
            remark = rmk_item.text().strip() if rmk_item else ""
            if word:
                self._entries.append(
                    {
                        "word": word,
                        "exception_items": exception_items,
                        "remark": remark,
                    }
                )
        coalesced = self._coalesce_duplicate_entries(
            self._entries,
            key_field="word",
            values_field="exception_items",
            comment_field="remark",
        )
        if coalesced is None:
            return
        self._entries = coalesced
        if self._write_toml():
            self.accept()


# ── ProofreadTab ────────────────────────────────────────────────────────────


class ProofreadTab(BaseSettingsTab):
    """Proofread engine settings tab."""

    def __init__(self, view_model: SettingsViewModel) -> None:
        self._vm = view_model
        self._rule_transfer = ProofreadRuleTransfer(self, view_model)
        self._symbol_pairing: QCheckBox = _cast(QCheckBox, None)
        self._symbol_correction: QCheckBox = _cast(QCheckBox, None)
        self._typos_rule: QCheckBox = _cast(QCheckBox, None)
        self._sensitive_word: QCheckBox = _cast(QCheckBox, None)
        self._skip_code_blocks: QCheckBox = _cast(QCheckBox, None)
        self._skip_quote_blocks: QCheckBox = _cast(QCheckBox, None)
        super().__init__()
        self._load_values()

    def _create_interface(self) -> None:
        # ── Rules card ──────────────────────────────────────────────────
        _rules_card, rules_form = self.add_settings_card(t("settings.proofread.validation_section", "Validation Rules"))
        self._symbol_pairing = self.create_settings_toggle(
            t("settings.proofread.enable_symbol_pairing", "Enable Symbol Pairing Check")
        )
        rules_form.addRow(self._symbol_pairing)
        self._symbol_correction = self.create_settings_toggle(
            t("settings.proofread.enable_symbol_correction", "Enable Symbol Correction")
        )
        rules_form.addRow(self._symbol_correction)
        self._typos_rule = self.create_settings_toggle(
            t("settings.proofread.enable_typos_rule", "Enable Typos Detection")
        )
        rules_form.addRow(self._typos_rule)
        self._sensitive_word = self.create_settings_toggle(
            t("settings.proofread.enable_sensitive_word", "Enable Sensitive Word Check")
        )
        rules_form.addRow(self._sensitive_word)
        self._skip_code_blocks = self.create_settings_toggle(
            t("settings.proofread.skip_code_blocks", "Skip Code Blocks")
        )
        rules_form.addRow(self._skip_code_blocks)
        self._skip_quote_blocks = self.create_settings_toggle(
            t("settings.proofread.skip_quote_blocks", "Skip Quote Blocks")
        )
        rules_form.addRow(self._skip_quote_blocks)

        # ── Dictionary editor cards ─────────────────────────────────────
        # Symbol Mapping Editor
        _card_sym, form_sym = self.add_settings_card(
            t("settings.proofread.symbol_mapping_section", "Punctuation Pairing Rules"),
            t("settings.proofread.symbol_mapping_desc", "Check that opening and closing punctuation is paired."),
        )
        self._add_editor_buttons(form_sym, self._open_symbol_pairing_editor, "proofread/pairs.toml")

        # Symbol Error Editor
        _card_err, form_err = self.add_settings_card(
            t("settings.proofread.symbol_correction_section", "Symbol Correction Editor"),
            t("settings.proofread.symbol_correction_desc", "Manage fullwidth/halfwidth symbol correction entries."),
        )
        self._add_editor_buttons(form_err, self._open_symbol_error_editor, "proofread/symbol_map.toml")

        # Typos Dictionary Editor
        _card_typo, form_typo = self.add_settings_card(
            t("settings.proofread.typos_section", "Typos Dictionary Editor"),
            t("settings.proofread.typos_desc", "Manage common typo correction entries."),
        )
        self._add_editor_buttons(form_typo, self._open_typos_editor, "proofread/typos.toml")

        # Sensitive Words Editor
        _card_sw, form_sw = self.add_settings_card(
            t("settings.proofread.sensitive_words_section", "Sensitive Words Editor"),
            t("settings.proofread.sensitive_words_desc", "Manage sensitive word detection entries."),
        )
        self._add_editor_buttons(form_sw, self._open_sensitive_words_editor, "proofread/sensitive_words.toml")

        # Wire toggles
        self._symbol_pairing.toggled.connect(lambda v: self._vm.set_field(SECTION_PROOFREAD, "symbol_pairing", v))
        self._symbol_correction.toggled.connect(lambda v: self._vm.set_field(SECTION_PROOFREAD, "symbol_correction", v))
        self._typos_rule.toggled.connect(lambda v: self._vm.set_field(SECTION_PROOFREAD, "typos_rule", v))
        self._sensitive_word.toggled.connect(lambda v: self._vm.set_field(SECTION_PROOFREAD, "sensitive_word", v))
        self._skip_code_blocks.toggled.connect(lambda v: self._vm.set_field(SECTION_PROOFREAD, "skip_code_blocks", v))
        self._skip_quote_blocks.toggled.connect(lambda v: self._vm.set_field(SECTION_PROOFREAD, "skip_quote_blocks", v))

    def _add_editor_buttons(self, form, slot, config_name: str) -> None:
        """All three actions address the same complete editable source."""
        button_row = ChoiceGroup(self, responsive=True)
        button_layout = button_row.content_layout
        btn = QPushButton(t("settings.proofread.edit", "Edit"), button_row)
        btn.clicked.connect(slot)
        button_layout.addWidget(btn, 1)
        import_button = QPushButton(t("settings.rule_transfer.import"), button_row)
        import_button.clicked.connect(lambda: self._rule_transfer.import_rules(config_name))
        button_layout.addWidget(import_button, 1)
        export_button = QPushButton(t("settings.rule_transfer.export"), button_row)
        export_button.clicked.connect(lambda: self._rule_transfer.export_rules(config_name))
        button_layout.addWidget(export_button, 1)
        form.addRow(button_row)

    def _load_values(self) -> None:
        proof = self._vm.config.proofread
        self._symbol_pairing.setChecked(proof.symbol_pairing)
        self._symbol_correction.setChecked(proof.symbol_correction)
        self._typos_rule.setChecked(proof.typos_rule)
        self._sensitive_word.setChecked(proof.sensitive_word)
        self._skip_code_blocks.setChecked(proof.skip_code_blocks)
        self._skip_quote_blocks.setChecked(proof.skip_quote_blocks)

    def reload_from_config(self) -> None:
        self._load_values()

    # ── Editor dialogs ──────────────────────────────────────────────────────

    def _open_config_editor(self, editor_type: Any, config_name: str) -> None:
        source_text = self._vm.read_config_file_text(config_name)
        if source_text is None:
            QMessageBox.warning(
                self,
                t("common.error", "Error"),
                t("editors.mapping.load_failed", "Could not load the editable configuration source."),
            )
            return
        dlg = editor_type(
            config_name,
            self,
            source_text=source_text,
            save_callback=self._vm.save_config_file_text,
        )
        dlg.exec()

    def _open_symbol_pairing_editor(self) -> None:
        """Open the pairing editor against the injected editable source."""
        self._open_config_editor(_SymbolPairingEditor, "proofread/pairs.toml")

    def _open_symbol_error_editor(self) -> None:
        """Open the symbol-correction editor against the injected source."""
        self._open_config_editor(_SymbolErrorEditor, "proofread/symbol_map.toml")

    def _open_typos_editor(self) -> None:
        """Open the typos editor against the injected editable source."""
        self._open_config_editor(_TyposDictionaryEditor, "proofread/typos.toml")

    def _open_sensitive_words_editor(self) -> None:
        """Open the sensitive-words editor against the injected source."""
        self._open_config_editor(_SensitiveWordEditor, "proofread/sensitive_words.toml")

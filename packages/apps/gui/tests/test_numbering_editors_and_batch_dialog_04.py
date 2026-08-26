"""Focused tests split from test_numbering_editors_and_batch_dialog.py."""

from __future__ import annotations

from ._numbering_editors_and_batch_dialog_support import (
    Path,
    QApplication,
    QMessageBox,
    _patch_all_modals,
    pytest,
)

pytestmark = pytest.mark.gui


class TestTomlEditorWidget:
    """Test the TOML editor widget."""

    def test_create_editor(self, qapp: QApplication, tmp_path: Path) -> None:
        from docwen_gui.widgets.settings.toml_editor import TomlEditorWidget

        toml_file = tmp_path / "test.toml"
        toml_file.write_text('[section]\nkey = "value"\n', encoding="utf-8")

        def path_resolver(name: str) -> Path:
            return toml_file

        widget = TomlEditorWidget(
            None,
            config_name="test",
            path_resolver=path_resolver,
        )
        assert widget._editor is not None
        assert "key" in widget._editor.toPlainText()
        widget.close()

    def test_save_valid_toml(self, qapp: QApplication, tmp_path: Path) -> None:
        from docwen_gui.widgets.settings.toml_editor import TomlEditorWidget

        toml_file = tmp_path / "save_test.toml"
        toml_file.write_text("", encoding="utf-8")

        reloaded: list[bool] = []

        def path_resolver(name: str) -> Path:
            return toml_file

        def reload_callback() -> None:
            reloaded.append(True)

        widget = TomlEditorWidget(
            None,
            config_name="test",
            path_resolver=path_resolver,
            reload_callback=reload_callback,
        )
        assert widget._editor is not None
        widget._editor.setPlainText('[new_section]\nkey = "saved"\n')
        ok = widget.save_to_disk(show_success=False)
        assert ok
        assert toml_file.read_text(encoding="utf-8").strip() == '[new_section]\nkey = "saved"'
        assert len(reloaded) == 1
        widget.close()

    def test_fallback_save_uses_runtime_atomic_writer(
        self,
        qapp: QApplication,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        from docwen_gui.widgets.settings import toml_editor

        toml_file = tmp_path / "atomic.toml"
        toml_file.write_text("[old]\nvalue = true\n", encoding="utf-8")
        writes: list[tuple[Path, str]] = []

        def atomic_write(path: str | Path, content: str) -> None:
            writes.append((Path(path), content))

        monkeypatch.setattr(toml_editor, "atomic_write_text", atomic_write)
        widget = toml_editor.TomlEditorWidget(
            None,
            config_name="test",
            path_resolver=lambda _name: toml_file,
        )
        assert widget._editor is not None
        widget._editor.setPlainText("[new]\nvalue = true\n")

        assert widget.save_to_disk(show_success=False) is True
        assert writes == [(toml_file, "[new]\nvalue = true\n")]
        assert toml_file.read_text(encoding="utf-8") == "[old]\nvalue = true\n"
        widget.close()

    def test_reload_callback_exception_fails_closed(
        self,
        qapp: QApplication,
        tmp_path: Path,
    ) -> None:
        from docwen_gui.widgets.settings.toml_editor import TomlEditorWidget

        toml_file = tmp_path / "reload_failure.toml"
        toml_file.write_text("", encoding="utf-8")

        def reload_callback() -> None:
            raise RuntimeError("reload failed")

        errors: list[tuple[str, str]] = []
        widget = TomlEditorWidget(
            None,
            config_name="test",
            path_resolver=lambda _name: toml_file,
            reload_callback=reload_callback,
        )
        widget._show_error = lambda title, message: errors.append((title, message))  # type: ignore[method-assign]
        assert widget._editor is not None
        widget._editor.setPlainText("[new]\nvalue = true\n")

        assert widget.save_to_disk(show_success=False) is False
        assert errors and "reload failed" in errors[-1][1]
        widget.close()

    def test_save_invalid_toml(self, qapp: QApplication, tmp_path: Path) -> None:
        from docwen_gui.widgets.settings.toml_editor import TomlEditorWidget

        toml_file = tmp_path / "invalid_test.toml"
        toml_file.write_text("", encoding="utf-8")

        def path_resolver(name: str) -> Path:
            return toml_file

        # Suppress the error dialog that save_to_disk shows on invalid TOML
        widget = TomlEditorWidget(
            None,
            config_name="test",
            path_resolver=path_resolver,
        )
        widget._show_error = lambda title, message: None  # type: ignore[method-assign]
        assert widget._editor is not None
        widget._editor.setPlainText("not valid toml {{{")
        ok = widget.save_to_disk(show_success=False)
        assert not ok
        widget.close()

    def test_reload_from_disk(self, qapp: QApplication, tmp_path: Path) -> None:
        from docwen_gui.widgets.settings.toml_editor import TomlEditorWidget

        toml_file = tmp_path / "reload_test.toml"
        toml_file.write_text("original", encoding="utf-8")

        def path_resolver(name: str) -> Path:
            return toml_file

        widget = TomlEditorWidget(
            None,
            config_name="test",
            path_resolver=path_resolver,
        )
        assert widget._editor is not None
        assert widget._editor.toPlainText() == "original"

        toml_file.write_text("updated", encoding="utf-8")
        widget.reload_from_disk()
        assert widget._editor.toPlainText() == "updated"
        widget.close()

    def test_config_switching(self, qapp: QApplication, tmp_path: Path) -> None:
        from docwen_gui.widgets.settings.toml_editor import TomlEditorWidget

        file_a = tmp_path / "a.toml"
        file_b = tmp_path / "b.toml"
        file_a.write_text("content A", encoding="utf-8")
        file_b.write_text("content B", encoding="utf-8")

        def path_resolver(name: str) -> Path:
            if name == "a":
                return file_a
            return file_b

        widget = TomlEditorWidget(
            None,
            config_name="a",
            choices=[("File A", "a"), ("File B", "b")],
            path_resolver=path_resolver,
        )
        assert widget._editor is not None
        assert widget._editor.toPlainText() == "content A"

        widget.set_config_name("b")
        assert widget._editor.toPlainText() == "content B"
        widget.close()

    def test_file_not_found(self, qapp: QApplication, tmp_path: Path) -> None:
        from docwen_gui.widgets.settings.toml_editor import TomlEditorWidget

        nonexistent = tmp_path / "nonexistent.toml"

        def path_resolver(name: str) -> Path:
            return nonexistent

        widget = TomlEditorWidget(
            None,
            config_name="test",
            path_resolver=path_resolver,
        )
        assert widget._editor is not None
        assert widget._editor.toPlainText() == ""
        widget.close()

    def test_save_callback_path_delegates_and_does_not_touch_path_resolver(
        self, qapp: QApplication, tmp_path: Path
    ) -> None:
        """When save_callback is provided, save_to_disk must delegate to it
        and NOT fall back to path_resolver/write_text."""
        from docwen_gui.widgets.settings.toml_editor import TomlEditorWidget

        saved: list[tuple[str, str]] = []
        reloaded: list[bool] = []

        def save_callback(config_name: str, content: str) -> bool:
            saved.append((config_name, content))
            return True

        def reload_callback() -> None:
            reloaded.append(True)

        # path_resolver points to a file that should NEVER be written when
        # save_callback is in use — if the fallback fired, this assertion breaks
        sentinel = tmp_path / "sentinel.toml"
        sentinel.write_text("untouched", encoding="utf-8")

        def path_resolver(name: str) -> Path:
            return sentinel

        widget = TomlEditorWidget(
            None,
            config_name="myconfig",
            path_resolver=path_resolver,
            reload_callback=reload_callback,
            save_callback=save_callback,
        )
        assert widget._editor is not None
        widget._editor.setPlainText('[x]\nk = "v"\n')
        ok = widget.save_to_disk(show_success=False)
        assert ok
        assert saved == [("myconfig", '[x]\nk = "v"\n')]
        assert len(reloaded) == 1
        # path_resolver file was NOT written by the save
        assert sentinel.read_text(encoding="utf-8") == "untouched"
        widget.close()

    def test_save_callback_returning_false_blocks_reload_and_reports_error(
        self, qapp: QApplication, tmp_path: Path
    ) -> None:
        from docwen_gui.widgets.settings.toml_editor import TomlEditorWidget

        def save_callback(config_name: str, content: str) -> bool:
            return False  # ConfigLoader rejected (e.g. unknown rel_path)

        reloaded: list[bool] = []

        def reload_callback() -> None:
            reloaded.append(True)

        widget = TomlEditorWidget(
            None,
            config_name="myconfig",
            path_resolver=lambda name: tmp_path / "x.toml",
            reload_callback=reload_callback,
            save_callback=save_callback,
        )
        widget._show_error = lambda title, message: None  # type: ignore[method-assign]
        assert widget._editor is not None
        widget._editor.setPlainText('[x]\nk = "v"\n')
        ok = widget.save_to_disk(show_success=False)
        assert ok is False
        assert reloaded == []  # reload must not fire on failed save
        widget.close()

    def test_fallback_atomic_writer_exception_fails_closed(
        self,
        qapp: QApplication,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        from docwen_gui.widgets.settings import toml_editor

        errors: list[tuple[str, str]] = []
        reloaded: list[bool] = []

        def fail_atomic_write(_path: str | Path, _content: str) -> None:
            raise OSError("atomic write failed")

        monkeypatch.setattr(toml_editor, "atomic_write_text", fail_atomic_write)
        widget = toml_editor.TomlEditorWidget(
            None,
            config_name="test",
            path_resolver=lambda _name: tmp_path / "atomic-failure.toml",
            reload_callback=lambda: reloaded.append(True),
        )
        widget._show_error = lambda title, message: errors.append((title, message))  # type: ignore[method-assign]
        assert widget._editor is not None
        widget._editor.setPlainText("[new]\nvalue = true\n")

        assert widget.save_to_disk(show_success=False) is False
        assert errors and "atomic write failed" in errors[-1][1]
        assert reloaded == []
        widget.close()

    def test_save_callback_exception_blocks_reload_and_fails_closed(
        self,
        qapp: QApplication,
        tmp_path: Path,
    ) -> None:
        from docwen_gui.widgets.settings.toml_editor import TomlEditorWidget

        errors: list[tuple[str, str]] = []
        reloaded: list[bool] = []

        def fail_save(_config_name: str, _content: str) -> bool:
            raise RuntimeError("save callback failed")

        widget = TomlEditorWidget(
            None,
            config_name="test",
            path_resolver=lambda _name: tmp_path / "unused.toml",
            reload_callback=lambda: reloaded.append(True),
            save_callback=fail_save,
        )
        widget._show_error = lambda title, message: errors.append((title, message))  # type: ignore[method-assign]
        assert widget._editor is not None
        widget._editor.setPlainText("[new]\nvalue = true\n")

        assert widget.save_to_disk(show_success=False) is False
        assert errors and "save callback failed" in errors[-1][1]
        assert reloaded == []
        widget.close()

    def test_reload_callback_false_fails_closed_after_save(
        self,
        qapp: QApplication,
        tmp_path: Path,
    ) -> None:
        from docwen_gui.widgets.settings.toml_editor import TomlEditorWidget

        errors: list[tuple[str, str]] = []
        saved: list[tuple[str, str]] = []

        def save_callback(config_name: str, content: str) -> bool:
            saved.append((config_name, content))
            return True

        widget = TomlEditorWidget(
            None,
            config_name="test",
            path_resolver=lambda _name: tmp_path / "unused.toml",
            reload_callback=lambda: False,
            save_callback=save_callback,
        )
        widget._show_error = lambda title, message: errors.append((title, message))  # type: ignore[method-assign]
        assert widget._editor is not None
        widget._editor.setPlainText("[new]\nvalue = true\n")

        assert widget.save_to_disk(show_success=False) is False
        assert saved == [("test", "[new]\nvalue = true\n")]
        assert errors
        widget.close()

    def test_successful_save_with_show_success_notifies(
        self,
        qapp: QApplication,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        from docwen_gui.dialogs import feedback
        from docwen_gui.widgets.settings.toml_editor import TomlEditorWidget

        notifications: list[tuple[str, str, str, object]] = []

        def capture_notify(
            level: str,
            title: str,
            message: str,
            *,
            parent: object = None,
            duration_ms: int | None = None,
        ) -> None:
            assert duration_ms is None
            notifications.append((level, title, message, parent))

        monkeypatch.setattr(feedback, "notify", capture_notify)
        widget = TomlEditorWidget(
            None,
            config_name="test",
            path_resolver=lambda _name: tmp_path / "unused.toml",
            save_callback=lambda _name, _content: True,
        )
        assert widget._editor is not None
        widget._editor.setPlainText("[new]\nvalue = true\n")

        assert widget.save_to_disk(show_success=True) is True
        assert len(notifications) == 1
        level, title, message, parent = notifications[0]
        assert level == "success"
        assert title
        assert message
        assert parent is widget
        widget.close()


class TestBatchAddFailedDialog:
    """Test the show_batch_add_failed_dialog function."""

    def test_empty_list_returns_early(self) -> None:
        """Empty failed list should not show any dialog."""
        from docwen_gui.widgets.batch_dialogs import show_batch_add_failed_dialog

        # Should not raise
        show_batch_add_failed_dialog(None, [])

    def test_non_empty_list_creates_dialog(self, qapp: QApplication, monkeypatch: pytest.MonkeyPatch) -> None:
        """Non-empty list should create a localized, identifiable warning dialog."""
        from docwen_gui.widgets.batch_dialogs import show_batch_add_failed_dialog

        captured: list[QMessageBox] = []

        def fake_exec(box: QMessageBox) -> int:
            captured.append(box)
            return int(QMessageBox.StandardButton.Ok)

        monkeypatch.setattr(QMessageBox, "exec", fake_exec)

        failed = [("/path/to/file.md", "Unsupported format")]
        show_batch_add_failed_dialog(None, failed)
        assert len(captured) == 1
        box = captured[0]
        assert box.objectName() == "feedbackWarningMessageBox"
        assert box.windowTitle() == "批量添加失败"
        assert box.text() == "1 个文件无法添加到批量列表。"
        assert "原因" in box.detailedText()
        assert "Unsupported format" in box.detailedText()


class TestDirtyTracking:
    """Test dirty-state behavior across both editors."""

    def test_numbering_add_dirty_on_edit(self, qapp: QApplication) -> None:
        from docwen_gui.widgets.settings.numbering_add_editor import (
            NumberingAddDialog,
        )

        dlg = NumberingAddDialog(config_data={})
        _patch_all_modals(dlg)
        dlg._create_new_scheme()
        dlg._dirty = False  # simulate saved
        dlg._saved_state = dlg._capture_state()
        dlg._refresh_dirty_state()
        assert not dlg._dirty

        dlg.name_edit.setText("Changed Name")
        assert dlg._dirty
        dlg.close()

    def test_numbering_clean_dirty_on_edit(self, qapp: QApplication) -> None:
        from docwen_gui.widgets.settings.numbering_clean_editor import (
            NumberingCleanDialog,
        )

        dlg = NumberingCleanDialog(config_data={})
        _patch_all_modals(dlg)
        dlg._create_new_rule()
        dlg._dirty = False  # simulate saved
        dlg._saved_state = dlg._capture_state()
        dlg._refresh_dirty_state()
        assert not dlg._dirty

        dlg.pattern_edit.setText("^changed_pattern")
        assert dlg._dirty
        dlg.close()

"""Capture and admit one user-triggered clipboard input.

The caller owns the store and its inspection/execution/history leases. This
handler only captures, projects and submits inputs through InputAreaViewModel;
it never publishes conversion outputs or retains an independent input list.
"""

from __future__ import annotations

from collections.abc import Callable
from pathlib import Path
from typing import TYPE_CHECKING

from PySide6.QtWidgets import QApplication, QWidget

from docwen_gui.i18n import t as _t

if TYPE_CHECKING:
    from docwen_gui.clipboard_inputs import ClipboardInputStore
    from docwen_gui.view_models.input_area_vm import InputAreaViewModel


class ClipboardPasteHandler:
    def __init__(
        self,
        *,
        input_area_vm: InputAreaViewModel,
        store_for_paste: Callable[[], ClipboardInputStore],
        busy: Callable[[], bool],
        message: Callable[..., None],
        sync_visible: Callable[[], None],
        parent: QWidget,
    ) -> None:
        self._input_area_vm = input_area_vm
        self._store_for_paste = store_for_paste
        self._busy = busy
        self._message = message
        self._sync_visible = sync_visible
        self._parent = parent

    def paste(self, *, plain_text_only: bool = False) -> None:
        if self._busy():
            self._message(_t("components.file_drop.input_busy"), "warning")
            return
        clipboard = QApplication.clipboard()
        mime_data = clipboard.mimeData()
        if not plain_text_only and mime_data is not None and mime_data.hasUrls():
            file_paths = self._input_area_vm.extract_urls_from_mime_data(mime_data.urls())
            if file_paths:
                self._paste_file_paths(file_paths)
                return

        capture = None
        if not plain_text_only and mime_data is not None:
            from docwen_gui.clipboard_capture import freeze_clipboard_mime

            capture = freeze_clipboard_mime(mime_data)
            if self._try_structured_clipboard_paste(capture):
                return
            if capture.image is not None:
                self._paste_standalone_clipboard_image(capture.image)
                return
            if capture.image_error_code and capture.plain_text is None:
                self._message(
                    _t(
                        "clipboard.image_invalid",
                        "Clipboard image data is invalid or exceeds the supported size.",
                    ),
                    "danger",
                )
                return

        text = capture.plain_text if capture is not None else (mime_data.text() if mime_data is not None else None)
        if text is None:
            self._message(
                _t("components.file_drop.clipboard_non_text", "The clipboard does not contain plain text."),
                "warning",
            )
            return
        if not text.strip():
            self._message(
                _t(
                    "components.file_drop.clipboard_empty",
                    "The clipboard text is empty.",
                ),
                "warning",
            )
            return
        try:
            store = self._store_for_paste()
            snapshot = store.create(
                text,
                display_name_template=_t(
                    "components.file_drop.clipboard_name",
                    "Clipboard Markdown {index}.md",
                    index="{index}",
                ),
            )
        except (OSError, TypeError, ValueError):
            self._message(
                _t(
                    "components.file_drop.clipboard_create_failed",
                    "Clipboard text could not be prepared as a Markdown input.",
                ),
                "danger",
            )
            return

        def completed(outcome) -> None:
            if not outcome.added:
                store.discard_if_unowned(snapshot.path)
            self._sync_visible()

        from docwen_core.detection import inspect_utf8_markdown_snapshot

        self._input_area_vm.add_files(
            [snapshot.path],
            completed=completed,
            file_inspector=inspect_utf8_markdown_snapshot,
        )

    def _try_structured_clipboard_paste(self, capture) -> bool:
        """Handle one already-frozen rich clipboard document."""

        from docwen_core.models.clipboard_document import (
            ClipboardImageRef,
            clipboard_document_preview,
            clipboard_document_to_bytes,
            iter_clipboard_inlines,
        )
        from docwen_gui.clipboard_rich_document import (
            ClipboardRichDocumentError,
            project_frozen_rich_document,
        )

        try:
            decision = project_frozen_rich_document(capture)
        except ClipboardRichDocumentError:
            self._message(
                _t(
                    "clipboard.structured_invalid",
                    "Clipboard table/image structure is invalid and was not imported.",
                ),
                "danger",
            )
            return True
        if decision.plain_fallback:
            self._message(
                _t(
                    "clipboard.structured_plain_fallback",
                    "Rich table/image structure could not be matched reliably to plain text; the complete plain text was kept.",
                ),
                "warning",
            )
            return False
        projection = decision.projection
        if not decision.handled or projection is None:
            return False

        document = projection.document
        try:
            payload = clipboard_document_to_bytes(document)
            store = self._store_for_paste()
            bundle = store.create_bundle(
                payload,
                display_name_template=_t(
                    "components.file_drop.clipboard_document_name",
                    "Clipboard Document {index}.dwclip",
                    index="{index}",
                ),
                preview=clipboard_document_preview(document),
                resources=projection.resources,
            )
        except (OSError, TypeError, ValueError):
            self._message(
                _t(
                    "clipboard.structured_invalid",
                    "Clipboard table/image structure is invalid and was not imported.",
                ),
                "danger",
            )
            return True

        missing_count = sum(
            isinstance(item, ClipboardImageRef) and item.resource_id is None
            for item in iter_clipboard_inlines(document)
        )

        def completed(outcome) -> None:
            if not outcome.added:
                store.discard_if_unowned(bundle.main.path)
            else:
                if projection.html_only:
                    self._message(
                        _t(
                            "clipboard.structured_html_only",
                            "The clipboard had no plain-text body; validated HTML was used as the content source.",
                        ),
                        "warning",
                    )
                if missing_count:
                    self._message(
                        _t("clipboard.images_omitted", count=missing_count),
                        "warning",
                    )
            self._sync_visible()

        from docwen_core.detection import inspect_structured_clipboard_snapshot

        self._input_area_vm.add_files(
            [bundle.main.path],
            completed=completed,
            file_inspector=inspect_structured_clipboard_snapshot,
        )
        return True

    def _paste_standalone_clipboard_image(self, frozen) -> None:
        """Publish one normalized Qt image through the ordinary PNG route."""

        try:
            store = self._store_for_paste()
            snapshot = store.create_binary(
                frozen.payload,
                suffix=".png",
                display_name_template=_t(
                    "components.file_drop.clipboard_image_name",
                    "Clipboard Image {index}.png",
                    index="{index}",
                ),
                preview=f"{frozen.width}×{frozen.height} PNG",
            )
        except (OSError, TypeError, ValueError):
            self._message(
                _t(
                    "clipboard.image_invalid",
                    "Clipboard image data is invalid or exceeds the supported size.",
                ),
                "danger",
            )
            return

        def completed(outcome) -> None:
            if not outcome.added:
                store.discard_if_unowned(snapshot.path)
            self._sync_visible()

        self._input_area_vm.add_files([snapshot.path], completed=completed)

    def _paste_file_paths(self, file_paths: list[str]) -> None:
        """Paste one captured local-file list through normal file admission."""

        needs_batch = self._input_area_vm.mode == "single" and (len(file_paths) != 1 or Path(file_paths[0]).is_dir())
        if not needs_batch:
            self._input_area_vm.add_files(file_paths)
            return

        # Show the ordinary non-mutating rejection first. Declining the mode
        # switch therefore preserves both the current input and its feedback.
        self._input_area_vm.add_files(file_paths)
        from docwen_gui.dialogs.feedback import confirm

        reason = (
            _t("messages.no_folder_in_single_mode", "Single mode does not support folders")
            if len(file_paths) == 1 and Path(file_paths[0]).is_dir()
            else _t("components.file_drop.single_mode_only_one", "Please select exactly one file in single mode")
        )
        if not confirm(
            _t("components.file_drop.batch_mode", "Batch"),
            reason,
            parent=self._parent,
            confirm_label=_t("components.file_drop.batch_mode", "Batch"),
        ):
            return

        # Use the captured list. The clipboard may have changed while the
        # confirmation dialog was open and must not be read a second time.
        self._input_area_vm.set_mode("batch")
        self._input_area_vm.add_files(file_paths)

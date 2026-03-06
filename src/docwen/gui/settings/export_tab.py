from __future__ import annotations

import logging
import tkinter as tk
from typing import TYPE_CHECKING, Any

import ttkbootstrap as tb

from docwen.i18n import t
from docwen.utils.dpi_utils import scale

from .base_tab import BaseSettingsTab
from .config import SectionStyle

logger = logging.getLogger(__name__)

if TYPE_CHECKING:
    from docwen.gui.components.config_combobox import ConfigCombobox


class ExportTab(BaseSettingsTab):
    export_image_ext_mode_combo: "ConfigCombobox | None" = None
    export_ocr_placement_combo: "ConfigCombobox | None" = None

    def _create_interface(self):
        frame = self.create_section_frame(self.scrollable_frame, t("settings.export.md_export_section"), SectionStyle.PRIMARY)

        try:
            ext_mode = self.config_manager.get_export_to_md_image_extraction_mode()
            ocr_place = self.config_manager.get_export_to_md_ocr_placement_mode()
        except Exception:
            ext_mode = "base64"
            ocr_place = "main_md"

        self._create_extraction_mode_selectors(frame, ext_mode, ocr_place, "export")

        self._create_ocr_title_override_section(frame)
        self._create_base64_compress_section(frame)

    def _create_ocr_title_override_section(self, parent: tk.Widget) -> None:
        container = tb.Frame(parent)
        container.pack(fill="x", pady=(scale(10), 0))

        self.create_label_with_info(
            container,
            t("settings.extraction.ocr_blockquote_title_text_label"),
            t("settings.extraction.ocr_blockquote_title_text_tooltip"),
        )

        try:
            from docwen.i18n import t as t_runtime

            default_text = t_runtime("conversion.ocr_output.blockquote_prefix", default="")
        except Exception:
            default_text = ""

        current = ""
        try:
            current = self.config_manager.get_ocr_blockquote_title_override_text() or ""
        except Exception:
            current = ""

        self.ocr_title_text_var = tk.StringVar(value=current or default_text)

        row = tb.Frame(container)
        row.pack(fill="x")

        entry = tb.Entry(row, textvariable=self.ocr_title_text_var)
        entry.pack(side="left", fill="x", expand=True)

        reset_btn = tb.Button(
            row,
            text=t("settings.extraction.ocr_blockquote_title_text_reset"),
            bootstyle="secondary",
            command=lambda: self._reset_ocr_title_text(default_text),
        )
        reset_btn.pack(side="left", padx=(scale(8), 0))

    def _reset_ocr_title_text(self, default_text: str) -> None:
        if getattr(self, "ocr_title_text_var", None) is not None:
            self.ocr_title_text_var.set(default_text)

    def _create_base64_compress_section(self, parent: tk.Widget) -> None:
        frame = tb.Frame(parent)
        frame.pack(fill="x", pady=(scale(10), 0))

        self.create_label_with_info(
            frame,
            t("settings.export.base64_compress_section"),
            t("settings.export.base64_compress_tooltip"),
        )

        enabled = bool(self.config_manager.get_export_base64_compress_enabled())
        threshold_kb = int(self.config_manager.get_export_base64_compress_threshold_kb())

        self.base64_compress_enabled_var = tk.BooleanVar(value=enabled)
        self.base64_threshold_var = tk.IntVar(value=threshold_kb)

        enabled_toggle = tb.Checkbutton(
            frame,
            text=t("settings.export.base64_compress_enabled"),
            variable=self.base64_compress_enabled_var,
            bootstyle="round-toggle",
        )
        enabled_toggle.pack(anchor="w", pady=(0, scale(8)))

        threshold_row = tb.Frame(frame)
        threshold_row.pack(fill="x", pady=(0, scale(8)))
        tb.Label(threshold_row, text=t("settings.export.base64_compress_threshold_label"), bootstyle="secondary").pack(
            side="left"
        )
        tb.Spinbox(
            threshold_row,
            from_=50,
            to=5000,
            increment=10,
            textvariable=self.base64_threshold_var,
            width=10,
        ).pack(side="left", padx=(scale(8), 0))
        tb.Label(threshold_row, text="KB", bootstyle="secondary").pack(side="left", padx=(scale(6), 0))

    def get_settings(self) -> dict[str, Any]:
        ext_mode = "base64"
        ocr_place = "main_md"
        if self.export_image_ext_mode_combo is not None:
            ext_mode = self.export_image_ext_mode_combo.get_config_value()
        if self.export_ocr_placement_combo is not None:
            ocr_place = self.export_ocr_placement_combo.get_config_value()

        title_enabled = self._get_ocr_blockquote_title_enabled_setting()

        title_text = ""
        if getattr(self, "ocr_title_text_var", None) is not None:
            title_text = str(self.ocr_title_text_var.get() or "")

        compress_enabled = bool(getattr(self, "base64_compress_enabled_var", tk.BooleanVar(value=True)).get())
        threshold_kb = int(getattr(self, "base64_threshold_var", tk.IntVar(value=200)).get())

        return {
            "to_md_image_extraction_mode": ext_mode,
            "to_md_ocr_placement_mode": ocr_place,
            "ocr_blockquote_title_enabled": title_enabled,
            "ocr_blockquote_title_text": title_text,
            "base64_compress_enabled": compress_enabled,
            "base64_compress_threshold_kb": threshold_kb,
        }

    def apply_settings(self) -> bool:
        try:
            settings = self.get_settings()
            ok = True

            for key in ["to_md_image_extraction_mode", "to_md_ocr_placement_mode"]:
                if not self.config_manager.update_config_value("conversion_defaults", "export", key, settings[key]):
                    ok = False

            if not self.config_manager.update_config_value(
                "conversion_config",
                "ocr_output",
                "show_blockquote_title",
                bool(settings["ocr_blockquote_title_enabled"]),
            ):
                ok = False

            locale = None
            try:
                locale = self.config_manager.get_locale()
            except Exception:
                locale = None

            if locale:
                raw = settings.get("ocr_blockquote_title_text") or ""
                try:
                    from docwen.i18n import t as t_runtime

                    default_text = str(t_runtime("conversion.ocr_output.blockquote_prefix", default="") or "")
                except Exception:
                    default_text = ""

                ocr_output = self.config_manager.get_ocr_output_config() or {}
                overrides = ocr_output.get("blockquote_title_override_by_locale", {})
                if not isinstance(overrides, dict):
                    overrides = {}

                if raw.strip() and raw.strip() != default_text.strip():
                    overrides[locale] = raw
                else:
                    overrides.pop(locale, None)

                if not self.config_manager.update_config_value(
                    "conversion_config",
                    "ocr_output",
                    "blockquote_title_override_by_locale",
                    overrides,
                ):
                    ok = False

            if not self.config_manager.update_config_value(
                "conversion_config",
                "export",
                "base64_compress_enabled",
                bool(settings["base64_compress_enabled"]),
            ):
                ok = False
            if not self.config_manager.update_config_value(
                "conversion_config",
                "export",
                "base64_compress_threshold_kb",
                int(settings["base64_compress_threshold_kb"]),
            ):
                ok = False

            return ok
        except Exception as e:
            logger.error(f"应用导出设置失败: {e}", exc_info=True)
            return False

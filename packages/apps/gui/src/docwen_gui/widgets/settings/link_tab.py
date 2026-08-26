"""Link settings tab — link format, non-embed, embed, embedding depth.

Matches old LinkTab (7 QComboBox + 1 QSpinBox, 3 cards).
"""

from __future__ import annotations

from typing import cast as _cast

from PySide6.QtWidgets import QComboBox, QSpinBox

from ...i18n import t
from ...view_models.settings_vm import SECTION_LINK, SettingsViewModel
from .base_tab import BaseSettingsTab


def _link_modes() -> list[tuple[str, str]]:
    return [
        (t("settings.link.modes.keep", "Keep Original"), "keep"),
        (t("settings.link.modes.extract_text", "Extract Text"), "extract_text"),
        (t("settings.link.modes.remove", "Remove Link"), "remove"),
        (t("settings.link.modes.hyperlink", "Hyperlink"), "hyperlink"),
    ]


def _embed_modes() -> list[tuple[str, str]]:
    return [
        (t("settings.link.modes.keep", "Keep Original"), "keep"),
        (t("settings.link.modes.extract_text", "Extract Text"), "extract_text"),
        (t("settings.link.modes.remove", "Remove Link"), "remove"),
        (t("settings.link.modes.embed", "Embed Content"), "embed"),
    ]


class LinkTab(BaseSettingsTab):
    """Link processing settings tab."""

    def __init__(self, view_model: SettingsViewModel) -> None:
        self._vm = view_model
        self._image_link_style: QComboBox = _cast(QComboBox, None)
        self._md_file_link_style: QComboBox = _cast(QComboBox, None)
        self._wiki_link_mode: QComboBox = _cast(QComboBox, None)
        self._md_link_mode: QComboBox = _cast(QComboBox, None)
        self._wiki_embed_image_mode: QComboBox = _cast(QComboBox, None)
        self._md_embed_image_mode: QComboBox = _cast(QComboBox, None)
        self._embed_md_file_mode: QComboBox = _cast(QComboBox, None)
        self._max_depth: QSpinBox
        super().__init__()
        self._load_values()

    def _create_interface(self) -> None:
        # ── Link format card ────────────────────────────────────────────
        _c1, f1 = self.add_settings_card(
            t("settings.link.format_section", "Link Format"),
            t("settings.link.format_desc", "Configure how links are rendered in output."),
            object_name="linkFormatCard",
        )
        self._image_link_style = self.create_combobox(
            [
                (t("settings.link.image_styles.wiki_embed", "Wiki Embed"), "wiki_embed"),
                (t("settings.link.image_styles.wiki_link", "Wiki Link"), "wiki_link"),
                (t("settings.link.image_styles.md_embed", "Markdown Embed"), "markdown_embed"),
                (t("settings.link.image_styles.md_link", "Markdown Link"), "markdown_link"),
            ],
            t("settings.link.image_link_style_tooltip", "Image link style in Markdown output"),
        )
        self.add_form_row(f1, t("settings.link.image_link_style_label", "Image Link Style:"), self._image_link_style)

        self._md_file_link_style = self.create_combobox(
            [
                (t("settings.link.md_file_styles.wiki_embed", "Wiki Embed"), "wiki_embed"),
                (t("settings.link.md_file_styles.wiki_link", "Wiki Link"), "wiki_link"),
                (t("settings.link.md_file_styles.md_link", "Markdown Link"), "markdown_link"),
            ],
            t("settings.link.md_file_link_style_tooltip", "MD file link style in Markdown output"),
        )
        self.add_form_row(
            f1, t("settings.link.md_file_link_style_label", "MD File Link Style:"), self._md_file_link_style
        )

        # ── Non-embed links card ────────────────────────────────────────
        _c2, f2 = self.add_settings_card(
            t("settings.link.non_embed_section", "Non-Embedded Links"),
            t("settings.link.non_embed_desc", "How to handle non-embedded links."),
            object_name="linkNonEmbedCard",
        )
        self._wiki_link_mode = self.create_combobox(
            _link_modes(), t("settings.link.wiki_link_tooltip", "How to process wiki-style links")
        )
        self.add_form_row(f2, t("settings.link.wiki_link_mode", "Wiki Link Mode:"), self._wiki_link_mode)

        self._md_link_mode = self.create_combobox(
            _link_modes(), t("settings.link.markdown_link_tooltip", "How to process Markdown-style links")
        )
        self.add_form_row(f2, t("settings.link.markdown_link_mode", "Markdown Link Mode:"), self._md_link_mode)

        # ── Embed links card ────────────────────────────────────────────
        _c3, f3 = self.add_settings_card(
            t("settings.link.embed_section", "Embedded Links"),
            t("settings.link.embed_desc", "How to handle embedded links and images."),
            object_name="linkEmbedCard",
        )
        self._wiki_embed_image_mode = self.create_combobox(
            _embed_modes(), t("settings.link.wiki_embed_image_tooltip", "How to process wiki-embedded images")
        )
        self.add_form_row(
            f3, t("settings.link.wiki_embed_image_mode", "Wiki Embed Image Mode:"), self._wiki_embed_image_mode
        )

        self._md_embed_image_mode = self.create_combobox(
            _embed_modes(), t("settings.link.markdown_embed_image_tooltip", "How to process Markdown-embedded images")
        )
        self.add_form_row(
            f3, t("settings.link.markdown_embed_image_mode", "Markdown Embed Image Mode:"), self._md_embed_image_mode
        )

        self._embed_md_file_mode = self.create_combobox(
            _embed_modes(), t("settings.link.embed_md_file_tooltip", "How to process embedded MD files")
        )
        self.add_form_row(f3, t("settings.link.embed_md_file_mode", "Embed MD File Mode:"), self._embed_md_file_mode)

        self._max_depth = self.create_spinbox(
            1, 20, t("settings.link.max_depth_tooltip", "Maximum link embedding depth (1-20)"), default=3
        )
        self.add_form_row(f3, t("settings.link.max_depth_label", "Max Embedding Depth:"), self._max_depth)

        # Wire signals
        self._image_link_style.currentIndexChanged.connect(lambda _: self._on_change("image_link_style"))
        self._md_file_link_style.currentIndexChanged.connect(lambda _: self._on_change("md_file_link_style"))
        self._wiki_link_mode.currentIndexChanged.connect(lambda _: self._on_change("wiki_link_mode"))
        self._md_link_mode.currentIndexChanged.connect(lambda _: self._on_change("markdown_link_mode"))
        self._wiki_embed_image_mode.currentIndexChanged.connect(lambda _: self._on_change("wiki_embed_image_mode"))
        self._md_embed_image_mode.currentIndexChanged.connect(lambda _: self._on_change("markdown_embed_image_mode"))
        self._embed_md_file_mode.currentIndexChanged.connect(lambda _: self._on_change("embed_md_file_mode"))
        self._max_depth.valueChanged.connect(lambda value: self._vm.set_field(SECTION_LINK, "max_depth", value))

    def _on_change(self, key: str) -> None:
        combo_map = {
            "image_link_style": self._image_link_style,
            "md_file_link_style": self._md_file_link_style,
            "wiki_link_mode": self._wiki_link_mode,
            "markdown_link_mode": self._md_link_mode,
            "wiki_embed_image_mode": self._wiki_embed_image_mode,
            "markdown_embed_image_mode": self._md_embed_image_mode,
            "embed_md_file_mode": self._embed_md_file_mode,
        }
        combo = combo_map.get(key)
        if combo:
            self._vm.set_field(SECTION_LINK, key, self.get_combo_data(combo))

    def _load_values(self) -> None:
        link = self._vm.config.link
        self.set_combo_data(self._image_link_style, link.image_link_style)
        self.set_combo_data(self._md_file_link_style, link.md_file_link_style)
        self.set_combo_data(self._wiki_link_mode, link.wiki_link_mode)
        self.set_combo_data(self._md_link_mode, link.markdown_link_mode)
        self.set_combo_data(self._wiki_embed_image_mode, link.wiki_embed_image_mode)
        self.set_combo_data(self._md_embed_image_mode, link.markdown_embed_image_mode)
        self.set_combo_data(self._embed_md_file_mode, link.embed_md_file_mode)
        self._max_depth.setValue(link.max_depth)

    def reload_from_config(self) -> None:
        self._load_values()

"""Display changes cannot silently change an admitted task's language."""

from __future__ import annotations

from pathlib import Path

import pytest

from docwen_core.models.request import ConversionRequest, FileRef, OutputPolicy
from docwen_runtime._request_admission import project_markdown_ocr_options
from docwen_runtime.config import build_document_style_catalog
from docwen_runtime.config.loader import ConfigLoader
from docwen_runtime.config.ocr_output import build_ocr_blockquote_title
from docwen_runtime.config.task_language import output_yaml_labels

pytestmark = pytest.mark.contract
CONFIGS = Path(__file__).resolve().parents[3] / "configs"


@pytest.mark.parametrize("display", ["zh_CN", "en_US", "ar_SA", "ja_JP"])
def test_fixed_task_language_is_independent_of_display(display):
    snapshot = {
        "gui": {"language": {"locale": display}},
        "conversion": {"output_locale": "en_US"},
        "ocr": {"language": "japanese"},
    }
    catalog = build_document_style_catalog(snapshot)
    assert catalog.locale == "en_US"
    assert catalog.name_for("figure_caption") == "Figure Caption"
    assert build_ocr_blockquote_title(snapshot) == build_ocr_blockquote_title({}, requested_locale="en_US")
    request = ConversionRequest(
        request_id="language-independence",
        input_refs=[FileRef(path="input.docx", format="docx", category="document")],
        target_format="md",
        options={},
        output_policy=OutputPolicy(),
    )
    assert project_markdown_ocr_options(request, snapshot) == {"locale": "en_US", "ocr_language": "japanese"}
    assert request.options == {}


def test_explicit_task_locale_wins_without_changing_snapshot():
    snapshot = {"conversion": {"output_locale": "en_US"}}
    assert build_document_style_catalog(snapshot, request_options={"locale": "ar_SA"}).locale == "ar_SA"
    assert snapshot == {"conversion": {"output_locale": "en_US"}}
    assert output_yaml_labels("de_DE") == {"title": "Titel", "subtitle": "Untertitel"}


def test_legacy_profile_freezes_previous_language_once(tmp_path):
    user = tmp_path / "user"
    user.mkdir()
    (user / "gui.toml").write_text('[language]\nlocale="en_US"\n', encoding="utf-8")
    loader = ConfigLoader(base_dir=CONFIGS, user_dir=user)
    assert loader.config.conversion.output_locale == "en_US"
    saved = (user / "conversion.toml").read_bytes()
    loader.reload()
    assert (user / "conversion.toml").read_bytes() == saved
    assert loader.set_value("gui.language.locale", "ar_SA")
    assert loader.config.gui.language.locale == "ar_SA"
    assert loader.config.conversion.output_locale == "en_US"
    assert (user / "conversion.toml").read_bytes() == saved


def test_new_profile_ui_change_preserves_default_task_language(tmp_path):
    user = tmp_path / "user"
    loader = ConfigLoader(base_dir=CONFIGS, user_dir=user)
    assert not (user / "conversion.toml").exists()
    assert loader.config.conversion.output_locale == "zh_CN"
    assert loader.set_value("gui.language.locale", "ar_SA")
    assert loader.config.conversion.output_locale == "zh_CN"
    assert loader.set_value("conversion.output_locale", "en_US")
    assert loader.config.gui.language.locale == "ar_SA"
    assert loader.config.conversion.output_locale == "en_US"

"""Check fictional screenshot inputs without opening the GUI or modifying templates."""

from pathlib import Path
from xml.etree import ElementTree
from zipfile import ZipFile

import pytest
from scripts.maintenance.render_screenshots import write_sample

pytestmark = pytest.mark.contract

_SAMPLE_NAMES = {
    "ar_SA": "تقرير المشروع",
    "de_DE": "Projektbericht",
    "en_US": "Project report",
    "es_ES": "Informe del proyecto",
    "fr_FR": "Rapport de projet",
    "ja_JP": "プロジェクト報告",
    "ko_KR": "프로젝트 보고서",
    "pt_BR": "Relatório do projeto",
    "ru_RU": "Отчёт по проекту",
    "vi_VN": "Báo cáo dự án",
    "zh_CN": "项目周报",
    "zh_TW": "專案週報",
}


def test_screenshot_samples_cover_every_application_language() -> None:
    locales = Path(__file__).resolve().parents[2] / "i18n/locales"
    assert set(_SAMPLE_NAMES) == {path.stem for path in locales.glob("*.toml")}


@pytest.mark.parametrize("locale,title", _SAMPLE_NAMES.items())
def test_screenshot_document_and_markdown_use_matching_localized_content(
    tmp_path: Path, locale: str, title: str
) -> None:
    markdown = write_sample(tmp_path, locale, document=False)
    document = write_sample(tmp_path, locale, document=True)
    assert markdown.name == f"{title}.md"
    assert document.name == f"{title}.docx"
    text = markdown.read_text(encoding="utf-8")
    assert text.startswith(f"# {title}\n")
    with ZipFile(document) as package:
        body = ElementTree.fromstring(package.read("word/document.xml"))
    namespace = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    contents = [node.text for node in body.findall(".//w:t", namespace)]
    assert len(contents) == 7
    assert contents[0] == title
    assert all(value and value in text for value in contents)
    if locale != "en_US":
        assert "Project report" not in text
        assert "Document review" not in text


def test_unknown_screenshot_language_is_not_silently_rendered_in_english(tmp_path: Path) -> None:
    with pytest.raises(KeyError):
        write_sample(tmp_path, "unknown", document=False)
    assert list(tmp_path.iterdir()) == []

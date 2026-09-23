"""Request-local DOCX comment presentation; rule identities never depend on language."""

from __future__ import annotations

from collections.abc import Mapping
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from docwen_plugin_proofread.text_validator import TextError

# typo label, symbol correction label, sensitive notice, unmatched notice, crossed notice
_TEXT: dict[str, tuple[str, str, str, str, str]] = {
    "zh_CN": (
        "错别字",
        "符号修正",
        "命中敏感词：‘{text}’。请结合上下文核查。",
        "‘{text}’未找到匹配符号，请核查。",
        "‘{text}’处存在交叉嵌套，请核查符号配对顺序。",
    ),
    "zh_TW": (
        "錯別字",
        "符號修正",
        "命中敏感詞：‘{text}’。請結合上下文核查。",
        "‘{text}’未找到配對符號，請核查。",
        "‘{text}’處存在交叉巢狀，請核查符號配對順序。",
    ),
    "en_US": (
        "Typo",
        "Symbol correction",
        "Sensitive word found: ‘{text}’. Review it in context.",
        "‘{text}’ has no matching symbol. Please check.",
        "‘{text}’ forms crossed nesting. Check the pairing order.",
    ),
    "de_DE": (
        "Tippfehler",
        "Zeichenkorrektur",
        "Sensibles Wort gefunden: ‘{text}’. Bitte im Kontext prüfen.",
        "Für ‘{text}’ fehlt das passende Zeichen. Bitte prüfen.",
        "Bei ‘{text}’ überkreuzen sich Zeichenpaare. Bitte die Reihenfolge prüfen.",
    ),
    "es_ES": (
        "Error tipográfico",
        "Corrección de símbolo",
        "Palabra sensible detectada: ‘{text}’. Revísela en su contexto.",
        "‘{text}’ no tiene símbolo de cierre o apertura correspondiente. Revíselo.",
        "‘{text}’ forma un anidamiento cruzado. Revise el orden de los símbolos.",
    ),
    "fr_FR": (
        "Faute de frappe",
        "Correction de symbole",
        "Mot sensible détecté : ‘{text}’. Vérifiez-le dans son contexte.",
        "‘{text}’ n’a pas de symbole correspondant. Veuillez vérifier.",
        "Les paires de symboles se croisent à ‘{text}’. Vérifiez leur ordre.",
    ),
    "ja_JP": (
        "誤字",
        "記号の修正",
        "要確認語句：‘{text}’。文脈に照らして確認してください。",
        "‘{text}’に対応する記号がありません。確認してください。",
        "‘{text}’で入れ子の順序が交差しています。記号の対応を確認してください。",
    ),
    "ko_KR": (
        "오타",
        "기호 수정",
        "민감한 단어 발견: ‘{text}’. 문맥을 확인하세요.",
        "‘{text}’에 대응하는 기호가 없습니다. 확인하세요.",
        "‘{text}’에서 중첩 순서가 교차합니다. 기호의 대응 순서를 확인하세요.",
    ),
    "pt_BR": (
        "Erro de digitação",
        "Correção de símbolo",
        "Palavra sensível encontrada: ‘{text}’. Verifique o contexto.",
        "‘{text}’ não tem símbolo correspondente. Verifique.",
        "‘{text}’ apresenta aninhamento cruzado. Verifique a ordem dos pares.",
    ),
    "ru_RU": (
        "Опечатка",
        "Исправление символа",
        "Найдено чувствительное слово: ‘{text}’. Проверьте контекст.",
        "Для ‘{text}’ не найден парный символ. Проверьте текст.",
        "У ‘{text}’ нарушен порядок вложенности. Проверьте порядок парных символов.",
    ),
    "vi_VN": (
        "Lỗi chính tả",
        "Sửa ký hiệu",
        "Phát hiện từ nhạy cảm: ‘{text}’. Hãy kiểm tra ngữ cảnh.",
        "‘{text}’ không có ký hiệu tương ứng. Hãy kiểm tra.",
        "‘{text}’ có thứ tự lồng nhau bị chéo. Hãy kiểm tra thứ tự các cặp ký hiệu.",
    ),
}
SUPPORTED_COMMENT_LOCALES = tuple(_TEXT)


def normalize_comment_locale(value: object) -> str:
    candidate = str(value or "").replace("-", "_")
    if candidate in _TEXT:
        return candidate
    aliases = {locale.split("_")[0]: locale for locale in _TEXT}
    aliases["zh"] = "zh_CN"
    return aliases.get(candidate.lower(), "en_US")


def resolve_comment_locale(context: object) -> str:
    request = getattr(context, "request", None)
    options = getattr(request, "options", {})
    if isinstance(options, Mapping):
        explicit = options.get("locale") or options.get("lang")
        if explicit:
            return normalize_comment_locale(explicit)
    snapshot = getattr(request, "config_snapshot", {})
    gui = snapshot.get("gui", {}) if isinstance(snapshot, Mapping) else {}
    language = gui.get("language", {}) if isinstance(gui, Mapping) else {}
    configured = language.get("locale") if isinstance(language, Mapping) else language
    if not configured:
        get_config = getattr(getattr(context, "config", None), "get", None)
        configured = get_config("gui.language.locale", "en_US") if callable(get_config) else "en_US"
    return normalize_comment_locale(configured)


def format_comment(error: TextError, locale: str) -> str:
    language = normalize_comment_locale(locale)
    typo, symbol, sensitive, unmatched, crossed = _TEXT[language]
    if error.replacement is not None:
        label = typo if error.source == "typo" else symbol
        separator = "：" if language in {"zh_CN", "zh_TW", "ja_JP"} else ": "
        return f"{label}{separator}{error.error_text} → {error.replacement}"
    template = sensitive if error.source == "sensitive" else crossed if error.pairing_reason == "crossed" else unmatched
    return template.format(text=error.error_text)

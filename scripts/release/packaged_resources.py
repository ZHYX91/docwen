from __future__ import annotations

from pathlib import Path

from docwen_runtime.pymupdf_layout_resources import (
    pymupdf_layout_resource_paths,
    verify_installed_pymupdf_layout_distribution,
    verify_pymupdf_layout_resource_root,
)

RESOURCE_DIRS = ("templates", "configs", "models")
REQUIRED_CONFIG_FILES = (
    "conversion.toml",
    "document.toml",
    "export.toml",
    "field_processors.toml",
    "gui.toml",
    "image.toml",
    "layout.toml",
    "link.toml",
    "logger.toml",
    "numbering/add.toml",
    "numbering/cleanup.toml",
    "optimize.toml",
    "other.toml",
    "output.toml",
    "proofread/engine.toml",
    "proofread/pairs.toml",
    "proofread/sensitive_words.toml",
    "proofread/skip.toml",
    "proofread/symbol_map.toml",
    "proofread/typos.toml",
    "software.toml",
    "spreadsheet.toml",
    "text.toml",
)
REQUIRED_TEMPLATE_FILES = (
    "Deutsche Allgemeine Vorlage.docx",
    "English General Template.docx",
    "English Sample Sheet Template.xlsx",
    "Modelo Geral Português.docx",
    "Modèle Général Français.docx",
    "Mẫu Chung Tiếng Việt.docx",
    "Plantilla General Española.docx",
    "Русский Общий Шаблон.docx",
    "日本語汎用テンプレート.docx",
    "简体中文通用模板.docx",
    "繁體中文通用模板.docx",
    "한국어 범용 템플릿.docx",
)
REQUIRED_MODEL_FILES = (
    "rapidocr/arabic_PP-OCRv4_rec_infer.onnx",
    "rapidocr/ch_PP-OCRv4_det_infer.onnx",
    "rapidocr/ch_PP-OCRv4_rec_infer.onnx",
    "rapidocr/ch_ppocr_mobile_v2.0_cls_infer.onnx",
    "rapidocr/chinese_cht_PP-OCRv3_rec_infer.onnx",
    "rapidocr/cyrillic_PP-OCRv3_rec_infer.onnx",
    "rapidocr/en_PP-OCRv4_rec_infer.onnx",
    "rapidocr/japan_PP-OCRv4_rec_infer.onnx",
    "rapidocr/korean_PP-OCRv4_rec_infer.onnx",
    "rapidocr/latin_PP-OCRv3_rec_infer.onnx",
)
REQUIRED_ASSET_FILES = (
    "icon.ico",
    "icon.png",
    "icon.svg",
    "about_icon.png",
    "complete_icon.png",
    "fail_icon.png",
    "location_icon.png",
    "move_top_icon.png",
    "remove_icon.png",
    "settings_icon.png",
    "skip_icon.png",
    "file_drop_empty_state.svg",
    "icons/about.svg",
    "icons/complete.svg",
    "icons/document.svg",
    "icons/error.svg",
    "icons/export.svg",
    "icons/font_size.svg",
    "icons/formatting.svg",
    "icons/general.svg",
    "icons/image.svg",
    "icons/info.svg",
    "icons/layout.svg",
    "icons/link.svg",
    "icons/logging.svg",
    "icons/open_folder.svg",
    "icons/other.svg",
    "icons/output.svg",
    "icons/proofread.svg",
    "icons/settings.svg",
    "icons/skip.svg",
    "icons/spreadsheet.svg",
    "icons/sync.svg",
    "icons/text.svg",
)
LOCALE_DIR_CANDIDATES = (
    "i18n/locales",
    "_internal/docwen/i18n/locales",
    "docwen/i18n/locales",
)
REQUIRED_LOCALE_FILES = (
    "de_DE.toml",
    "en_US.toml",
    "es_ES.toml",
    "fr_FR.toml",
    "ja_JP.toml",
    "ko_KR.toml",
    "pt_BR.toml",
    "ru_RU.toml",
    "vi_VN.toml",
    "zh_CN.toml",
    "zh_TW.toml",
)

PYMUPDF_LAYOUT_PACKAGED_RESOURCE_ROOT = Path("_internal/pymupdf/layout/resources")

_FORBIDDEN_QT_NETWORK_FILENAMES = {
    "qt6network.dll",
    "qt6networkauth.dll",
    "qt6websockets.dll",
    "qtnetwork.pyd",
    "qtnetworkauth.pyd",
    "qtwebsockets.pyd",
}
_FORBIDDEN_QT_NETWORK_PREFIXES = (
    "libqt6network.so",
    "libqt6networkauth.so",
    "libqt6websockets.so",
)
_FORBIDDEN_QT_PLUGIN_DIRS = {"networkinformation", "tls"}


def missing_files(path: Path, required: tuple[str, ...]) -> list[str]:
    return [name for name in required if not (path / name).is_file()]


def pymupdf_layout_source_resource_files() -> tuple[str, ...]:
    """Verify and return the explicitly pinned source resource contract."""

    verification = verify_installed_pymupdf_layout_distribution()
    if not verification.available:
        raise RuntimeError(f"pymupdf_layout_source_contract_failed:{verification.reason}")
    return pymupdf_layout_resource_paths()


def verify_pymupdf_layout_resource_layout(binary_dir: Path, *, error_prefix: str) -> None:
    resource_root = binary_dir / PYMUPDF_LAYOUT_PACKAGED_RESOURCE_ROOT
    verification = verify_pymupdf_layout_resource_root(resource_root)
    if not verification.available:
        raise RuntimeError(f"{error_prefix}_pymupdf_layout_resources_invalid:{verification.reason}")


def verify_no_bundled_qt_network_stack(binary_dir: Path, *, error_prefix: str) -> None:
    """Reject unused native Qt networking surfaces from frozen candidates."""

    forbidden: list[str] = []
    for path in binary_dir.rglob("*"):
        relative = path.relative_to(binary_dir)
        lowered_parts = tuple(part.lower() for part in relative.parts)
        lowered_name = path.name.lower()
        if path.is_dir() and "plugins" in lowered_parts and lowered_name in _FORBIDDEN_QT_PLUGIN_DIRS:
            forbidden.append(relative.as_posix())
            continue
        if not path.is_file():
            continue
        if (
            lowered_name in _FORBIDDEN_QT_NETWORK_FILENAMES
            or lowered_name.startswith(_FORBIDDEN_QT_NETWORK_PREFIXES)
            or "qtnetwork.framework" in lowered_parts
            or "qtwebsockets.framework" in lowered_parts
        ):
            forbidden.append(relative.as_posix())
    if forbidden:
        raise RuntimeError(f"{error_prefix}_forbidden_qt_network_stack: {sorted(forbidden)}")


def verify_common_resource_layout(binary_dir: Path, *, error_prefix: str) -> None:
    verify_no_bundled_qt_network_stack(binary_dir, error_prefix=error_prefix)

    missing_dirs = [name for name in RESOURCE_DIRS if not (binary_dir / name).is_dir()]
    if missing_dirs:
        raise RuntimeError(f"{error_prefix}_resources_missing: {missing_dirs}")

    missing_templates = missing_files(binary_dir / "templates", REQUIRED_TEMPLATE_FILES)
    if missing_templates:
        raise RuntimeError(f"{error_prefix}_templates_missing: {missing_templates}")

    missing_configs = missing_files(binary_dir / "configs", REQUIRED_CONFIG_FILES)
    if missing_configs:
        raise RuntimeError(f"{error_prefix}_configs_missing: {missing_configs}")

    missing_models = missing_files(binary_dir / "models", REQUIRED_MODEL_FILES)
    if missing_models:
        raise RuntimeError(f"{error_prefix}_models_missing: {missing_models}")

    locale_dirs = [binary_dir / rel_path for rel_path in LOCALE_DIR_CANDIDATES]
    if not any(path.is_dir() and not missing_files(path, REQUIRED_LOCALE_FILES) for path in locale_dirs):
        missing_by_dir = {
            str(path): missing_files(path, REQUIRED_LOCALE_FILES) if path.is_dir() else list(REQUIRED_LOCALE_FILES)
            for path in locale_dirs
        }
        raise RuntimeError(f"{error_prefix}_locales_missing: {missing_by_dir}")

    verify_pymupdf_layout_resource_layout(binary_dir, error_prefix=error_prefix)

"""Fail-closed guards for the current-only documentation tree."""

from __future__ import annotations

import re
from pathlib import Path
from urllib.parse import unquote

import pytest

pytestmark = pytest.mark.unit

ROOT = Path(__file__).resolve().parents[2]
DOCS = ROOT / "docs"

REQUIRED_DOCS = {
    "README.md",
    "overview.md",
    "capabilities.md",
    "cli.md",
    "configuration.md",
    "architecture.md",
    "runtime-artifacts.md",
    "external-dependencies.md",
    "testing.md",
    "packaging.md",
    "development.md",
    "specs/routes-and-actions.md",
    "specs/plugin-manifest.md",
    "specs/json-contracts.md",
    "specs/json-contracts.schema.json",
    "specs/golden-regression-suite.md",
    "specs/gui-behavior.md",
    "specs/markdown-compatibility.md",
    "specs/structured-numbering-phases.md",
    "specs/templates-and-styles.md",
    "specs/public-api-boundaries.md",
    "maintenance/troubleshooting.md",
    "maintenance/docs-style-guide.md",
}

LOCALIZED_READMES = (
    "README.ar.md",
    "README.de-DE.md",
    "README.es-ES.md",
    "README.fr-FR.md",
    "README.ja-JP.md",
    "README.ko-KR.md",
    "README.pt-BR.md",
    "README.ru-RU.md",
    "README.vi-VN.md",
    "README.zh-CN.md",
    "README.zh-TW.md",
)

LANGUAGE_NAV = (
    "[English](https://github.com/ZHYX91/docwen/blob/main/README.md)"
    " · [简体中文](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.zh-CN.md)"
    " · [繁體中文](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.zh-TW.md)"
    " · [Deutsch](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.de-DE.md)"
    " · [Français](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.fr-FR.md)"
    " · [Español](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.es-ES.md)"
    " · [Português](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.pt-BR.md)"
    " · [Русский](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.ru-RU.md)"
    " · [日本語](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.ja-JP.md)"
    " · [한국어](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.ko-KR.md)"
    " · [Tiếng Việt](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.vi-VN.md)"
    " · [العربية](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.ar.md)"
)
LOGO_BLOCK = (
    '<p align="center">\n'
    '  <img src="https://raw.githubusercontent.com/ZHYX91/docwen/main/assets/icon.svg" '
    'alt="DocWen logo" width="120">\n'
    "</p>"
)

NETWORK_GUARD_IP_OPERATIONS = ("bind", "connect", "connect_ex", "sendto", "sendmsg")

NETWORK_GUARD_EXTERNAL_PROCESS_MARKERS = {
    "README.ar.md": "هذا القيد لا يشملها",
    "README.md": "They are outside this restriction",
    "README.de-DE.md": "Sie fallen nicht unter diese Beschränkung",
    "README.es-ES.md": "Esta restricción no los cubre",
    "README.fr-FR.md": "Cette restriction ne les couvre pas",
    "README.ja-JP.md": "この制限の対象ではないため",
    "README.ko-KR.md": "이 제한의 적용 대상이 아닙니다",
    "README.pt-BR.md": "Essa restrição não se aplica a eles",
    "README.ru-RU.md": "Это ограничение на них не распространяется",
    "README.vi-VN.md": "Hạn chế này không áp dụng cho chúng",
    "README.zh-CN.md": "不受这项限制",
    "README.zh-TW.md": "不受這項限制",
}

SCREENSHOTS = {"markdown-generation-en.png", "official-documents-en.png", "template-management-en.png"}

BANNED_CURRENT_TOKENS = (
    "gui-ui-ux-acceptance",
    "gui-visual-parity-baseline",
    "linux-compatibility-acceptance",
    "三项目证据化审计清单",
    "协作交接",
    "parity_goal_prompt.txt",
    "cli_json_output_schema",
    "docs/i18n/",
    "docs/guides/",
    "docs/architecture/",
)


def test_current_documentation_inventory_is_complete() -> None:
    missing = sorted(path for path in REQUIRED_DOCS if not (DOCS / path).is_file())
    assert not missing
    assert {path.name for path in (DOCS / "user-guides").glob("*.md")} == set(LOCALIZED_READMES) | {"macos.md"}
    assert {path.name for path in (DOCS / "assets" / "screenshots").glob("*.png")} == SCREENSHOTS


def test_v090_changelog_explicitly_declares_removed_python_import_surfaces() -> None:
    text = (DOCS / "CHANGELOG.md").read_text(encoding="utf-8")
    chinese_match = re.search(
        r"### 破坏性变化（中文）\n(?P<section>.*?)(?=\n### Breaking changes \(English\))",
        text,
        re.DOTALL,
    )
    english_match = re.search(
        r"### Breaking changes \(English\)\n(?P<section>.*?)(?=\n## v0\.8\.5)",
        text,
        re.DOTALL,
    )
    assert chinese_match is not None
    assert english_match is not None

    removed_surfaces = (
        "OfficeSoftwareNotFoundError",
        "docwen.cli.api",
        "docwen.converter.docx.to_md.shared.docx_utils",
    )
    migration_contract = (
        "docwen_cli",
        "protocol 3",
        "dependency_missing",
        "conversion_failed",
        "operation_cancelled",
    )
    for section in (chinese_match.group("section"), english_match.group("section")):
        assert all(surface in section for surface in removed_surfaces)
        assert all(token in section for token in migration_contract)


def test_public_readmes_share_one_language_navigation_contract() -> None:
    public_readmes = [ROOT / "README.md", *(DOCS / "user-guides" / name for name in LOCALIZED_READMES)]

    for path in public_readmes:
        text = path.read_text(encoding="utf-8")
        assert text.startswith(f"# DocWen\n\n{LOGO_BLOCK}\n\n{LANGUAGE_NAV}")
        assert text.count(LOGO_BLOCK) == 1
        assert text.count(LANGUAGE_NAV) == 1

    docs_readme = (DOCS / "README.md").read_text(encoding="utf-8")
    assert LOGO_BLOCK in docs_readme
    assert LANGUAGE_NAV in docs_readme


def _public_readme_paths() -> list[Path]:
    return [ROOT / "README.md", *(DOCS / "user-guides" / name for name in LOCALIZED_READMES)]


CLI_SECTION_HEADING_MARKERS = {
    "README.ar.md": "استخدام سطر الأوامر",
    "README.md": "Command Line Usage",
    "README.de-DE.md": "CLI-Verwendung",
    "README.es-ES.md": "Uso de la CLI",
    "README.fr-FR.md": "Utilisation en ligne de commande",
    "README.ja-JP.md": "コマンドライン利用",
    "README.ko-KR.md": "명령줄 사용",
    "README.pt-BR.md": "Uso da CLI",
    "README.ru-RU.md": "Использование CLI",
    "README.vi-VN.md": "Sử dụng CLI",
    "README.zh-CN.md": "命令行使用",
    "README.zh-TW.md": "命令列使用",
}


def test_public_readmes_keep_cli_integration_anchors() -> None:
    required = (
        "inspect <file> [--json]",
        "resources list formats --json",
        "schema convert",
        "--dry-run --json",
        "serve --stdio",
        "Machine Protocol v2",
        "resources list optimizations --json",
        "gui open",
    )
    for path in _public_readme_paths():
        text = path.read_text(encoding="utf-8")
        cli = _named_h2_section(text, CLI_SECTION_HEADING_MARKERS[path.name])
        assert all(token in cli for token in required), path.name
        assert cli.count("| `gui open [ABSOLUTE_FILE]") == 1, path.name
        ordered = [
            cli.index("inspect <file> [--json]"),
            cli.index("resources list formats --json"),
            cli.index("schema convert"),
            cli.index("--dry-run --json"),
        ]
        assert ordered == sorted(ordered), path.name


ASSISTANT_PLATFORM_SUPPORT_MARKERS = {
    "README.ar.md": "يدعم ملحق DocWen Assistant سطح المكتب على Windows وLinux",
    "README.md": "supports Windows and Linux desktop hosts",
    "README.de-DE.md": "unterstützt Windows- und Linux-Desktop-Hosts",
    "README.es-ES.md": "admite hosts de escritorio Windows y Linux",
    "README.fr-FR.md": "prend en charge les postes de bureau Windows et Linux",
    "README.ja-JP.md": "Windows と Linux のデスクトップ環境をサポート",
    "README.ko-KR.md": "Windows와 Linux 데스크톱 호스트를 지원",
    "README.pt-BR.md": "oferece suporte a desktops Windows e Linux",
    "README.ru-RU.md": "поддерживает настольные системы Windows и Linux",
    "README.vi-VN.md": "hỗ trợ máy tính để bàn Windows và Linux",
    "README.zh-CN.md": "支持 Windows 和 Linux 桌面端",
    "README.zh-TW.md": "支援 Windows 與 Linux 桌面端",
}

ASSISTANT_MACOS_UNSUPPORTED_MARKERS = {
    "README.ar.md": "لا يدعم DocWen Assistant نظام macOS حاليًا",
    "README.md": "does not currently support macOS",
    "README.de-DE.md": "macOS wird vom Assistant derzeit nicht unterstützt",
    "README.es-ES.md": "Assistant no admite macOS actualmente",
    "README.fr-FR.md": "Assistant ne prend actuellement pas en charge macOS",
    "README.ja-JP.md": "Assistant は現在 macOS をサポートしていません",
    "README.ko-KR.md": "Assistant는 현재 macOS를 지원하지 않습니다",
    "README.pt-BR.md": "Assistant atualmente não oferece suporte ao macOS",
    "README.ru-RU.md": "macOS в Assistant сейчас не поддерживается",
    "README.vi-VN.md": "Assistant hiện chưa hỗ trợ macOS",
    "README.zh-CN.md": "Assistant 目前不支持 macOS",
    "README.zh-TW.md": "Assistant 目前不支援 macOS",
}

STALE_ASSISTANT_PLATFORM_MARKERS = (
    "remains Windows desktop-only",
    "仍仅限 Windows 桌面端",
    "仍僅限 Windows 桌面端",
    "bleibt auf Windows-Desktop beschränkt",
    "reste limité au bureau Windows",
    "sigue limitado al escritorio de Windows",
    "permanece exclusivo para desktop Windows",
    "остаётся только для Windows desktop",
    "Windows デスクトップ専用のまま",
    "Windows 데스크톱 전용",
    "vẫn chỉ dành cho Windows",
)


def _named_h2_section(text: str, token: str) -> str:
    lines = text.splitlines()
    start = next((index for index, line in enumerate(lines) if line.startswith("## ") and token in line), None)
    assert start is not None
    end = next((index for index in range(start + 1, len(lines)) if lines[index].startswith("## ")), len(lines))
    return "\n".join(lines[start:end])


def test_public_readmes_describe_current_assistant_platform_boundary() -> None:
    for path in _public_readme_paths():
        text = path.read_text(encoding="utf-8")
        obsidian = _named_h2_section(text, "Obsidian")
        support_marker = ASSISTANT_PLATFORM_SUPPORT_MARKERS[path.name]
        assert support_marker in obsidian, path.name
        assert text.count(support_marker) == 1, path.name
        assert ASSISTANT_MACOS_UNSUPPORTED_MARKERS[path.name] in obsidian, path.name
        assert not any(marker in text for marker in STALE_ASSISTANT_PLATFORM_MARKERS), path.name


def test_public_readmes_and_architecture_describe_the_network_guard_boundary() -> None:
    public_readmes = [ROOT / "README.md", *(DOCS / "user-guides" / name for name in LOCALIZED_READMES)]
    assert {path.name for path in public_readmes} == set(NETWORK_GUARD_EXTERNAL_PROCESS_MARKERS)

    # Public guides explain the user-visible scope and link to one technical
    # description that retains the exact operations and implementation limits.
    architecture = (DOCS / "architecture.md").read_text(encoding="utf-8")
    assert "CPython audit hook" in architecture
    assert "all DNS/name resolution" in architecture
    assert "AF_INET/AF_INET6" in architecture
    assert all(f"`{operation}`" in architecture for operation in NETWORK_GUARD_IP_OPERATIONS)
    assert "Windows named pipes" in architecture and "Unix-domain sockets" in architecture
    assert "dedicated Office helper" in architecture
    assert "not an operating-system sandbox" in architecture

    stale_scope_markers = (
        "non-local DNS",
        "nicht lokale DNS",
        "DNS no local",
        "DNS non local",
        "非ローカル DNS",
        "비로컬 DNS",
        "DNS não local",
        "нелокальный DNS",
        "DNS không cục bộ",
        "非本地 DNS",
    )
    assert not any(marker in architecture for marker in stale_scope_markers)
    for path in public_readmes:
        text = path.read_text(encoding="utf-8")
        assert "Python" in text
        external_processes = next(line for line in text.splitlines() if "Word/Office" in line)
        assert "WPS" in external_processes and "LibreOffice" in external_processes
        assert NETWORK_GUARD_EXTERNAL_PROCESS_MARKERS[path.name] in external_processes, path.name
        architecture_link = "docs/architecture.md" if path.parent == ROOT else "../architecture.md"
        assert f"({architecture_link})" in text, path.name
        assert not any(marker in text for marker in stale_scope_markers)


def test_readmes_do_not_reintroduce_stale_localization_guidance() -> None:
    """Guidance must reflect current conversion defaults and UI settings."""
    import tomllib

    config = tomllib.loads((ROOT / "configs" / "conversion.toml").read_text(encoding="utf-8"))
    assert config["md_to_docx"]["each_line_paragraph"] is False

    policy = (DOCS / "specs" / "task-language.md").read_text(encoding="utf-8")
    assert "Follow task output language" in policy

    obsolete = (
        "Every non-empty line is treated as a separate paragraph by default.",
        "每个非空行默认作为独立段落处理",
        "每個非空行默認作為獨立段落處理",
        "Jede nicht leere Zeile wird standardmäßig als separater Absatz behandelt.",
        "Chaque ligne non vide est traitée par défaut comme un paragraphe séparé.",
        "Cada línea no vacía se trata como un párrafo independiente por defecto.",
        "空でない行はデフォルトで別々の段落として扱われます。",
        "비어 있지 않은 각 줄은 기본적으로 독립된 단락으로 처리됩니다.",
        "Cada linha não vazia é tratada como um parágrafo separado por padrão.",
        "Каждая непустая строка по умолчанию обрабатывается как отдельный абзац.",
        "Mỗi dòng không rỗng được xem là một đoạn riêng theo mặc định.",
        "設定按選項卡組織",
        "Die Einstellungen sind in Tabs organisiert",
        "Les paramètres sont organisés en onglets",
        "La configuración está organizada en pestañas",
        "設定はタブで構成されています",
        "설정은 탭으로 구성됩니다",
        "As configurações são organizadas em abas",
        "Настройки организованы по вкладкам",
        "Cài đặt được tổ chức theo tab",
    )
    guides = _public_readme_paths()
    assert len(guides) == 12

    # File-drop button labels come from the shipping locale catalog, not English screenshots.
    file_drop_locales = {
        "README.md": "en_US",
        "README.zh-CN.md": "zh_CN",
        "README.zh-TW.md": "zh_TW",
        "README.de-DE.md": "de_DE",
        "README.fr-FR.md": "fr_FR",
        "README.es-ES.md": "es_ES",
        "README.ja-JP.md": "ja_JP",
        "README.ko-KR.md": "ko_KR",
        "README.pt-BR.md": "pt_BR",
        "README.ru-RU.md": "ru_RU",
        "README.vi-VN.md": "vi_VN",
    }
    for readme in guides:
        locale = file_drop_locales.get(readme.name)
        if locale is None:  # The Arabic overview does not describe the file-drop button.
            continue
        source = tomllib.loads((ROOT / "i18n" / "locales" / f"{locale}.toml").read_text(encoding="utf-8"))
        label = source["components"]["file_drop"]["add_button"]
        guide = readme.read_text(encoding="utf-8")
        assert f'"{label}"' in guide or f"「{label}」" in guide, readme.name

    for path in guides:
        content = path.read_text(encoding="utf-8")
        assert not any(phrase in content for phrase in obsolete), path.name

    arabic = (DOCS / "user-guides" / "README.ar.md").read_text(encoding="utf-8")
    ocr = next(line for line in arabic.splitlines() if line.startswith("**التعرف على النص (OCR):**"))
    assert "لغة إخراج المهمة" in ocr
    assert "اتباع لغة الواجهة" not in ocr

    for locale in ("es-ES", "ko-KR", "vi-VN"):
        guide = (DOCS / "user-guides" / f"README.{locale}.md").read_text(encoding="utf-8")
        assert "title: Test Document" not in guide
        assert "## Test Title" not in guide
        assert "This is the test body content." not in guide
        assert "## I. Work Requirements." not in guide
        assert "This meeting..." not in guide

    vietnamese = (DOCS / "user-guides" / "README.vi-VN.md").read_text(encoding="utf-8")
    assert "Text Proofreading" not in vietnamese
    assert "Export MD" not in vietnamese


def test_public_guides_explain_supported_proofreading_and_release_selection() -> None:
    """Guides must not narrow CLI validation or pin companion release advice."""
    for path in _public_readme_paths():
        guide = path.read_text(encoding="utf-8")
        prose = re.sub(r"```.*?```", "", guide, flags=re.DOTALL)
        validation_paragraphs = [paragraph for paragraph in prose.split("\n\n") if "validate" in paragraph]
        assert any(
            all(extension in paragraph for extension in (".md", ".docx", ".doc", ".wps", ".rtf", ".odt"))
            for paragraph in validation_paragraphs
        ), path.name
        assert "validate" in guide and "--report" in guide, path.name
        assert "https://github.com/ZHYX91/obsidian-docwen-assistant/releases" in _named_h2_section(guide, "Obsidian"), (
            path.name
        )
        openclaw_section = _named_h2_section(guide, "Obsidian" if path.name == "README.ar.md" else "OpenClaw")
        assert "https://github.com/ZHYX91/docwen-openclaw/releases" in openclaw_section, path.name
        assert "Assistant 3.3" not in guide, path.name
        assert "Plugin 3.0" not in guide, path.name
        screenshot = (
            "docs/assets/screenshots/official-documents-en.png"
            if path.parent == ROOT
            else "../assets/screenshots/official-documents-en.png"
        )
        assert screenshot in guide, path.name
        assert "--optimization gongwen" in guide, path.name

    # Arabic is intentionally a compact starting point, not a partial copy
    # presented as an equally detailed full reference.
    arabic = (DOCS / "user-guides" / "README.ar.md").read_text(encoding="utf-8")
    assert "دليل عربي مختصر" in arabic
    assert "(../../README.md)" in arabic


def test_current_docs_do_not_depend_on_completed_audit_material() -> None:
    current_files = [ROOT / "README.md", *DOCS.rglob("*.md")]
    current_files = [path for path in current_files if path.name != "CHANGELOG.md"]

    violations: list[str] = []
    for path in current_files:
        text = path.read_text(encoding="utf-8")
        for token in BANNED_CURRENT_TOKENS:
            if token in text:
                violations.append(f"{path.relative_to(ROOT)}: {token}")
        if re.search(r"\bVIS-\d", text) or re.search(r"\bF-\d", text):
            violations.append(f"{path.relative_to(ROOT)}: stage identifier")
    assert not violations, "\n".join(violations)


def test_current_documentation_local_links_resolve() -> None:
    current_files = [ROOT / "README.md", *DOCS.rglob("*.md")]
    current_files = [path for path in current_files if path.name != "CHANGELOG.md"]
    unresolved: list[str] = []
    link_pattern = re.compile(r"!?\[[^]]*\]\(([^)]+)\)")

    for path in current_files:
        in_fence = False
        for line in path.read_text(encoding="utf-8").splitlines():
            if line.lstrip().startswith("```"):
                in_fence = not in_fence
                continue
            if in_fence:
                continue
            for raw_target in link_pattern.findall(line):
                target = raw_target.strip().strip("<>")
                if target == "url" or target.startswith(("http://", "https://", "mailto:", "#")):
                    continue
                relative = unquote(target.split("#", 1)[0])
                if relative and not (path.parent / relative).resolve().exists():
                    unresolved.append(f"{path.relative_to(ROOT)} -> {target}")

    assert not unresolved, "\n".join(unresolved)

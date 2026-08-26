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
)
LOGO_BLOCK = (
    '<p align="center">\n'
    '  <img src="https://raw.githubusercontent.com/ZHYX91/docwen/main/assets/icon.svg" '
    'alt="DocWen logo" width="120">\n'
    "</p>"
)

NETWORK_GUARD_DNS_MARKERS = {
    "README.md": "all DNS/name resolution",
    "README.de-DE.md": "sämtliche DNS-/Namensauflösung",
    "README.es-ES.md": "toda resolución DNS/de nombres",
    "README.fr-FR.md": "toute résolution DNS/de noms",
    "README.ja-JP.md": "すべての DNS／名前解決",
    "README.ko-KR.md": "모든 DNS/이름 확인",
    "README.pt-BR.md": "toda resolução DNS/de nomes",
    "README.ru-RU.md": "любое разрешение DNS/имён",
    "README.vi-VN.md": "mọi hoạt động phân giải DNS/tên",
    "README.zh-CN.md": "所有 DNS/名称解析",
    "README.zh-TW.md": "所有 DNS／名稱解析",
}

NETWORK_GUARD_IP_OPERATIONS = ("bind", "connect", "connect_ex", "sendto", "sendmsg")

NETWORK_GUARD_HELPER_MARKERS = {
    "README.md": "dedicated Office helper",
    "README.de-DE.md": "Office-Helfers",
    "README.es-ES.md": "ayudante de Office",
    "README.fr-FR.md": "assistant Office",
    "README.ja-JP.md": "専用 Office ヘルパー",
    "README.ko-KR.md": "전용 Office 헬퍼",
    "README.pt-BR.md": "auxiliar do Office",
    "README.ru-RU.md": "служебный процесс Office",
    "README.vi-VN.md": "Office helper chuyên dụng",
    "README.zh-CN.md": "专用 Office helper",
    "README.zh-TW.md": "專用 Office helper",
}

NETWORK_GUARD_SANDBOX_MARKERS = {
    "README.md": "operating-system sandbox",
    "README.de-DE.md": "Betriebssystem-Sandbox",
    "README.es-ES.md": "sandbox del sistema operativo",
    "README.fr-FR.md": "bac à sable du système",
    "README.ja-JP.md": "OS サンドボックス",
    "README.ko-KR.md": "운영체제 샌드박스",
    "README.pt-BR.md": "sandbox do sistema operacional",
    "README.ru-RU.md": "песочница операционной системы",
    "README.vi-VN.md": "sandbox của hệ điều hành",
    "README.zh-CN.md": "操作系统级沙箱",
    "README.zh-TW.md": "作業系統級沙箱",
}

SCREENSHOTS = {
    f"{surface}-{theme}.png"
    for surface in (
        "main",
        "settings",
        "batch",
        "conversion-document",
        "conversion-spreadsheet",
        "conversion-image",
        "conversion-layout",
        "about",
    )
    for theme in ("light", "dark")
}

BANNED_CURRENT_TOKENS = (
    "gui-ui-ux-acceptance",
    "gui-visual-parity-baseline",
    "linux-compatibility-acceptance",
    "三项目证据化审计清单",
    "AI协作交接",
    "parity_goal_prompt.txt",
    "cli_json_output_schema",
    "docs/i18n/",
    "docs/guides/",
    "docs/architecture/",
)


def test_current_documentation_inventory_is_complete() -> None:
    missing = sorted(path for path in REQUIRED_DOCS if not (DOCS / path).is_file())
    assert not missing
    assert {path.name for path in (DOCS / "user-guides").glob("*.md")} == set(LOCALIZED_READMES)
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


def test_public_readmes_describe_the_exact_network_guard_boundary() -> None:
    public_readmes = [ROOT / "README.md", *(DOCS / "user-guides" / name for name in LOCALIZED_READMES)]
    assert {path.name for path in public_readmes} == set(NETWORK_GUARD_DNS_MARKERS)
    assert set(NETWORK_GUARD_SANDBOX_MARKERS) == set(NETWORK_GUARD_DNS_MARKERS)
    assert set(NETWORK_GUARD_HELPER_MARKERS) == set(NETWORK_GUARD_DNS_MARKERS)

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
    for path in public_readmes:
        text = path.read_text(encoding="utf-8")
        assert NETWORK_GUARD_DNS_MARKERS[path.name] in text
        assert "AF_INET/AF_INET6" in text
        assert all(f"`{operation}`" in text for operation in NETWORK_GUARD_IP_OPERATIONS)
        assert "Windows" in text and "Unix" in text
        assert "Office" in text and "WPS" in text and "LibreOffice" in text
        assert NETWORK_GUARD_HELPER_MARKERS[path.name] in text
        assert NETWORK_GUARD_SANDBOX_MARKERS[path.name] in text
        assert not any(marker in text for marker in stale_scope_markers)


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


def test_capability_inventory_has_one_row_per_current_feature() -> None:
    text = (DOCS / "capabilities.md").read_text(encoding="utf-8")
    rows = [line for line in text.splitlines() if line.startswith("| FEAT-")]
    assert len(rows) == 160
    assert len({line.split("|", 2)[1].strip() for line in rows}) == 160


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

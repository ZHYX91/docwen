from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class MissingPath:
    ref: str
    resolved: str


def _repo_root() -> Path:
    return Path(__file__).resolve().parents[2]


def _load_doc_text(root: Path) -> str:
    return (root / "doc" / "技术文档.md").read_text(encoding="utf-8")


def _iter_backticked(text: str) -> list[str]:
    return [m.group(1).strip() for m in re.finditer(r"`([^`\n]+)`", text)]


def _resolve_repo_path(root: Path, token: str) -> Path | None:
    if token.startswith("http://") or token.startswith("https://"):
        return None

    if "\\" in token:
        try:
            p = Path(token)
            if p.is_absolute():
                return p
        except Exception:
            return None

    rp = token
    if (
        rp.startswith("converter/")
        or rp.startswith("services/")
        or rp.startswith("utils/")
        or rp.startswith("gui/")
        or rp.startswith("cli/")
        or rp.startswith("config/")
        or rp.startswith("ipc/")
        or rp.startswith("security/")
        or rp.startswith("template/")
        or rp.startswith("i18n/")
        or rp.startswith("docx_spell/")
        or rp.startswith("table_merger/")
    ):
        rp = "src/docwen/" + rp
    elif rp.startswith(
        (
            "configs/",
            "templates/",
            "tests/",
            "assets/",
            "samples/",
            "scripts/",
            "doc/",
            "src/",
        )
    ):
        rp = rp
    else:
        if rp.endswith(".py") and "/" not in rp and "\\" not in rp:
            rp = "src/docwen/" + rp
        else:
            return None

    return root / rp


def _extract_symbol_names(text: str) -> set[str]:
    names: set[str] = set()

    in_py = False
    for line in text.splitlines():
        if line.strip().startswith("```"):
            tag = line.strip()[3:].strip()
            if not in_py and tag == "python":
                in_py = True
                continue
            if in_py:
                in_py = False
                continue

        if not in_py:
            continue

        m = re.match(r"\s*(def|class)\s+([A-Za-z_][A-Za-z0-9_]*)\s*\(", line)
        if m:
            names.add(m.group(2))

    for m in re.finditer(r"`([A-Za-z_][A-Za-z0-9_]*)\(\)`", text):
        names.add(m.group(1))

    return names


def _index_py_symbols(root: Path) -> set[str]:
    found: set[str] = set()
    pat = re.compile(r"^\s*(def|class)\s+([A-Za-z_][A-Za-z0-9_]*)\b")
    for fp in root.rglob("*.py"):
        try:
            txt = fp.read_text(encoding="utf-8", errors="ignore")
        except Exception:
            continue
        for ln in txt.splitlines():
            m = pat.match(ln)
            if m:
                found.add(m.group(2))
    return found


def main() -> int:
    root = _repo_root()
    text = _load_doc_text(root)

    tokens = _iter_backticked(text)
    missing_paths: list[MissingPath] = []
    checked_paths = 0
    for t in sorted(set(tokens)):
        p = _resolve_repo_path(root, t)
        if p is None:
            continue
        checked_paths += 1
        if not p.exists():
            missing_paths.append(MissingPath(ref=t, resolved=str(p)))

    names = _extract_symbol_names(text)
    index = _index_py_symbols(root)
    missing_symbols = sorted(n for n in names if n not in index)

    print(f"PATH_CHECK_TOTAL {checked_paths}")
    print(f"PATH_MISSING_TOTAL {len(missing_paths)}")
    for item in missing_paths:
        print(f"MISSING_PATH {item.ref} -> {item.resolved}")

    print(f"SYMBOL_CHECK_TOTAL {len(names)}")
    print(f"SYMBOL_MISSING_TOTAL {len(missing_symbols)}")
    for n in missing_symbols:
        print(f"MISSING_SYMBOL {n}")

    return 0 if not missing_paths and not missing_symbols else 2


if __name__ == "__main__":
    raise SystemExit(main())

from __future__ import annotations

import argparse
import ast
import re
from collections.abc import Iterable, Mapping
from pathlib import Path
from typing import Any

import tomlkit

HEADER_RE = re.compile(r"^\[(?P<name>[A-Za-z0-9_.-]+)\]\s*$")
ASSIGN_RE = re.compile(r"^(?P<key>[A-Za-z0-9_]+)\s*=")


def get_locales_dir() -> Path:
    return Path(__file__).resolve().parent / "locales"


def get_reference_locale_path() -> Path:
    return get_locales_dir() / "zh_CN.toml"


def get_all_locale_paths() -> list[Path]:
    return sorted(get_locales_dir().glob("*.toml"))


def _read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def _write_text(path: Path, text: str) -> None:
    path.write_text(text, encoding="utf-8", newline="\n")


def _load_toml(path: Path) -> Any:
    with path.open("r", encoding="utf-8") as f:
        return tomlkit.load(f)


def _save_toml(path: Path, doc: Any) -> None:
    with path.open("w", encoding="utf-8", newline="\n") as f:
        tomlkit.dump(doc, f)


def _is_mapping(value: Any) -> bool:
    return isinstance(value, Mapping)


def extract_top_level_section_order(toml_text: str) -> list[str]:
    order: list[str] = []
    for line in toml_text.splitlines():
        m = HEADER_RE.match(line.strip())
        if not m:
            continue
        top = m.group("name").split(".", 1)[0]
        if top not in order:
            order.append(top)
    return order


def extract_styles_key_order(toml_text: str) -> list[str]:
    lines = toml_text.splitlines()
    start = None
    for i, line in enumerate(lines):
        if line.strip() == "[styles]":
            start = i + 1
            break
    if start is None:
        return []
    keys: list[str] = []
    for line in lines[start:]:
        if HEADER_RE.match(line.strip()):
            break
        m = ASSIGN_RE.match(line)
        if m:
            keys.append(m.group("key"))
    return keys


def _flatten_leaf_keys(obj: Any, prefix: str = "") -> set[str]:
    keys: set[str] = set()
    if not _is_mapping(obj):
        return keys
    for k, v in obj.items():
        if not isinstance(k, str):
            continue
        full = f"{prefix}.{k}" if prefix else k
        if _is_mapping(v):
            keys |= _flatten_leaf_keys(v, full)
        else:
            keys.add(full)
    return keys


def extract_all_leaf_keys(toml_path: Path) -> set[str]:
    return _flatten_leaf_keys(_load_toml(toml_path))


def extract_translation_keys_for_unused_check(toml_path: Path) -> set[str]:
    doc = _load_toml(toml_path)
    keys: set[str] = set()

    skip_top_level = {"meta", "style_formats"}
    skip_suffixes = {"_locales"}

    def walk(obj: Any, prefix: str = "") -> None:
        if not _is_mapping(obj):
            return
        for k, v in obj.items():
            if not isinstance(k, str):
                continue
            if not prefix and k in skip_top_level:
                continue
            if any(k.endswith(s) for s in skip_suffixes):
                continue
            full = f"{prefix}.{k}" if prefix else k
            if _is_mapping(v):
                walk(v, full)
            else:
                keys.add(full)

    walk(doc)
    return keys


def find_project_root(start: Path | None = None) -> Path:
    base = (start or Path.cwd()).resolve()
    for candidate in [base, *base.parents]:
        if (candidate / "pyproject.toml").exists() and (candidate / "src" / "docwen").exists():
            return candidate
    return base


def scan_python_files_for_used_keys(src_dirs: Iterable[Path]) -> set[str]:
    used: set[str] = set()
    patterns = [
        re.compile(r"\bt\s*\(\s*[\"']([a-zA-Z_][a-zA-Z0-9_\.]*)[\"']"),
        re.compile(r"\bcli_t\s*\(\s*[\"']([a-zA-Z_][a-zA-Z0-9_\.]*)[\"']"),
        re.compile(r"\bt_all_locales\s*\(\s*[\"']([a-zA-Z_][a-zA-Z0-9_\.]*)[\"']"),
        re.compile(r"\bt_locale\s*\(\s*[\"']([a-zA-Z_][a-zA-Z0-9_\.]*)[\"']"),
        re.compile(r"\.t\s*\(\s*[\"']([a-zA-Z_][a-zA-Z0-9_\.]*)[\"']"),
    ]

    var_assign_patterns = [
        re.compile(r"\b([A-Za-z_][A-Za-z0-9_]*)\s*=\s*[\"']([a-zA-Z_][a-zA-Z0-9_\.]*)[\"']"),
        re.compile(
            r"\b([A-Za-z_][A-Za-z0-9_]*)\s*=\s*[\"']([a-zA-Z_][a-zA-Z0-9_\.]*)[\"']\s*if\s+.+?\s+else\s+[\"']([a-zA-Z_][a-zA-Z0-9_\.]*)[\"']"
        ),
    ]

    var_call_pattern = re.compile(r"\b(?:t|cli_t|t_all_locales|t_locale)\s*\(\s*([A-Za-z_][A-Za-z0-9_]*)\s*(?:,|\))")

    for src_dir in src_dirs:
        for py_file in src_dir.rglob("*.py"):
            try:
                content = py_file.read_text(encoding="utf-8")
            except Exception:
                try:
                    content = py_file.read_text(encoding="utf-8-sig")
                except Exception:
                    continue

            var_key_map: dict[str, set[str]] = {}
            for pat in var_assign_patterns:
                for match in pat.findall(content):
                    if isinstance(match, tuple):
                        if len(match) == 2:
                            var_name, key = match
                            var_key_map.setdefault(var_name, set()).add(key)
                        elif len(match) == 3:
                            var_name, key_a, key_b = match
                            var_key_map.setdefault(var_name, set()).update([key_a, key_b])

            for pat in patterns:
                used.update(pat.findall(content))

            for var_name in var_call_pattern.findall(content):
                if var_name in var_key_map:
                    used.update(var_key_map[var_name])

    return used


def scan_translation_key_maps(src_dirs: Iterable[Path]) -> set[str]:
    used: set[str] = set()
    key_re = re.compile(r"^[a-zA-Z_][a-zA-Z0-9_\.]*$")

    def is_key(value: str) -> bool:
        return bool(value) and bool(key_re.match(value))

    def extract_str_list(expr: ast.AST, list_map: dict[str, list[str]]) -> list[str] | None:
        if isinstance(expr, ast.Name):
            return list_map.get(expr.id)
        if not isinstance(expr, (ast.List, ast.Tuple)):
            return None
        items: list[str] = []
        for e in expr.elts:
            if isinstance(e, ast.Constant) and isinstance(e.value, str) and e.value:
                items.append(e.value)
        return items or None

    def extract_fstring_pattern(expr: ast.AST, var_name: str) -> tuple[str, str] | None:
        if not isinstance(expr, ast.JoinedStr):
            return None
        prefix_parts: list[str] = []
        suffix_parts: list[str] = []
        seen_var = False
        for part in expr.values:
            if isinstance(part, ast.Constant) and isinstance(part.value, str):
                (suffix_parts if seen_var else prefix_parts).append(part.value)
                continue
            if isinstance(part, ast.FormattedValue) and isinstance(part.value, ast.Name) and part.value.id == var_name:
                if seen_var:
                    return None
                seen_var = True
                continue
            return None
        if not seen_var:
            return None
        prefix = "".join(prefix_parts)
        suffix = "".join(suffix_parts)
        return (prefix, suffix) if prefix else None

    for src_dir in src_dirs:
        for py_file in src_dir.rglob("*.py"):
            try:
                tree = ast.parse(py_file.read_text(encoding="utf-8-sig"))
            except Exception:
                continue

            list_map: dict[str, list[str]] = {}
            for node in tree.body:
                if not isinstance(node, ast.Assign) or len(node.targets) != 1:
                    continue
                target = node.targets[0]
                if not isinstance(target, ast.Name):
                    continue
                items = extract_str_list(node.value, list_map)
                if items:
                    list_map[target.id] = items

            for node in tree.body:
                if not isinstance(node, ast.Assign) or len(node.targets) != 1:
                    continue
                target = node.targets[0]
                if not isinstance(target, ast.Name):
                    continue
                name = target.id
                if "TRANSLATE_KEY" not in name and "TRANSLATE_KEYS" not in name:
                    continue

                value = node.value
                if isinstance(value, ast.Dict):
                    for v in value.values:
                        if isinstance(v, ast.Constant) and isinstance(v.value, str) and is_key(v.value):
                            used.add(v.value)
                    continue

                if isinstance(value, ast.DictComp) and len(value.generators) == 1:
                    gen = value.generators[0]
                    if not isinstance(gen.target, ast.Name):
                        continue
                    var_name = gen.target.id
                    items = extract_str_list(gen.iter, list_map)
                    if not items:
                        continue
                    pat = extract_fstring_pattern(value.value, var_name)
                    if not pat:
                        continue
                    prefix, suffix = pat
                    for item in items:
                        k = f"{prefix}{item}{suffix}"
                        if is_key(k):
                            used.add(k)

    return used


def scan_config_files_for_i18n_keys(config_dir: Path) -> set[str]:
    def load_toml(file_path: Path) -> Any:
        if not file_path.exists():
            return {}
        with file_path.open("r", encoding="utf-8") as f:
            data = tomlkit.load(f)
        return data if _is_mapping(data) else {}

    used: set[str] = set()

    add_cfg = load_toml(config_dir / "heading_numbering_add.toml")
    schemes = add_cfg.get("schemes", {}) if _is_mapping(add_cfg) else {}
    if _is_mapping(schemes):
        for _, scheme in schemes.items():
            if not _is_mapping(scheme):
                continue
            name_key = scheme.get("name_key")
            if isinstance(name_key, str) and name_key:
                used.add(f"editors.numbering_add.names.{name_key}")
            desc_key = scheme.get("description_key")
            if isinstance(desc_key, str) and desc_key:
                used.add(f"editors.numbering_add.descriptions.{desc_key}")

    clean_cfg = load_toml(config_dir / "heading_numbering_clean.toml")
    rules = clean_cfg.get("rules", {}) if _is_mapping(clean_cfg) else {}
    if _is_mapping(rules):
        for _, rule in rules.items():
            if not _is_mapping(rule):
                continue
            name_key = rule.get("name_key")
            if isinstance(name_key, str) and name_key:
                used.add(f"editors.numbering_clean.names.{name_key}")
            desc_key = rule.get("description_key")
            if isinstance(desc_key, str) and desc_key:
                used.add(f"editors.numbering_clean.descriptions.{desc_key}")

    return used


def scan_cli_dynamic_keys(src_root: Path) -> set[str]:
    cli_utils_path = src_root / "cli" / "utils.py"
    if not cli_utils_path.exists():
        return set()
    try:
        tree = ast.parse(cli_utils_path.read_text(encoding="utf-8-sig"))
    except Exception:
        return set()

    categories: set[str] = set()
    for node in tree.body:
        if isinstance(node, ast.FunctionDef) and node.name == "detect_category":
            for sub in ast.walk(node):
                if isinstance(sub, ast.Return) and isinstance(sub.value, ast.Constant):
                    value = sub.value.value
                    if isinstance(value, str) and value:
                        categories.add(value)
    categories.discard("unknown")
    return {f"cli.categories.{c}" for c in categories}


def scan_styles_dynamic_keys(src_root: Path) -> set[str]:
    def parse_file(file_path: Path) -> ast.AST | None:
        if not file_path.exists():
            return None
        try:
            return ast.parse(file_path.read_text(encoding="utf-8-sig"))
        except Exception:
            return None

    def collect_style_keys_from_calls(tree: ast.AST) -> tuple[set[str], list[tuple[str, str, str]]]:
        literal: set[str] = set()
        patterns: list[tuple[str, str, str]] = []
        for node in ast.walk(tree):
            if not isinstance(node, ast.Call) or not node.args:
                continue
            func = node.func
            is_injection_call = False
            if (isinstance(func, ast.Name) and func.id == "get_injection_name") or (
                isinstance(func, ast.Attribute) and func.attr == "get_injection_name"
            ):
                is_injection_call = True
            if not is_injection_call:
                continue

            arg0 = node.args[0]
            if isinstance(arg0, ast.Constant) and isinstance(arg0.value, str):
                if arg0.value:
                    literal.add(arg0.value)
                continue

            if isinstance(arg0, ast.JoinedStr):
                parts = arg0.values
                prefix = (
                    parts[0].value
                    if parts and isinstance(parts[0], ast.Constant) and isinstance(parts[0].value, str)
                    else ""
                )
                var = (
                    parts[1].value.id
                    if len(parts) >= 2
                    and isinstance(parts[1], ast.FormattedValue)
                    and isinstance(parts[1].value, ast.Name)
                    else ""
                )
                suffix = (
                    parts[2].value
                    if len(parts) >= 3 and isinstance(parts[2], ast.Constant) and isinstance(parts[2].value, str)
                    else ""
                )
                if prefix and var and suffix == "":
                    patterns.append((prefix, var, suffix))
        return literal, patterns

    def collect_style_keys_from_mapping(tree: ast.AST, var_name: str) -> set[str]:
        keys: set[str] = set()
        if not isinstance(tree, ast.Module):
            return keys
        for node in tree.body:
            if not isinstance(node, ast.ClassDef) or node.name != "StyleNameResolver":
                continue
            for stmt in node.body:
                if not isinstance(stmt, ast.Assign) or len(stmt.targets) != 1:
                    continue
                target = stmt.targets[0]
                if not isinstance(target, ast.Name) or target.id != var_name:
                    continue
                if isinstance(stmt.value, ast.Dict):
                    for k in stmt.value.keys:
                        if isinstance(k, ast.Constant) and isinstance(k.value, str) and k.value:
                            keys.add(k.value)
        return keys

    def expand_for_range_patterns(tree: ast.AST, patterns: list[tuple[str, str, str]]) -> set[str]:
        expanded: set[str] = set()
        for node in ast.walk(tree):
            if not isinstance(node, ast.For):
                continue
            if not isinstance(node.target, ast.Name):
                continue
            var = node.target.id
            if (
                not isinstance(node.iter, ast.Call)
                or not isinstance(node.iter.func, ast.Name)
                or node.iter.func.id != "range"
            ):
                continue

            args = node.iter.args
            start = 0
            end = None
            if len(args) == 1 and isinstance(args[0], ast.Constant) and isinstance(args[0].value, int):
                end = args[0].value
            elif (
                len(args) >= 2
                and isinstance(args[0], ast.Constant)
                and isinstance(args[0].value, int)
                and isinstance(args[1], ast.Constant)
                and isinstance(args[1].value, int)
            ):
                start = args[0].value
                end = args[1].value
            if end is None:
                continue

            for prefix, pat_var, _ in patterns:
                if pat_var != var:
                    continue
                if prefix not in {"quote_", "horizontal_rule_"}:
                    continue
                for i in range(start, end):
                    if prefix == "horizontal_rule_" and i not in {1, 2, 3}:
                        continue
                    expanded.add(f"{prefix}{i}")
        return expanded

    style_keys: set[str] = set()
    style_resolver_path = src_root / "i18n" / "style_resolver.py"
    templates_path = src_root / "converter" / "md2docx" / "style" / "templates.py"
    injector_path = src_root / "converter" / "md2docx" / "style" / "injector.py"

    for file_path in [style_resolver_path, templates_path, injector_path]:
        tree = parse_file(file_path)
        if tree is None:
            continue

        literal, call_patterns = collect_style_keys_from_calls(tree)
        style_keys |= literal
        style_keys |= expand_for_range_patterns(tree, call_patterns)

        if file_path == style_resolver_path:
            style_keys |= collect_style_keys_from_mapping(tree, "STYLE_CONFIG_MAPPING")
            style_keys |= collect_style_keys_from_mapping(tree, "STYLE_ALIAS_KEY_MAPPING")

    style_keys.discard("")
    return {f"styles.{k}" for k in style_keys}


def scan_fstring_translation_prefixes(src_dirs: Iterable[Path]) -> set[str]:
    """扫描翻译函数中的 f-string 调用，提取动态键前缀。

    检测如 ``t_all_locales(f"placeholders.{var}")`` 这类模式，
    提取前缀 ``"placeholders."``，然后将参考语言文件中所有匹配该前缀的键标记为已使用。
    这避免了 check-unused 因无法解析 f-string 动态变量而误删键。
    """
    translation_funcs = {"t", "cli_t", "t_all_locales", "t_locale"}
    prefixes: set[str] = set()

    for src_dir in src_dirs:
        for py_file in src_dir.rglob("*.py"):
            try:
                tree = ast.parse(py_file.read_text(encoding="utf-8-sig"))
            except Exception:
                continue

            for node in ast.walk(tree):
                if not isinstance(node, ast.Call) or not node.args:
                    continue
                func = node.func
                func_name = None
                if isinstance(func, ast.Name) and func.id in translation_funcs:
                    func_name = func.id
                elif isinstance(func, ast.Attribute) and func.attr in translation_funcs:
                    func_name = func.attr
                if not func_name:
                    continue

                arg0 = node.args[0]
                if not isinstance(arg0, ast.JoinedStr):
                    continue

                # 提取 f-string 中的纯文本前缀部分（在第一个 {} 之前）
                prefix_parts: list[str] = []
                for part in arg0.values:
                    if isinstance(part, ast.Constant) and isinstance(part.value, str):
                        prefix_parts.append(part.value)
                    else:
                        break
                prefix = "".join(prefix_parts)
                if prefix and prefix.endswith("."):
                    prefixes.add(prefix)

    if not prefixes:
        return set()

    # 从参考语言文件中找出所有匹配前缀的键
    ref_keys = extract_translation_keys_for_unused_check(get_reference_locale_path())
    used: set[str] = set()
    for key in ref_keys:
        for prefix in prefixes:
            if key.startswith(prefix):
                used.add(key)
    return used


def scan_used_keys_in_code(project_root: Path) -> set[str]:
    """扫描代码中实际使用的全部 i18n 键。

    聚合多个扫描策略的结果：正则匹配、AST 分析、配置文件引用、
    CLI 动态键、样式动态键，以及 f-string 动态前缀。
    """
    src_root = project_root / "src" / "docwen"
    config_dir = project_root / "configs"
    src_dirs = [src_root]

    used = scan_python_files_for_used_keys(src_dirs)
    used |= scan_translation_key_maps(src_dirs)
    used |= scan_config_files_for_i18n_keys(config_dir)
    used |= scan_cli_dynamic_keys(src_root)
    used |= scan_styles_dynamic_keys(src_root)
    used |= scan_fstring_translation_prefixes(src_dirs)
    return used


def find_unused_keys(project_root: Path | None = None) -> set[str]:
    root = project_root or find_project_root()
    defined = extract_translation_keys_for_unused_check(get_reference_locale_path())
    used = scan_used_keys_in_code(root)
    return defined - used


def remove_key_from_doc(doc: Any, key_path: str) -> bool:
    parts = key_path.split(".")
    current = doc
    for part in parts[:-1]:
        if not _is_mapping(current) or part not in current:
            return False
        current = current[part]
    last = parts[-1]
    if _is_mapping(current) and last in current:
        del current[last]
        return True
    return False


def remove_keys_from_all_locales(keys: Iterable[str]) -> dict[str, int]:
    keys_list = list(keys)
    results: dict[str, int] = {}
    for path in get_all_locale_paths():
        doc = _load_toml(path)
        removed = 0
        for k in keys_list:
            if remove_key_from_doc(doc, k):
                removed += 1
        if removed:
            _save_toml(path, doc)
        results[path.name] = removed
    return results


def _group_by_top_level(text: str) -> tuple[list[str], dict[str, list[str]]]:
    lines = text.splitlines(keepends=True)
    preamble: list[str] = []
    groups: dict[str, list[str]] = {}

    current_top: str | None = None
    current_lines: list[str] = []

    for line in lines:
        m = HEADER_RE.match(line.strip())
        if not m:
            current_lines.append(line)
            continue

        top = m.group("name").split(".", 1)[0]
        if current_top is None:
            preamble.extend(current_lines)
            current_lines = [line]
            current_top = top
            continue

        if top == current_top:
            current_lines.append(line)
            continue

        trail: list[str] = []
        i = len(current_lines) - 1
        while i >= 0:
            s = current_lines[i].strip()
            if s == "" or s.startswith("#"):
                trail.insert(0, current_lines[i])
                i -= 1
                continue
            break
        if trail:
            current_lines = current_lines[: i + 1]

        groups.setdefault(current_top, []).extend(current_lines)
        current_top = top
        current_lines = [*trail, line]

    if current_top is None:
        preamble.extend(current_lines)
    else:
        groups.setdefault(current_top, []).extend(current_lines)

    return preamble, groups


def _build_reordered_text(preamble: list[str], groups: dict[str, list[str]], order: list[str]) -> str:
    out: list[str] = []
    out.extend(preamble)

    for top in order:
        if top in groups:
            out.extend(groups[top])

    for top, group in groups.items():
        if top not in order:
            out.extend(group)

    text = "".join(out)
    if not text.endswith("\n"):
        text += "\n"
    return text


def _reorder_styles_block(toml_text: str, ref_styles_order: list[str]) -> str:
    if not ref_styles_order:
        return toml_text
    lines = toml_text.splitlines(keepends=True)
    start = None
    for i, line in enumerate(lines):
        if line.strip() == "[styles]":
            start = i
            break
    if start is None:
        return toml_text

    end = start + 1
    while end < len(lines):
        if HEADER_RE.match(lines[end].strip()):
            break
        end += 1

    block = lines[start:end]
    head = [block[0]]
    body = block[1:]

    assigns: dict[str, str] = {}
    other: list[str] = []
    for line in body:
        m = ASSIGN_RE.match(line)
        if m:
            assigns[m.group("key")] = line
        else:
            other.append(line)

    if not assigns:
        return toml_text

    ordered: list[str] = []
    seen: set[str] = set()
    for k in ref_styles_order:
        if k in assigns:
            ordered.append(assigns[k])
            seen.add(k)
    for k, line in assigns.items():
        if k not in seen:
            ordered.append(line)

    new_block = head + other + ordered
    return "".join(lines[:start] + new_block + lines[end:])


def sync_order_to_reference(target_path: Path) -> None:
    ref_text = _read_text(get_reference_locale_path())
    ref_top_order = extract_top_level_section_order(ref_text)
    ref_styles_order = extract_styles_key_order(ref_text)

    original = _read_text(target_path)
    preamble, groups = _group_by_top_level(original)
    reordered = _build_reordered_text(preamble, groups, ref_top_order)
    reordered = _reorder_styles_block(reordered, ref_styles_order)
    if reordered != original:
        _write_text(target_path, reordered)


def find_missing_and_extra_keys(target_path: Path) -> tuple[set[str], set[str]]:
    ref = extract_all_leaf_keys(get_reference_locale_path())
    target = extract_all_leaf_keys(target_path)
    return ref - target, target - ref


def _cmd_check_unused(args: argparse.Namespace) -> int:
    root = find_project_root(Path(args.project_root) if args.project_root else None)
    unused = sorted(find_unused_keys(root))
    for k in unused:
        print(k)
    return 0 if not unused else 1


def _cmd_remove_unused(args: argparse.Namespace) -> int:
    root = find_project_root(Path(args.project_root) if args.project_root else None)
    unused = sorted(find_unused_keys(root))
    if not unused:
        print("no unused keys")
        return 0

    if not args.yes:
        print(f"will remove {len(unused)} keys from all locales")
        confirm = input("type 'yes' to continue: ").strip().lower()
        if confirm != "yes":
            print("cancelled")
            return 2

    results = remove_keys_from_all_locales(unused)
    for name, removed in sorted(results.items()):
        print(f"{name}\t{removed}")
    return 0


def _cmd_check_completeness(args: argparse.Namespace) -> int:
    ref = get_reference_locale_path()
    failures = 0
    for path in get_all_locale_paths():
        if path == ref:
            continue
        missing, extra = find_missing_and_extra_keys(path)
        if missing or extra:
            failures += 1
        print(path.name)
        if missing:
            print(f"  missing: {len(missing)}")
            for k in sorted(missing)[:50]:
                print(f"    - {k}")
        if extra:
            print(f"  extra: {len(extra)}")
            for k in sorted(extra)[:50]:
                print(f"    + {k}")
    return 0 if failures == 0 else 1


def _cmd_sync_order(args: argparse.Namespace) -> int:
    ref = get_reference_locale_path()
    targets = []
    targets = [get_locales_dir() / args.file] if args.file else [p for p in get_all_locale_paths() if p != ref]
    for path in targets:
        sync_order_to_reference(path)
        print(path.name)
    return 0


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="python -m docwen.i18n.locale_validator")
    parser.add_argument("--project-root", default=None)
    sub = parser.add_subparsers(dest="cmd", required=True)

    p1 = sub.add_parser("check-unused")
    p1.set_defaults(func=_cmd_check_unused)

    p2 = sub.add_parser("remove-unused")
    p2.add_argument("--yes", action="store_true")
    p2.set_defaults(func=_cmd_remove_unused)

    p3 = sub.add_parser("check-completeness")
    p3.set_defaults(func=_cmd_check_completeness)

    p4 = sub.add_parser("sync-order")
    p4.add_argument("--file", default=None)
    p4.set_defaults(func=_cmd_sync_order)

    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_arg_parser()
    args = parser.parse_args(argv)
    return int(args.func(args))


if __name__ == "__main__":
    raise SystemExit(main())

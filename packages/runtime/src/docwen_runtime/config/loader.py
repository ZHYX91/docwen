"""Configuration loader — three-layer (base / user / runtime) model.

Loads configuration from:
    base/  ── shipped, read-only defaults in ResourceRegistry.default().configs_dir()
    user/  ── sparse overrides in platformdirs.user_config_dir("docwen")/configs
    runtime ── in-memory overrides passed to the constructor (never persisted)

Each file is declared in the ConfigFileSpec registry
(packages/runtime/src/docwen_runtime/config/registry.py).
``reload()`` reads base + user per spec, deep-merges per-file, and wraps
each file under its declared namespace. User files are never backfilled or
migrated. Invalid user files are quarantined; invalid base/runtime values
fail closed. Base files must exist and user files are optional.
"""

from __future__ import annotations

import logging
import os
import threading
from collections.abc import Mapping, MutableMapping
from copy import deepcopy
from datetime import datetime
from pathlib import Path
from tomllib import TOMLDecodeError
from typing import Any

import docwen_runtime.config.transaction as config_transaction
from docwen_runtime.config.registry import (
    CONFIG_FILES,
    relative_key_for_spec,
    require_spec,
    reset_plan_for_group,
    spec_for_dotted_key,
    wrap_namespace,
)
from docwen_runtime.config.validation import ConfigSemanticError, validate_config_file
from docwen_runtime.toml_io import (
    atomic_write_bytes,
    atomic_write_text,
    durable_unlink,
    read_toml_file,
    write_toml_file,
)

logger = logging.getLogger(__name__)
_CONFIG_TRANSACTION_LOCK = threading.RLock()
_CONFIG_TRANSACTION_STATE = threading.local()


def _default_user_config_dir() -> Path:
    """Return the user override config directory.

    Release verification and isolated host tests may provide
    ``DOCWEN_CONFIG_DIR``. This is an internal isolation hook, not a public
    configuration surface.

    Uses ``platformdirs`` to pick a platform-appropriate location, with a
    CWD fallback if ``platformdirs`` is unavailable.
    """
    isolated_dir = os.environ.get("DOCWEN_CONFIG_DIR", "").strip()
    if isolated_dir:
        return Path(isolated_dir)
    try:
        from platformdirs import user_config_dir

        return Path(user_config_dir("docwen", appauthor=False)) / "configs"
    except ImportError:
        return Path.cwd() / "configs"


# ---------------------------------------------------------------------------
# Files excluded from reset (user-curated data)
# ---------------------------------------------------------------------------
RESET_EXCLUDED: frozenset[str] = frozenset(
    {
        "proofread/symbol_map.toml",
        "proofread/typos.toml",
        "proofread/sensitive_words.toml",
    }
)


# ===================================================================
#  Helper: deep merge
# ===================================================================


def deep_merge(default: dict[str, Any], user: dict[str, Any]) -> dict[str, Any]:
    """Recursively merge *user* into *default*, returning a new dict."""
    result = default.copy()
    for key, user_value in user.items():
        if key not in result:
            result[key] = user_value
        elif isinstance(result[key], dict) and isinstance(user_value, dict):
            result[key] = deep_merge(result[key], user_value)
        else:
            result[key] = user_value
    return result


def _reject_reserved_user_keys(user: Mapping[str, Any]) -> None:
    """Reject removed internal metadata instead of giving it configuration meaning."""
    reserved = sorted(str(key) for key in user if str(key).startswith("__docwen_"))
    if reserved:
        joined = ", ".join(reserved)
        raise ValueError(f"Reserved internal keys are not valid user configuration: {joined}")


def _merge_keyed_records(base: Any, user: Any) -> list[Any]:
    """Merge ordered records by stable ``id`` while retaining shipped items."""
    if not isinstance(base, (list, tuple)) or not isinstance(user, (list, tuple)):
        source = user if isinstance(user, (list, tuple)) else base
        return list(deepcopy(source)) if isinstance(source, (list, tuple)) else []

    user_by_id: dict[str, Any] = {}
    for item in user:
        if isinstance(item, Mapping) and item.get("id") is not None:
            user_by_id[str(item["id"])] = item

    merged: list[Any] = []
    consumed: set[str] = set()
    for item in base:
        if not isinstance(item, Mapping) or item.get("id") is None:
            merged.append(deepcopy(item))
            continue
        item_id = str(item["id"])
        user_item = user_by_id.get(item_id)
        if isinstance(user_item, Mapping):
            merged.append(deep_merge(deepcopy(dict(item)), deepcopy(dict(user_item))))
            consumed.add(item_id)
        else:
            merged.append(deepcopy(item))

    for item in user:
        if isinstance(item, Mapping) and item.get("id") is not None:
            item_id = str(item["id"])
            if item_id in consumed or any(
                isinstance(base_item, Mapping) and str(base_item.get("id")) == item_id for base_item in base
            ):
                continue
        merged.append(deepcopy(item))
    return merged


def _merge_file_layers(spec: Any, base: dict[str, Any], user: dict[str, Any]) -> dict[str, Any]:
    """Merge one file according to its explicit registry ownership contract."""
    _reject_reserved_user_keys(user)
    result = deep_merge(deepcopy(base), deepcopy(user))
    for section in spec.keyed_list_sections:
        if section in spec.replace_sections or section not in user:
            continue
        result[section] = _merge_keyed_records(
            base.get(section, []),
            user.get(section, []),
        )
    for section in spec.replace_sections:
        if section in user:
            result[section] = deepcopy(user[section])
    return result


def _merge_toml_documents(
    base_doc: Any,
    user_doc: Any,
    replace_sections: frozenset[str],
    keyed_list_sections: frozenset[str] = frozenset(),
) -> Any:
    """Overlay TOML documents while retaining keyed-item comment trivia."""
    merged = deepcopy(base_doc)

    def overlay(target: Any, source: Any, *, top_level: bool) -> None:
        for key, value in source.items():
            key_text = str(key)
            if top_level and key_text in replace_sections:
                target[key] = deepcopy(value)
                continue
            current = target.get(key)
            if isinstance(current, Mapping) and isinstance(value, Mapping):
                overlay(current, value, top_level=False)
            else:
                target[key] = deepcopy(value)

    overlay(merged, user_doc, top_level=True)
    for section in keyed_list_sections:
        if section in replace_sections or section not in user_doc:
            continue
        merged[section] = _merge_keyed_records(base_doc.get(section, []), user_doc.get(section, []))
    return merged


def _toml_comment_texts(node: Any) -> list[str]:
    """Collect comment trivia recursively from a tomlkit document/item."""
    comments: list[str] = []
    visited: set[int] = set()

    def visit(value: Any) -> None:
        value_id = id(value)
        if value_id in visited:
            return
        visited.add(value_id)

        body = getattr(value, "body", None)
        if isinstance(body, list):
            for _key, item in body:
                try:
                    trivia = item.trivia
                except (AttributeError, RuntimeError):
                    trivia = None
                comment = str(getattr(trivia, "comment", "") or "").strip()
                if comment:
                    comments.append(comment)
                try:
                    nested = item.value
                except (AttributeError, RuntimeError):
                    nested = None
                if nested is not None:
                    visit(nested)
            return
        if isinstance(value, Mapping):
            for item in value.values():
                visit(item)
        elif isinstance(value, (list, tuple)):
            for item in value:
                visit(item)

    visit(node)
    return comments


def _editable_document_text(merged_doc: Any, user_doc: Any) -> str:
    """Serialize an overlay without silently dropping user comment text.

    tomlkit preserves inline trivia when a keyed item is copied, but detached
    comments that precede an overlaid table/key are not part of that item. Any
    such otherwise-lost comments are retained in a preamble. Inline dictionary
    remarks remain attached to their values; only detached comments may move.
    """
    merged_comments = _toml_comment_texts(merged_doc)
    missing_comments: list[str] = []
    for comment in _toml_comment_texts(user_doc):
        if comment in merged_comments:
            merged_comments.remove(comment)
        else:
            missing_comments.append(comment)

    merged_text = str(merged_doc.as_string())
    if not missing_comments:
        return merged_text
    return "".join(f"{comment}\n" for comment in missing_comments) + "\n" + merged_text


# ===================================================================
#  TOML I/O helper for the three-layer loader
# ===================================================================


def _read_toml_file(filepath: Path, *, missing_ok: bool = False) -> dict[str, Any]:
    """Read a TOML file for the three-layer loader.

    * ``missing_ok=False`` (base layer): a missing file raises ``FileNotFoundError``.
    * ``missing_ok=True`` (user layer): a missing file returns ``{}`` (no override).

    Parse errors are surfaced so the caller can distinguish an invalid shipped
    file from a recoverable user override.
    """
    if not filepath.exists():
        if missing_ok:
            return {}
        raise FileNotFoundError(f"Base config file not found: {filepath}")
    return read_toml_file(filepath)


def _quarantine_invalid_user_file(filepath: Path, *, reason: str) -> Path:
    """Durably preserve and remove one invalid sparse user override."""
    if filepath.is_symlink():
        raise ConfigSemanticError(f"refusing to quarantine symlinked user config: {filepath}")
    original = filepath.read_bytes()
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    backup = filepath.with_name(f"{filepath.name}.bak_{reason}_{timestamp}")
    atomic_write_bytes(backup, original)
    durable_unlink(filepath)
    logger.warning(
        "Invalid user config quarantined: %s -> %s | reason=%s",
        filepath,
        backup,
        reason,
    )
    return backup


_MISSING_CONFIG_VALUE = object()


_UserFilePreimage = config_transaction.UserFilePreimage


def _path_exists_or_is_symlink(path: Path) -> bool:
    """Return true for regular entries and broken symlink nodes."""
    return path.exists() or path.is_symlink()


def _capture_user_file_preimage(path: Path) -> _UserFilePreimage:
    """Capture bytes, link identity, and portable metadata for rollback."""
    return config_transaction.capture_user_file_preimage(path)


def _restore_user_file_preimage(preimage: _UserFilePreimage, *, operation: str) -> None:
    """Durably restore bytes, metadata, and logical symlink identity."""
    config_transaction.restore_user_file_preimage(
        preimage,
        operation=operation,
        atomic_write=atomic_write_bytes,
    )


def _restore_recovery_preimage(preimage: _UserFilePreimage, operation: str) -> None:
    _restore_user_file_preimage(preimage, operation=operation)


def _read_nested_value(container: Any, rel_key: tuple[str, ...]) -> Any:
    cursor = container
    for part in rel_key:
        if not isinstance(cursor, Mapping) or part not in cursor:
            return _MISSING_CONFIG_VALUE
        cursor = cursor[part]
    return cursor


def _write_nested_value(container: Any, rel_key: tuple[str, ...], value: Any) -> None:
    cursor = container
    for part in rel_key[:-1]:
        if part not in cursor or not isinstance(cursor[part], Mapping):
            cursor[part] = {}
        cursor = cursor[part]
    cursor[rel_key[-1]] = deepcopy(value)


def _delete_nested_value(
    container: Any,
    rel_key: tuple[str, ...],
    *,
    preserve_top_level: bool = False,
) -> None:
    cursor = container
    parents: list[tuple[Any, str]] = []
    for part in rel_key[:-1]:
        if not isinstance(cursor, Mapping) or part not in cursor:
            return
        parents.append((cursor, part))
        cursor = cursor[part]
    if isinstance(cursor, MutableMapping) and rel_key and rel_key[-1] in cursor:
        del cursor[rel_key[-1]]
    for parent, key in reversed(parents):
        if preserve_top_level and parent is container:
            continue
        try:
            child = parent[key]
        except Exception:
            continue
        if isinstance(child, Mapping) and len(child) == 0:
            del parent[key]


# ===================================================================
#  Config loader
# ===================================================================


class DocWenConfig:
    """Flat attribute-access wrapper around the merged config dict.

    Each top-level key of the merged dict is exposed as a property that
    lazily wraps sub-dicts in :class:`DocWenConfig` so you can write::

        config.gui.window.default_mode
        config.output.directory.mode
    """

    def __init__(self, data: dict[str, Any]) -> None:
        object.__setattr__(self, "_data", data)

    def __getattr__(self, name: str) -> Any:
        data = object.__getattribute__(self, "_data")
        if name in data:
            value = data[name]
            if isinstance(value, dict):
                return DocWenConfig(value)
            return value
        raise AttributeError(f"DocWenConfig 没有配置节: {name!r}")

    def __repr__(self) -> str:
        keys = list(object.__getattribute__(self, "_data").keys())
        return f"DocWenConfig({keys})"

    def as_dict(self) -> dict[str, Any]:
        """Return a deep copy of the raw dict."""
        import copy

        return copy.deepcopy(object.__getattribute__(self, "_data"))


class ConfigLoader:
    """Load and manage the DocWen configuration.

    Three-layer model:

    * **base** — shipped, read-only defaults in ``ResourceRegistry.default().configs_dir()``.
    * **user** — registry-governed overrides in
      ``platformdirs.user_config_dir("docwen")/configs``. Ordinary mappings are
      sparse; declared replacement sections are complete when present.
    * **runtime** — in-memory overrides passed via ``runtime_overrides`` (never persisted).

    Effective config is ``deep_merge(base_tree, user_overrides, runtime_overrides)``.
    ``reload()`` never writes disk and never backfills user files.

    Parameters:
        base_dir: Path to the read-only base ``configs/`` directory.  If *None*
            the loader uses ``ResourceRegistry.default().configs_dir()``.
        user_dir: Path to the writable user override directory.  If *None* the
            loader uses ``platformdirs.user_config_dir("docwen")/configs``.
        runtime_overrides: Optional dict merged on top of base+user in memory.
    """

    def __init__(
        self,
        *,
        base_dir: Path | str | None = None,
        user_dir: Path | str | None = None,
        runtime_overrides: Mapping[str, Any] | None = None,
    ) -> None:
        if base_dir is not None:
            self._base_dir = Path(base_dir)
        else:
            from docwen_runtime.resources import ResourceRegistry

            self._base_dir = ResourceRegistry.default().configs_dir()
        self._user_dir = Path(user_dir) if user_dir is not None else _default_user_config_dir()
        self._runtime_overrides: dict[str, Any] = dict(runtime_overrides or {})
        self._config: dict[str, Any] = {}
        self._config_state_trusted = False
        self.reload()

    # ---- public API -----------------------------------------------------------

    @property
    def config(self) -> DocWenConfig:
        """The merged configuration as a lazily-wrapped attribute tree."""
        return DocWenConfig(self._config)

    @property
    def config_state_trusted(self) -> bool:
        """Whether the last reload completed all cache and runtime wiring."""
        return self._config_state_trusted

    @property
    def base_dir(self) -> Path:
        """Absolute path to the read-only base ``configs/`` directory."""
        return self._base_dir

    @property
    def user_dir(self) -> Path:
        """Absolute path to the writable user override directory."""
        return self._user_dir

    def _allowed_user_paths(self) -> tuple[Path, ...]:
        """Return the registry-owned logical paths accepted by recovery."""
        return tuple(self._user_dir / spec.rel_path for spec in CONFIG_FILES)

    def reload(self) -> None:
        """Recover under the process lock, then publish one trusted reload."""
        with _CONFIG_TRANSACTION_LOCK, config_transaction.process_config_lock(self._user_dir):
            self._config_state_trusted = False
            if getattr(_CONFIG_TRANSACTION_STATE, "operation", None) is None:
                config_transaction.recover_transaction_journal(
                    self._user_dir,
                    self._allowed_user_paths(),
                    _restore_recovery_preimage,
                )
            self._reload_unlocked()
            self._config_state_trusted = True

    def _reload_unlocked(self) -> None:
        """Re-read all registry config files from base/user/runtime and merge.

        For every ``ConfigFileSpec`` in the registry: read its base file
        (must exist) and its user file (optional — missing means no override),
        deep-merge them per file, wrap the result under the spec's namespace,
        and accumulate into the merged tree. Runtime overrides are applied
        last. Invalid shipped/runtime data fails closed; an invalid pre-existing
        sparse user override is quarantined unless a write transaction is active.
        """
        merged: dict[str, Any] = {}
        shipped_tree: dict[str, Any] = {}
        for spec in CONFIG_FILES:
            base_path = self._base_dir / spec.rel_path
            user_path = self._user_dir / spec.rel_path
            base_data = _read_toml_file(base_path, missing_ok=False)
            base_data = validate_config_file(spec.rel_path, base_data, base_data)
            shipped_tree = deep_merge(shipped_tree, wrap_namespace(base_data, spec.namespace))

            try:
                user_data = _read_toml_file(user_path, missing_ok=True)
                file_data = _merge_file_layers(spec, base_data, user_data)
                file_data = validate_config_file(spec.rel_path, file_data, base_data)
            except (TOMLDecodeError, ConfigSemanticError) as exc:
                if getattr(_CONFIG_TRANSACTION_STATE, "operation", None) is not None:
                    raise
                try:
                    same_physical_path = base_path.resolve(strict=False) == user_path.resolve(strict=False)
                except OSError:
                    same_physical_path = base_path.absolute() == user_path.absolute()
                if same_physical_path or not _path_exists_or_is_symlink(user_path):
                    raise
                reason = "parse_failed" if isinstance(exc, TOMLDecodeError) else "schema_failed"
                _quarantine_invalid_user_file(user_path, reason=reason)
                file_data = deepcopy(base_data)
            merged = deep_merge(merged, wrap_namespace(file_data, spec.namespace))

        if self._runtime_overrides:
            merged = deep_merge(merged, self._runtime_overrides)

        for spec in CONFIG_FILES:
            effective_file: Any = merged
            shipped_file: Any = shipped_tree
            for namespace_part in spec.namespace:
                if isinstance(effective_file, Mapping):
                    effective_file = effective_file.get(namespace_part, {})
                if isinstance(shipped_file, Mapping):
                    shipped_file = shipped_file.get(namespace_part, {})
            validated_file = validate_config_file(spec.rel_path, effective_file, shipped_file)
            cursor = merged
            for namespace_part in spec.namespace[:-1]:
                cursor = cursor[namespace_part]
            cursor[spec.namespace[-1]] = validated_file

        self._config = merged
        self._wire_logging()

    def _wire_logging(self) -> None:
        """Wire the merged configuration to the logging subsystem.

        Reads the ``logger`` config section and calls :func:`~docwen_runtime.logging.init_logging`
        to build console and file handlers.  Called automatically from :meth:`reload`.
        """
        from docwen_runtime.logging import init_logging

        init_logging(self._config)

    # ---- write / reset / read API --------------------------------------------

    def _user_path_for_spec(self, spec: Any) -> Path:
        """Return the user override path for a registry spec."""
        return self._user_dir / spec.rel_path

    def _run_user_file_transaction(
        self,
        user_paths: list[Path] | tuple[Path, ...],
        mutate: Any,
        *,
        operation: str,
    ) -> bool:
        """Serialize, recover, and execute one durable config transaction."""
        with _CONFIG_TRANSACTION_LOCK:
            try:
                with config_transaction.process_config_lock(self._user_dir):
                    return self._run_user_file_transaction_process_locked(
                        user_paths,
                        mutate,
                        operation=operation,
                    )
            except Exception as exc:
                logger.error(
                    "Configuration transaction lock failed: operation=%s error=%s",
                    operation,
                    exc,
                )
                return False

    def _run_user_file_transaction_process_locked(
        self,
        user_paths: list[Path] | tuple[Path, ...],
        mutate: Any,
        *,
        operation: str,
    ) -> bool:
        """Recover and mutate while the user-directory process lock is held."""
        active_operation = getattr(_CONFIG_TRANSACTION_STATE, "operation", None)
        if active_operation is not None:
            _CONFIG_TRANSACTION_STATE.nested_attempted = True
            logger.error(
                "Nested configuration persistence rejected: outer=%s nested=%s",
                active_operation,
                operation,
            )
            return False

        try:
            recovered = config_transaction.recover_transaction_journal(
                self._user_dir,
                self._allowed_user_paths(),
                _restore_recovery_preimage,
            )
            if recovered is not None:
                self._config_state_trusted = False
                self._reload_unlocked()
                self._config_state_trusted = True
        except Exception as exc:
            self._config_state_trusted = False
            logger.error("Configuration transaction recovery failed: %s", exc)
            return False

        _CONFIG_TRANSACTION_STATE.operation = operation
        _CONFIG_TRANSACTION_STATE.nested_attempted = False
        try:
            return self._run_user_file_transaction_locked(
                user_paths,
                mutate,
                operation=operation,
            )
        finally:
            del _CONFIG_TRANSACTION_STATE.operation
            del _CONFIG_TRANSACTION_STATE.nested_attempted

    def _run_user_file_transaction_locked(
        self,
        user_paths: list[Path] | tuple[Path, ...],
        mutate: Any,
        *,
        operation: str,
    ) -> bool:
        """Apply one config operation with a durable before-image journal.

        Each individual write uses a same-directory staged replacement.
        ``PREPARED`` is durable before mutation. Every user-file replacement or
        deletion orders its parent directory. ``COMMITTED`` is published only
        after the new generation reloads successfully, and is the transaction
        linearization point; journal deletion is cleanup only.
        """
        paths = tuple(dict.fromkeys(Path(path) for path in user_paths))
        try:
            preimages = {path: _capture_user_file_preimage(path) for path in paths}
        except Exception as exc:
            logger.error(
                "Configuration transaction snapshot failed: operation=%s error=%s",
                operation,
                exc,
            )
            return False

        try:
            config_transaction.write_transaction_journal(
                self._user_dir,
                operation,
                preimages.values(),
                state="PREPARED",
            )
        except Exception as exc:
            logger.error(
                "Configuration transaction journal prepare failed: operation=%s error=%s",
                operation,
                exc,
            )
            return False

        try:
            mutate()
            if _CONFIG_TRANSACTION_STATE.nested_attempted:
                raise RuntimeError("nested configuration persistence was rejected")
            self.reload()
            if _CONFIG_TRANSACTION_STATE.nested_attempted:
                raise RuntimeError("nested configuration persistence was rejected")
            config_transaction.mark_transaction_committed(
                self._user_dir,
                operation,
                preimages.values(),
            )
        except Exception as exc:
            logger.error(
                "Configuration transaction failed: operation=%s error=%s; restoring %d file preimage(s)",
                operation,
                exc,
                len(preimages),
            )
            rollback_ok = True
            for preimage in reversed(preimages.values()):
                try:
                    _restore_user_file_preimage(preimage, operation=operation)
                except Exception as rollback_exc:
                    rollback_ok = False
                    logger.error(
                        "Configuration transaction rollback failed: operation=%s path=%s error=%s",
                        operation,
                        preimage.path,
                        rollback_exc,
                    )

            if rollback_ok:
                try:
                    config_transaction.remove_transaction_journal(self._user_dir)
                except Exception as cleanup_exc:
                    logger.warning(
                        "Prepared configuration journal cleanup deferred: operation=%s error=%s",
                        operation,
                        cleanup_exc,
                    )
            try:
                self.reload()
            except Exception as reload_exc:
                logger.error(
                    "Configuration transaction reconciliation failed: operation=%s rollback_ok=%s error=%s",
                    operation,
                    rollback_ok,
                    reload_exc,
                )
            if not rollback_ok:
                self._config_state_trusted = False
            return False

        try:
            config_transaction.remove_transaction_journal(self._user_dir)
        except Exception as cleanup_exc:
            logger.warning(
                "Committed configuration journal cleanup deferred: operation=%s error=%s",
                operation,
                cleanup_exc,
            )
        return True

    def _mutate_user_document(self, spec: Any, mutate: Any) -> None:
        """Mutate and staged-replace one user TOML document without reload."""
        from docwen_runtime.toml_io import (
            load_toml_document,
            new_toml_document,
            save_toml_document,
        )

        user_path = self._user_path_for_spec(spec)
        user_path.parent.mkdir(parents=True, exist_ok=True)
        doc = load_toml_document(user_path) if user_path.exists() else new_toml_document()
        _reject_reserved_user_keys(doc)
        mutate(doc)
        _reject_reserved_user_keys(doc)
        save_toml_document(user_path, doc)

    def _plan_reset_values(
        self,
        dotted_keys: list[str] | tuple[str, ...],
    ) -> dict[str, list[tuple[str, ...]]] | None:
        """Route selected dotted keys to relative paths grouped by owner file."""
        by_file: dict[str, list[tuple[str, ...]]] = {}
        for dotted_key in dotted_keys:
            spec = spec_for_dotted_key(dotted_key)
            if spec is None:
                logger.error("无法路由配置键: %s", dotted_key)
                return None
            try:
                rel_key = relative_key_for_spec(spec, dotted_key)
            except KeyError:
                return None
            by_file.setdefault(spec.rel_path, []).append(rel_key)
        return by_file

    def _reset_grouped_values_on_disk(
        self,
        by_file: Mapping[str, list[tuple[str, ...]]],
    ) -> None:
        """Reset already-routed user values without starting a transaction or reload."""
        for rel_path, rel_keys in by_file.items():
            spec = require_spec(rel_path)
            user_path = self._user_path_for_spec(spec)
            if user_path.is_symlink() and not user_path.exists():
                raise OSError(f"cannot reset selected values through broken symlink: {user_path}")
            if not user_path.exists():
                continue
            base_data = self.get_base_file_dict(rel_path)

            def _mutate(
                doc: Any,
                keys: list[tuple[str, ...]] = rel_keys,
                base: dict[str, Any] = base_data,
                file_spec: Any = spec,
            ) -> None:
                complete_sections = frozenset(file_spec.replace_sections).intersection(str(key) for key in doc)
                top_level_resets = {key[0] for key in keys if len(key) == 1}
                for rel_key in keys:
                    if not rel_key or (len(rel_key) > 1 and rel_key[0] in top_level_resets):
                        continue
                    if len(rel_key) > 1 and rel_key[0] in complete_sections:
                        base_value = _read_nested_value(base, rel_key)
                        if base_value is _MISSING_CONFIG_VALUE:
                            _delete_nested_value(doc, rel_key, preserve_top_level=True)
                        else:
                            _write_nested_value(doc, rel_key, base_value)
                        continue
                    _delete_nested_value(doc, rel_key)

            self._mutate_user_document(spec, _mutate)

    def reset_file(self, rel_path: str) -> bool:
        """Delete the user override for *rel_path*, revealing base defaults.

        - unknown path → ``False``
        - user file exists → delete, reload, return ``True``
        - user file missing → reload (effective config already equals base),
          return ``True``
        - never writes base file
        """
        if rel_path in RESET_EXCLUDED:
            logger.warning("词库配置不参与还原: %s", rel_path)
            return False
        try:
            spec = require_spec(rel_path)
        except KeyError:
            logger.error("未知的配置文件: %s", rel_path)
            return False
        user_path = self._user_path_for_spec(spec)

        def _delete_override() -> None:
            if _path_exists_or_is_symlink(user_path):
                durable_unlink(user_path)

        if not self._run_user_file_transaction(
            [user_path],
            _delete_override,
            operation=f"reset_file:{rel_path}",
        ):
            return False
        logger.info("配置已还原: %s", rel_path)
        return True

    def reset_section(self, dotted_section: str) -> bool:
        """Route *dotted_section* to its spec via registry and reset the user file.

        Since each spec maps to exactly one file, resetting a section is
        equivalent to resetting the file that owns that namespace.
        """
        if not dotted_section:
            return False
        spec = spec_for_dotted_key(dotted_section)
        if spec is None:
            logger.error("无法路由配置节: %s", dotted_section)
            return False
        return self.reset_file(spec.rel_path)

    def reset_values(self, dotted_keys: list[str] | tuple[str, ...]) -> bool:
        """Delete selected user override values, revealing base defaults.

        Unlike :meth:`reset_file`, this preserves sibling values in the same
        user override file. It is useful for GUI tabs that own only a few keys
        inside a shared file such as ``software.toml``.
        """
        by_file = self._plan_reset_values(dotted_keys)
        if by_file is None:
            return False
        user_paths = [self._user_path_for_spec(require_spec(rel_path)) for rel_path in by_file]

        def _reset_selected_values() -> None:
            self._reset_grouped_values_on_disk(by_file)

        return self._run_user_file_transaction(
            user_paths,
            _reset_selected_values,
            operation="reset_values",
        )

    def reset_group(self, group: str) -> bool:
        """Reset one logical settings group without crossing tab ownership.

        Whole-file groups are derived from the config registry.  Shared-file
        groups use the registry's dotted-key plan, and user-curated files in
        :data:`RESET_EXCLUDED` are deliberately preserved without making the
        otherwise successful group reset report a partial failure.
        """

        plan = reset_plan_for_group(group)
        if not plan.files and not plan.dotted_keys:
            logger.error("未知的配置组: %s", group)
            return False

        whole_specs = [require_spec(rel_path) for rel_path in plan.files if rel_path not in RESET_EXCLUDED]
        dotted_by_file = self._plan_reset_values(plan.dotted_keys)
        if dotted_by_file is None:
            return False
        dotted_specs = [require_spec(rel_path) for rel_path in dotted_by_file]
        user_paths = [self._user_path_for_spec(spec) for spec in (*whole_specs, *dotted_specs)]

        def _reset_group_on_disk() -> None:
            for spec in whole_specs:
                user_path = self._user_path_for_spec(spec)
                if _path_exists_or_is_symlink(user_path):
                    durable_unlink(user_path)
            self._reset_grouped_values_on_disk(dotted_by_file)

        return self._run_user_file_transaction(
            user_paths,
            _reset_group_on_disk,
            operation=f"reset_group:{group}",
        )

    def reset_all(self) -> bool:
        """Reset every registered config file except user-curated data."""
        reset_specs = [spec for spec in CONFIG_FILES if spec.rel_path not in RESET_EXCLUDED]
        user_paths = [self._user_path_for_spec(spec) for spec in reset_specs]

        def _reset_all_on_disk() -> None:
            for user_path in user_paths:
                if _path_exists_or_is_symlink(user_path):
                    durable_unlink(user_path)

        return self._run_user_file_transaction(
            user_paths,
            _reset_all_on_disk,
            operation="reset_all",
        )

    def set_value(self, dotted_key: str, value: Any) -> bool:
        """Write a single value into the correct user override file via registry.

        Uses ``spec_for_dotted_key`` to find the owning spec, then
        ``relative_key_for_spec`` to determine the relative key path within
        the file.  Creates user directories as needed.  Calls :meth:`reload`
        after writing so the in-memory config reflects the change.
        """
        return self.set_values({dotted_key: value})

    def set_values(self, values: Mapping[str, Any]) -> bool:
        """Write multiple dotted values, grouped by config file, then reload once.

        This is equivalent to calling :meth:`set_value` for each key, except
        related keys are coalesced per TOML file. Handled write or reload
        failures restore every target's byte-exact preimage before the loader
        reconciles its effective configuration.
        """
        by_file: dict[str, list[tuple[tuple[str, ...], Any]]] = {}
        try:
            for dotted_key, value in values.items():
                spec = spec_for_dotted_key(dotted_key)
                if spec is None:
                    logger.error("无法路由配置键: %s", dotted_key)
                    return False
                rel_key = relative_key_for_spec(spec, dotted_key)
                by_file.setdefault(spec.rel_path, []).append((rel_key, deepcopy(value)))
        except Exception as exc:
            logger.error("Configuration batch planning failed: %s", exc)
            return False

        if not by_file:
            return True
        user_paths = [self._user_path_for_spec(require_spec(rel_path)) for rel_path in by_file]

        def _write_grouped_values() -> None:
            for rel_path, entries in by_file.items():
                spec = require_spec(rel_path)
                user_path = self._user_path_for_spec(spec)
                user_path.parent.mkdir(parents=True, exist_ok=True)
                user_data = _read_toml_file(user_path, missing_ok=True)
                _reject_reserved_user_keys(user_data)
                for rel_key, value in entries:
                    cursor = user_data
                    for part in rel_key[:-1]:
                        if part not in cursor or not isinstance(cursor[part], dict):
                            cursor[part] = {}
                        cursor = cursor[part]
                    cursor[rel_key[-1]] = value
                _reject_reserved_user_keys(user_data)
                write_toml_file(user_path, user_data)

        return self._run_user_file_transaction(
            user_paths,
            _write_grouped_values,
            operation="set_values",
        )

    def write_file_content(self, rel_path: str, data: dict[str, Any]) -> bool:
        """Write entire content to a user override file, then reload.

        Args:
            rel_path: Registry path (e.g. ``"gui.toml"``).
            data: The complete dict to write.

        Returns:
            True on success, False on failure.
        """
        try:
            spec = require_spec(rel_path)
        except KeyError:
            logger.error("未知的配置文件: %s", rel_path)
            return False
        try:
            user_data = deepcopy(data)
            _reject_reserved_user_keys(user_data)
        except Exception as exc:
            logger.error("配置文件序列化准备失败: %s | 错误: %s", rel_path, exc)
            return False
        user_path = self._user_path_for_spec(spec)

        def _write_content() -> None:
            write_toml_file(user_path, user_data)

        return self._run_user_file_transaction(
            [user_path],
            _write_content,
            operation=f"write_file_content:{rel_path}",
        )

    def get_file_text(self, rel_path: str) -> str | None:
        """Return the editable base+user TOML source for a registered file.

        A missing user override returns the shipped source verbatim. Ordinary
        user mappings are overlaid sparsely. Registry-declared replacement
        sections replace the shipped section wholesale whenever present.
        Comments from the effective editable document remain available.
        Runtime-only overrides are deliberately excluded: opening an editor
        must not silently turn ephemeral request overrides into persisted user
        configuration.
        """
        try:
            spec = require_spec(rel_path)
        except KeyError:
            logger.error("未知的配置文件: %s", rel_path)
            return None

        base_path = self._base_dir / spec.rel_path
        user_path = self._user_dir / spec.rel_path
        try:
            base_text = base_path.read_text(encoding="utf-8")
            if not user_path.exists():
                return base_text

            from docwen_core.toml_tools import read_toml_text

            base_doc = read_toml_text(base_text)
            user_doc = read_toml_text(user_path.read_text(encoding="utf-8"))
            _reject_reserved_user_keys(user_doc)
            merged = _merge_toml_documents(
                base_doc,
                user_doc,
                spec.replace_sections,
                spec.keyed_list_sections,
            )
            return _editable_document_text(merged, user_doc)
        except Exception as exc:
            logger.error("读取可编辑配置源失败: %s | 错误: %s", rel_path, exc)
            return None

    def save_file_text(self, rel_path: str, content: str) -> bool:
        """Persist validated TOML *content* to the user override file.

        The content is written without hidden metadata. Registry-declared
        replacement sections are complete whenever present; all other mappings
        remain sparse. Calls :meth:`reload` after writing.

        Returns ``False`` for unknown *rel_path*, invalid TOML, or write
        failure.
        """
        try:
            spec = require_spec(rel_path)
        except KeyError:
            logger.error("未知的配置文件: %s", rel_path)
            return False

        # Validate TOML syntax and reserved keys before touching disk.
        try:
            from docwen_core.toml_tools import read_toml_text

            doc = read_toml_text(content or "")
            _reject_reserved_user_keys(doc)
        except Exception as exc:
            logger.error("TOML 语法错误 (%s): %s", rel_path, exc)
            return False

        user_path = self._user_path_for_spec(spec)

        def _write_text() -> None:
            atomic_write_text(user_path, content)

        return self._run_user_file_transaction(
            [user_path],
            _write_text,
            operation=f"save_file_text:{rel_path}",
        )

    def update_file_sections(self, rel_path: str, sections: dict[str, Any]) -> bool:
        """Write top-level *sections* into *rel_path*'s user file, preserving
        comments in untouched sections via tomlkit.

        Each key in *sections* replaces the entire corresponding TOML section
        wholesale (not per-key deep merge).  Calls :meth:`reload` after writing.
        """
        try:
            spec = require_spec(rel_path)
        except KeyError:
            logger.error("未知的配置文件: %s", rel_path)
            return False

        from docwen_runtime.toml_io import update_toml_document_sections

        user_path = self._user_path_for_spec(spec)

        def _update_sections(doc: Any) -> None:
            update_toml_document_sections(doc, sections)

        def _write_sections() -> None:
            self._mutate_user_document(spec, _update_sections)

        return self._run_user_file_transaction(
            [user_path],
            _write_sections,
            operation=f"update_file_sections:{rel_path}",
        )

    def update_file_document(self, rel_path: str, mutate: Any) -> bool:
        """Hand the tomlkit document for *rel_path*'s user file to *mutate(doc)*
        for fine-grained mutation, then write back with comments preserved.

        *mutate* receives a mutable ``TOMLDocument`` and should modify it
        in-place (no return value expected).  If *mutate* raises, this
        method returns ``False`` and the file is left untouched.

        Calls :meth:`reload` after a successful write.
        """
        try:
            spec = require_spec(rel_path)
        except KeyError:
            logger.error("未知的配置文件: %s", rel_path)
            return False

        user_path = self._user_path_for_spec(spec)

        def _write_document() -> None:
            self._mutate_user_document(spec, mutate)

        return self._run_user_file_transaction(
            [user_path],
            _write_document,
            operation=f"update_file_document:{rel_path}",
        )

    def get_file_dict(self, rel_path: str) -> dict[str, Any]:
        """Read the user override file as a plain dict (no deep-merge).

        Returns an empty dict when the file is missing, unknown, or
        parsing fails.
        """
        try:
            spec = require_spec(rel_path)
        except KeyError:
            logger.error("未知的配置文件: %s", rel_path)
            return {}
        raw = _read_toml_file(self._user_dir / spec.rel_path, missing_ok=True)
        _reject_reserved_user_keys(raw)
        return raw

    def get_base_file_dict(self, rel_path: str) -> dict[str, Any]:
        """Read the base (shipped) file as a plain dict.

        Raises ``KeyError`` for unknown *rel_path* and ``FileNotFoundError``
        for missing base file.
        """
        spec = require_spec(rel_path)
        return _read_toml_file(self._base_dir / spec.rel_path, missing_ok=False)

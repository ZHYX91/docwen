from __future__ import annotations

from collections.abc import Mapping
from pathlib import Path
from typing import Any


class ConfigPortAdapter:
    def __init__(
        self,
        loader: Any | None = None,
        *,
        base_dir: str | Path | None = None,
        user_dir: str | Path | None = None,
        runtime_overrides: Mapping[str, Any] | None = None,
    ) -> None:
        from docwen_runtime.config import ConfigLoader

        if loader is not None and any(value is not None for value in (base_dir, user_dir, runtime_overrides)):
            raise ValueError("Inject a ConfigLoader or construction paths, not both")
        self._loader = loader or ConfigLoader(
            base_dir=base_dir,
            user_dir=user_dir,
            runtime_overrides=runtime_overrides,
        )

    def _trusted_config(self) -> Any:
        """Return effective config only after cache/runtime wiring completed."""
        if getattr(self._loader, "config_state_trusted", True) is not True:
            raise RuntimeError("configuration state is untrusted after a failed reconciliation; reload required")
        return self._loader.config

    def get(self, key: str, default: Any = None) -> Any:
        data = self._trusted_config().as_dict()
        parts = [p for p in (key or "").split(".") if p]
        if not parts:
            return default
        cur: Any = data
        for part in parts:
            if not isinstance(cur, dict) or part not in cur:
                return default
            cur = cur[part]
        return cur

    def snapshot(self) -> dict[str, Any]:
        return self._trusted_config().as_dict()

    def set(self, key: str, value: Any) -> bool:
        return self._loader.set_value(key, value)

    def set_many(self, values: dict[str, Any]) -> bool:
        return self._loader.set_values(values)

    def get_file_text(self, rel_path: str) -> str | None:
        return self._loader.get_file_text(rel_path)

    def save_file_text(self, rel_path: str, content: str) -> bool:
        return self._loader.save_file_text(rel_path, content)

    def reset_file(self, rel_path: str) -> bool:
        return self._loader.reset_file(rel_path)

    def reset_group(self, group: str) -> bool:
        return self._loader.reset_group(group)

    def reset_all(self) -> bool:
        """Reset all registered config files to defaults."""
        return self._loader.reset_all()

    def reload(self) -> None:
        self._loader.reload()

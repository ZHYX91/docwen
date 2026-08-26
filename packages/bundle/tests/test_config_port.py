from __future__ import annotations

import pytest

pytestmark = pytest.mark.unit


def test_config_port_adapter_rejects_reads_while_loader_state_is_untrusted() -> None:
    from docwen_bundle.config_port import ConfigPortAdapter

    class _Config:
        def as_dict(self) -> dict[str, object]:
            return {"gui": {"theme": {"default_theme": "light"}}}

    class _Loader:
        config_state_trusted = False
        config = _Config()

    loader = _Loader()
    adapter = ConfigPortAdapter(loader)

    with pytest.raises(RuntimeError, match="configuration state is untrusted"):
        adapter.snapshot()
    with pytest.raises(RuntimeError, match="configuration state is untrusted"):
        adapter.get("gui.theme.default_theme")

    loader.config_state_trusted = True
    assert adapter.snapshot() == {"gui": {"theme": {"default_theme": "light"}}}
    assert adapter.get("gui.theme.default_theme") == "light"


def test_config_port_adapter_delegates_logical_group_reset() -> None:
    from docwen_bundle.config_port import ConfigPortAdapter

    calls: list[str] = []

    class _Loader:
        def reset_group(self, group: str) -> bool:
            calls.append(group)
            return group == "formatting"

        def reset_all(self) -> bool:
            calls.append("all")
            return True

    loader = _Loader()
    adapter = ConfigPortAdapter(loader)

    assert adapter.reset_group("formatting") is True
    assert adapter.reset_all() is True
    assert calls == ["formatting", "all"]


def test_config_port_adapter_delegates_editable_file_text() -> None:
    from docwen_bundle.config_port import ConfigPortAdapter

    calls: list[tuple[str, str]] = []

    class _Loader:
        def get_file_text(self, rel_path: str) -> str | None:
            calls.append(("read", rel_path))
            return "[entries]\n"

        def save_file_text(self, rel_path: str, content: str) -> bool:
            calls.append((rel_path, content))
            return True

    adapter = ConfigPortAdapter(_Loader())

    assert adapter.get_file_text("proofread/typos.toml") == "[entries]\n"
    assert adapter.save_file_text("proofread/typos.toml", "[entries]\n") is True
    assert calls == [
        ("read", "proofread/typos.toml"),
        ("proofread/typos.toml", "[entries]\n"),
    ]

"""Focused tests split from test_entrypoints.py."""

from __future__ import annotations

from ._entrypoints_support import (
    pytest,
)

pytestmark = pytest.mark.unit


class TestCliEntry:
    def test_cli_entry_wires_runtime_and_config_ports(self, monkeypatch: pytest.MonkeyPatch) -> None:
        import docwen_bundle.cli_entry as cli_entry
        import docwen_bundle.config_port as config_port_module
        import docwen_bundle.runtime_factory as runtime_factory_module
        import docwen_cli.main as cli_main_module
        import docwen_runtime.logging as runtime_logging

        runtime_port = object()
        _captured_runtime_factory: object | None = None
        _captured_config_factory: object | None = None
        _captured_gui_control_factory: object | None = None
        _captured_machine_server_factory: object | None = None
        _captured_argv: object | None = None
        runtime_loaders: list[object] = []
        pre_init_levels: list[str] = []
        monkeypatch.setattr(runtime_logging, "pre_init_logging", pre_init_levels.append)

        def _fake_runtime(*, config_loader):
            runtime_loaders.append(config_loader)
            return runtime_port

        monkeypatch.setattr(runtime_factory_module, "create_runtime_port", _fake_runtime)

        class _ConfigPort:
            def __init__(self, loader=None, **kwargs) -> None:
                self.loader = loader
                self.kwargs = kwargs

        monkeypatch.setattr(config_port_module, "ConfigPortAdapter", _ConfigPort)

        def _fake_cli_main(
            argv=None,
            *,
            runtime_port_factory=None,
            config_port_factory=None,
            gui_control_port_factory=None,
            machine_server_factory=None,
        ):
            nonlocal _captured_argv
            nonlocal _captured_runtime_factory
            nonlocal _captured_config_factory
            nonlocal _captured_gui_control_factory
            nonlocal _captured_machine_server_factory
            _captured_argv = argv
            _captured_runtime_factory = runtime_port_factory
            _captured_config_factory = config_port_factory
            _captured_gui_control_factory = gui_control_port_factory
            _captured_machine_server_factory = machine_server_factory
            return 7

        monkeypatch.setattr(cli_main_module, "main", _fake_cli_main)

        exit_code = cli_entry.main(["convert", "a.docx", "--to", "md", "--output", "a.md"])

        assert exit_code == 7
        assert _captured_argv == ["convert", "a.docx", "--to", "md", "--output", "a.md"]
        assert callable(_captured_runtime_factory)
        assert _captured_runtime_factory() is runtime_port
        assert callable(_captured_config_factory)
        config = _captured_config_factory()
        assert isinstance(config, _ConfigPort)
        assert config.kwargs == {}
        assert runtime_loaders == [config.loader]
        assert _captured_gui_control_factory is not None
        assert _captured_machine_server_factory is not None
        assert pre_init_levels == ["WARNING"]

    def test_pyi_cli_entry_delegates_and_preserves_exit_code(self, monkeypatch: pytest.MonkeyPatch) -> None:
        import docwen_bundle.cli_entry as cli_entry
        import docwen_bundle.pyi_cli_entry as pyi

        monkeypatch.setattr(cli_entry, "main", lambda argv=None: 11)

        exit_code = pyi._delegate()

        assert exit_code == 11

    def test_pyi_cli_entry_prepares_frozen_multiprocessing(self, monkeypatch: pytest.MonkeyPatch) -> None:
        import docwen_bundle.pyi_cli_entry as pyi

        calls: list[bool] = []
        monkeypatch.setattr(pyi.multiprocessing, "freeze_support", lambda: calls.append(True))

        pyi._prepare_multiprocessing()

        assert calls == [True]

"""Request-admission tests for RuntimePort OCR option projection."""

from __future__ import annotations

from copy import deepcopy
from pathlib import Path
from typing import Any

import pytest

from docwen_core.models.file_ref import FileRef
from docwen_core.models.request import ConversionRequest, OutputPolicy
from docwen_core.models.result import ConversionResult
from docwen_runtime.adapters import RuntimePortAdapter

pytestmark = pytest.mark.unit
PROJECT_CONFIGS = Path(__file__).resolve().parents[3] / "configs"


class _Recorder:
    def __init__(self) -> None:
        self.requests: list[ConversionRequest] = []

    def execute_single(self, request: ConversionRequest, on_event: Any = None) -> ConversionResult:
        self.requests.append(request)
        return ConversionResult(task_id=request.request_id, success=True)

    def execute_batch(
        self,
        request: ConversionRequest,
        on_event: Any = None,
    ) -> list[ConversionResult]:
        self.requests.append(request)
        return [ConversionResult(task_id=request.request_id, success=True)]

    def cancel(self, task_id: str) -> None:
        pass

    def cancel_all(self) -> int:
        return 0


class _SnapshotConfig:
    def __init__(self, snapshot: dict[str, Any]) -> None:
        self._snapshot = deepcopy(snapshot)

    def as_dict(self) -> dict[str, Any]:
        return deepcopy(self._snapshot)


class _ConfigLoader:
    def __init__(self, snapshot: dict[str, Any]) -> None:
        self.config = _SnapshotConfig(snapshot)


def _snapshot(*, ocr_language: str, locale: str) -> dict[str, Any]:
    return {
        "image": {"ocr_language": ocr_language},
        "gui": {"language": {"locale": locale}},
    }


@pytest.fixture
def runtime_input(tmp_path: Path) -> Path:
    source = tmp_path / "runtime-input.md"
    source.write_text("# Runtime input\n", encoding="utf-8")
    return source


def _request(
    input_path: Path,
    *,
    options: dict[str, Any] | None = None,
    config_snapshot: dict[str, Any] | None = None,
    target_format: str = "md",
    action_name: str = "",
) -> ConversionRequest:
    return ConversionRequest(
        request_id="runtime-ocr-projection",
        input_refs=[FileRef(path=str(input_path), format="markdown", category="markdown")],
        target_format=target_format,
        action_name=action_name,
        options=dict(options or {}),
        output_policy=OutputPolicy(),
        config_snapshot=deepcopy(config_snapshot or {}),
    )


def test_empty_options_use_the_same_loader_snapshot_as_runtime_context(runtime_input: Path) -> None:
    recorder = _Recorder()
    loader = _ConfigLoader(_snapshot(ocr_language="japanese", locale="ja_JP"))
    adapter = RuntimePortAdapter(recorder, config_loader=loader)  # type: ignore[arg-type]
    request = _request(runtime_input)

    adapter.execute(request)

    admitted = recorder.requests[0]
    assert admitted.config_snapshot == _snapshot(ocr_language="japanese", locale="ja_JP")
    assert admitted.options == {"ocr_language": "japanese", "locale": "ja_JP"}
    assert request.options == {}
    assert request.config_snapshot == {}


def test_real_config_loader_projects_reloaded_ocr_values_without_mutating_request(
    tmp_path: Path,
    runtime_input: Path,
) -> None:
    from docwen_runtime.config.loader import ConfigLoader

    loader = ConfigLoader(base_dir=PROJECT_CONFIGS, user_dir=tmp_path / "user-config")
    assert loader.set_values(
        {
            "image.ocr_language": "japanese",
            "gui.language.locale": "ja_JP",
        }
    )
    recorder = _Recorder()
    adapter = RuntimePortAdapter(recorder, config_loader=loader)  # type: ignore[arg-type]
    request = _request(runtime_input)

    adapter.execute(request)

    admitted = recorder.requests[0]
    assert admitted.options == {"ocr_language": "japanese", "locale": "ja_JP"}
    assert admitted.config_snapshot["image"]["ocr_language"] == "japanese"
    assert admitted.config_snapshot["gui"]["language"]["locale"] == "ja_JP"
    assert request.options == {}
    assert request.config_snapshot == {}


def test_request_snapshot_is_authoritative_over_the_live_loader(runtime_input: Path) -> None:
    recorder = _Recorder()
    loader = _ConfigLoader(_snapshot(ocr_language="english", locale="en_US"))
    adapter = RuntimePortAdapter(recorder, config_loader=loader)  # type: ignore[arg-type]
    request_snapshot = _snapshot(ocr_language="korean", locale="ko_KR")
    request = _request(runtime_input, config_snapshot=request_snapshot)

    adapter.execute(request)

    admitted = recorder.requests[0]
    assert admitted is not request
    assert admitted.config_snapshot == request_snapshot
    assert admitted.options == {"ocr_language": "korean", "locale": "ko_KR"}
    assert request.options == {}


@pytest.mark.parametrize(
    ("request_snapshot", "expected"),
    [
        (
            {"image": {"ocr_language": "japanese"}},
            {"ocr_language": "japanese", "locale": "zh_CN"},
        ),
        (
            {"gui": {"language": {"locale": "ja_JP"}}},
            {"ocr_language": "auto", "locale": "ja_JP"},
        ),
    ],
)
def test_partial_request_snapshot_uses_protocol_defaults_not_live_loader(
    runtime_input: Path,
    request_snapshot: dict[str, Any],
    expected: dict[str, Any],
) -> None:
    recorder = _Recorder()
    loader = _ConfigLoader(_snapshot(ocr_language="english", locale="en_US"))
    adapter = RuntimePortAdapter(recorder, config_loader=loader)  # type: ignore[arg-type]

    adapter.execute(_request(runtime_input, config_snapshot=request_snapshot))

    admitted = recorder.requests[0]
    assert admitted.config_snapshot == request_snapshot
    assert admitted.options == expected


def test_removed_scalar_gui_language_shape_does_not_act_as_a_locale_alias(
    runtime_input: Path,
) -> None:
    recorder = _Recorder()
    adapter = RuntimePortAdapter(recorder)  # type: ignore[arg-type]

    adapter.execute(
        _request(
            runtime_input,
            config_snapshot={
                "image": {"ocr_language": "english"},
                "gui": {"language": "ja_JP"},
            },
        )
    )

    assert recorder.requests[0].options == {
        "ocr_language": "english",
        "locale": "zh_CN",
    }


@pytest.mark.parametrize(
    ("options", "expected"),
    [
        (
            {"ocr_language": "english"},
            {"ocr_language": "english", "locale": "ja_JP"},
        ),
        (
            {"locale": "en_US"},
            {"ocr_language": "japanese", "locale": "en_US"},
        ),
    ],
)
def test_explicit_nonblank_option_wins_while_missing_peer_uses_snapshot(
    runtime_input: Path,
    options: dict[str, Any],
    expected: dict[str, Any],
) -> None:
    recorder = _Recorder()
    adapter = RuntimePortAdapter(recorder)  # type: ignore[arg-type]

    adapter.execute(
        _request(
            runtime_input,
            options=options,
            config_snapshot=_snapshot(ocr_language="japanese", locale="ja_JP"),
        )
    )

    assert recorder.requests[0].options == expected


def test_present_falsey_options_are_preserved_by_projection(runtime_input: Path) -> None:
    recorder = _Recorder()
    adapter = RuntimePortAdapter(recorder)  # type: ignore[arg-type]
    options = {
        "ocr_language": "  ",
        "locale": None,
        "to_md_enable_ocr": False,
        "custom_sentinel": "",
    }

    adapter.execute(
        _request(
            runtime_input,
            options=options,
            config_snapshot=_snapshot(ocr_language="latin", locale="fr_FR"),
        )
    )

    assert recorder.requests[0].options == options


def test_explicit_nonblank_ocr_options_are_not_overwritten(runtime_input: Path) -> None:
    recorder = _Recorder()
    adapter = RuntimePortAdapter(recorder)  # type: ignore[arg-type]
    options = {"ocr_language": "cyrillic", "locale": "ru_RU"}

    adapter.execute(
        _request(
            runtime_input,
            options=options,
            config_snapshot=_snapshot(ocr_language="japanese", locale="ja_JP"),
        )
    )

    assert recorder.requests[0].options == options


@pytest.mark.parametrize("target_format", ["docx", "markdown", "MD", " md "])
def test_noncanonical_markdown_targets_do_not_receive_ocr_options(
    runtime_input: Path,
    target_format: str,
) -> None:
    recorder = _Recorder()
    adapter = RuntimePortAdapter(recorder)  # type: ignore[arg-type]

    adapter.execute(
        _request(
            runtime_input,
            config_snapshot=_snapshot(ocr_language="japanese", locale="ja_JP"),
            target_format=target_format,
        )
    )

    assert recorder.requests[0].options == {}


def test_markdown_numbering_action_does_not_receive_unconsumed_ocr_options(
    runtime_input: Path,
) -> None:
    recorder = _Recorder()
    adapter = RuntimePortAdapter(recorder)  # type: ignore[arg-type]

    adapter.execute(
        _request(
            runtime_input,
            config_snapshot=_snapshot(ocr_language="japanese", locale="ja_JP"),
            action_name="process_md_numbering",
        )
    )

    assert recorder.requests[0].options == {}


def test_request_without_any_snapshot_keeps_empty_options(runtime_input: Path) -> None:
    recorder = _Recorder()
    adapter = RuntimePortAdapter(recorder)  # type: ignore[arg-type]

    adapter.execute(_request(runtime_input))

    assert recorder.requests[0].options == {}

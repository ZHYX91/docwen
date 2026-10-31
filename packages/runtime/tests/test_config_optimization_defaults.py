"""Domain resets preserve other optimizer selections and concurrent edits."""

from __future__ import annotations

from contextlib import contextmanager
from pathlib import Path

import pytest

pytestmark = pytest.mark.unit

PROJECT_CONFIGS = Path(__file__).resolve().parents[3] / "configs"


def test_shipped_optimization_defaults_use_one_explicit_value(tmp_path: Path) -> None:
    from docwen_runtime.config.loader import ConfigLoader

    snapshot = ConfigLoader(base_dir=PROJECT_CONFIGS, user_dir=tmp_path / "user").config.as_dict()

    for category in ("document", "layout", "image"):
        assert snapshot[category]["to_md_optimization"] == "none"
        assert "to_md_enable_optimization" not in snapshot[category]
        assert "to_md_optimization_type" not in snapshot[category]


@pytest.mark.parametrize("selection", ["gongwen", "other_optimizer", "none"])
def test_gongwen_reset_owns_field_switch_and_only_its_selected_default(tmp_path: Path, selection: str) -> None:
    from docwen_runtime.config.loader import ConfigLoader

    user_dir = tmp_path / "user"
    loader = ConfigLoader(base_dir=PROJECT_CONFIGS, user_dir=user_dir)
    assert loader.set_values(
        {
            "template_fill.gongwen_fields_enabled": False,
            "template_fill.gongwen_numbering_text": False,
            "template_fill.gongwen_numbering_comments": False,
            "template_fill.attachment_alignment": "spaces",
            "template_fill.list_separator": " / ",
            "document.to_md_optimization": selection,
            "document.to_md_keep_images": False,
            "image.to_md_optimization": "invoice_cn",
            "layout.to_md_optimization": "other_optimizer",
        }
    )
    before = {name: (user_dir / f"{name}.toml").read_bytes() for name in ("document", "image", "layout")}

    assert loader.reset_group("gongwen") is True

    snapshot = loader.config.as_dict()
    expected_template_fill = {
        "gongwen_fields_enabled": True,
        "gongwen_numbering_text": True,
        "gongwen_numbering_comments": True,
        "attachment_alignment": "tab",
        "list_separator": " / ",
    }
    assert snapshot["template_fill"] == expected_template_fill
    assert (
        ConfigLoader(base_dir=PROJECT_CONFIGS, user_dir=user_dir).config.as_dict()["template_fill"]
        == expected_template_fill
    )
    assert snapshot["document"]["to_md_optimization"] == ("none" if selection == "gongwen" else selection)
    assert snapshot["document"]["to_md_keep_images"] is False
    for name in ("image", "layout"):
        assert (user_dir / f"{name}.toml").read_bytes() == before[name]
    if selection != "gongwen":
        assert (user_dir / "document.toml").read_bytes() == before["document"]


@pytest.mark.parametrize(
    ("image_selection", "layout_selection"),
    [
        ("invoice_cn", "invoice_cn"),
        ("invoice_cn", "other_optimizer"),
        ("other_optimizer", "invoice_cn"),
        ("none", "other_optimizer"),
    ],
)
def test_invoice_reset_independently_clears_only_invoice_defaults(
    tmp_path: Path, image_selection: str, layout_selection: str
) -> None:
    from docwen_runtime.config.loader import ConfigLoader

    user_dir = tmp_path / "user"
    loader = ConfigLoader(base_dir=PROJECT_CONFIGS, user_dir=user_dir)
    assert loader.set_values(
        {
            "image.to_md_optimization": image_selection,
            "image.to_md_keep_images": False,
            "layout.to_md_optimization": layout_selection,
            "layout.to_md_keep_images": False,
            "document.to_md_optimization": "gongwen",
            "template_fill.gongwen_fields_enabled": False,
        }
    )
    before = {
        name: (user_dir / f"{name}.toml").read_bytes() for name in ("document", "image", "layout", "template_fill")
    }

    assert loader.reset_group("invoice") is True

    snapshot = loader.config.as_dict()
    for category, selection in (("image", image_selection), ("layout", layout_selection)):
        assert snapshot[category]["to_md_optimization"] == ("none" if selection == "invoice_cn" else selection)
        assert snapshot[category]["to_md_keep_images"] is False
        if selection != "invoice_cn":
            assert (user_dir / f"{category}.toml").read_bytes() == before[category]
    for name in ("document", "template_fill"):
        assert (user_dir / f"{name}.toml").read_bytes() == before[name]


def test_conditional_only_reset_succeeds_without_creating_overrides(tmp_path: Path) -> None:
    from docwen_runtime.config.loader import ConfigLoader

    user_dir = tmp_path / "user"
    loader = ConfigLoader(base_dir=PROJECT_CONFIGS, user_dir=user_dir)

    assert loader.reset_group("invoice") is True
    assert not (user_dir / "image.toml").exists()
    assert not (user_dir / "layout.toml").exists()


@pytest.mark.parametrize("category", ["document", "layout", "image"])
def test_general_input_reset_preserves_domain_default(tmp_path: Path, category: str) -> None:
    from docwen_runtime.config.loader import ConfigLoader

    loader = ConfigLoader(base_dir=PROJECT_CONFIGS, user_dir=tmp_path / "user")
    assert loader.set_values(
        {
            f"{category}.to_md_keep_images": False,
            f"{category}.to_md_optimization": "other_optimizer",
        }
    )

    assert loader.reset_group(category) is True

    snapshot = loader.config.as_dict()
    assert snapshot[category]["to_md_keep_images"] is True
    assert snapshot[category]["to_md_optimization"] == "other_optimizer"


def test_text_reset_preserves_gongwen_field_switch(tmp_path: Path) -> None:
    from docwen_runtime.config.loader import ConfigLoader

    loader = ConfigLoader(base_dir=PROJECT_CONFIGS, user_dir=tmp_path / "user")
    assert loader.set_values(
        {
            "template_fill.list_separator": " / ",
            "template_fill.gongwen_fields_enabled": False,
            "template_fill.gongwen_numbering_text": False,
            "template_fill.gongwen_numbering_comments": False,
            "template_fill.attachment_alignment": "tab",
        }
    )

    assert loader.reset_group("text") is True

    expected_template_fill = {
        "list_separator": ", ",
        "gongwen_fields_enabled": False,
        "gongwen_numbering_text": False,
        "gongwen_numbering_comments": False,
        "attachment_alignment": "tab",
    }
    assert loader.config.as_dict()["template_fill"] == expected_template_fill
    assert (
        ConfigLoader(base_dir=PROJECT_CONFIGS, user_dir=tmp_path / "user").config.as_dict()["template_fill"]
        == expected_template_fill
    )


@pytest.mark.parametrize(
    ("group", "category", "owned_selection"),
    [("gongwen", "document", "gongwen"), ("invoice", "image", "invoice_cn"), ("invoice", "layout", "invoice_cn")],
)
@pytest.mark.parametrize("latest_is_owned", [True, False])
def test_domain_reset_rechecks_disk_after_entering_transaction_lock(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    group: str,
    category: str,
    owned_selection: str,
    latest_is_owned: bool,
) -> None:
    from docwen_runtime import file_transactions
    from docwen_runtime.config.loader import ConfigLoader

    user_dir = tmp_path / "user"
    first = ConfigLoader(base_dir=PROJECT_CONFIGS, user_dir=user_dir)
    second = ConfigLoader(base_dir=PROJECT_CONFIGS, user_dir=user_dir)
    initial = "other_optimizer" if latest_is_owned else owned_selection
    latest = owned_selection if latest_is_owned else "other_optimizer"
    key = f"{category}.to_md_optimization"
    assert first.set_value(key, initial)
    original_lock = file_transactions.process_config_lock
    interleaved = False

    @contextmanager
    def interleave_other_window(*args, **kwargs):
        nonlocal interleaved
        if not interleaved:
            interleaved = True
            assert second.set_value(key, latest)
            assert first.config.as_dict()[category]["to_md_optimization"] == initial
        with original_lock(*args, **kwargs):
            yield

    monkeypatch.setattr(file_transactions, "process_config_lock", interleave_other_window)

    assert first.reset_group(group) is True

    assert interleaved
    expected = "none" if latest_is_owned else "other_optimizer"
    assert first.config.as_dict()[category]["to_md_optimization"] == expected
    assert (
        ConfigLoader(base_dir=PROJECT_CONFIGS, user_dir=user_dir).config.as_dict()[category]["to_md_optimization"]
        == expected
    )


@pytest.mark.parametrize(
    ("persisted", "runtime", "expected"),
    [("gongwen", "other_optimizer", "none"), ("other_optimizer", "gongwen", "other_optimizer")],
)
def test_domain_reset_compares_persisted_selection_instead_of_runtime_override(
    tmp_path: Path, persisted: str, runtime: str, expected: str
) -> None:
    from docwen_runtime.config.loader import ConfigLoader

    user_dir = tmp_path / "user"
    writer = ConfigLoader(base_dir=PROJECT_CONFIGS, user_dir=user_dir)
    assert writer.set_value("document.to_md_optimization", persisted)
    loader = ConfigLoader(
        base_dir=PROJECT_CONFIGS,
        user_dir=user_dir,
        runtime_overrides={"document": {"to_md_optimization": runtime}},
    )

    assert loader.reset_group("gongwen") is True

    assert loader.config.as_dict()["document"]["to_md_optimization"] == runtime
    assert (
        ConfigLoader(base_dir=PROJECT_CONFIGS, user_dir=user_dir).config.as_dict()["document"]["to_md_optimization"]
        == expected
    )


@pytest.mark.parametrize("group", ["gongwen", "invoice"])
def test_domain_reset_rolls_back_all_changed_files_on_later_write_failure(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, group: str
) -> None:
    from docwen_runtime import toml_io
    from docwen_runtime.config.loader import ConfigLoader

    user_dir = tmp_path / "user"
    loader = ConfigLoader(base_dir=PROJECT_CONFIGS, user_dir=user_dir)
    assert loader.set_values(
        {
            "template_fill.gongwen_fields_enabled": False,
            "document.to_md_optimization": "gongwen",
            "image.to_md_optimization": "invoice_cn",
            "layout.to_md_optimization": "invoice_cn",
        }
    )
    before = {
        name: (user_dir / f"{name}.toml").read_bytes() for name in ("document", "image", "layout", "template_fill")
    }
    snapshot = loader.config.as_dict()
    original_save = toml_io.save_toml_document
    attempted = []

    def fail_second_write(path, document):
        attempted.append(Path(path))
        if len(attempted) == 2:
            raise OSError("simulated second domain reset write failure")
        return original_save(path, document)

    monkeypatch.setattr(toml_io, "save_toml_document", fail_second_write)

    assert loader.reset_group(group) is False

    assert len(attempted) == 2
    assert loader.config_state_trusted is True
    assert loader.config.as_dict() == snapshot
    for name, content in before.items():
        assert (user_dir / f"{name}.toml").read_bytes() == content

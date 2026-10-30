"""Optimization-page defaults, persistence and conditional reset ownership."""

from __future__ import annotations

from pathlib import Path

import pytest
from tests.support.gui_vm_fakes import optimization_capability_projection

from docwen_application.controller import ApplicationController
from docwen_bundle.config_port import ConfigPortAdapter
from docwen_gui.view_models.settings_vm import SECTION_TEXT, SettingsViewModel
from docwen_runtime.config import ConfigLoader

pytestmark = pytest.mark.unit
PROJECT_CONFIGS = Path(__file__).resolve().parents[4] / "configs"


@pytest.fixture
def composition(tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
    loader = ConfigLoader(base_dir=PROJECT_CONFIGS, user_dir=tmp_path / "configs")
    port = ConfigPortAdapter(loader)
    controller = ApplicationController(config_port=port)
    monkeypatch.setattr(controller, "describe_runtime_capabilities", optimization_capability_projection)
    vm = SettingsViewModel(controller=controller)
    vm.begin_session()
    return vm, port, controller


def test_numbering_comments_persist_cancel_and_reset_independently(composition) -> None:
    vm, port, controller = composition
    assert vm.config.text.gongwen_numbering_comments is False
    vm.set_field(SECTION_TEXT, "gongwen_numbering_comments", True)
    vm.set_field(SECTION_TEXT, "gongwen_fields_enabled", False)
    assert vm.apply_changes()
    reopened = SettingsViewModel(controller=controller)
    assert reopened.config.text.gongwen_numbering_comments is True
    assert reopened.config.text.gongwen_fields_enabled is False
    assert port.get("template_fill.gongwen_numbering_comments") is True
    vm.set_field(SECTION_TEXT, "gongwen_numbering_comments", False)
    vm.cancel()
    assert vm.config.text.gongwen_numbering_comments is True
    assert vm.reset_group("gongwen")
    assert vm.config.text.gongwen_numbering_comments is False


def test_one_default_per_input_category_persists_and_reopens(composition) -> None:
    vm, port, controller = composition
    assert vm.config.text.attachment_alignment == port.get("template_fill.attachment_alignment") == "tab"
    assert vm.get_default_optimization("document") == "none"
    assert port.set("document.to_md_optimization", "contract")
    vm.load_from_controller_config()
    vm.begin_session()

    assert vm.set_default_optimization("document", "gongwen")
    assert vm.set_default_optimization("layout", "invoice_cn")
    assert vm.get_default_optimization("document") == "gongwen"
    assert vm.get_default_optimization("layout") == "invoice_cn"
    assert vm.get_default_optimization("image") == "none"
    assert not vm.clear_default_optimization("document", "contract")
    assert vm.get_default_optimization("document") == "gongwen"
    assert port.get("document.to_md_optimization") == "contract"

    assert vm.apply_changes()
    reopened = SettingsViewModel(controller=controller)
    assert reopened.get_default_optimization("document") == "gongwen"
    assert reopened.get_default_optimization("layout") == "invoice_cn"
    assert reopened.get_default_optimization("image") == "none"
    assert not reopened.is_dirty


@pytest.mark.parametrize("alignment", ["spaces", "tab"])
def test_attachment_alignment_persists_and_reopens_when_field_formatting_is_disabled(
    composition, alignment: str
) -> None:
    vm, port, controller = composition
    vm.set_field(SECTION_TEXT, "attachment_alignment", alignment)
    vm.set_field(SECTION_TEXT, "gongwen_fields_enabled", False)

    assert vm.apply_changes()
    assert port.get("template_fill.attachment_alignment") == alignment
    reopened = SettingsViewModel(controller=controller)
    assert reopened.config.text.attachment_alignment == alignment
    assert reopened.config.text.gongwen_fields_enabled is False
    assert not reopened.is_dirty


def test_gongwen_reset_restores_attachment_alignment_without_resetting_text_drafts(composition) -> None:
    vm, port, _controller = composition
    vm.set_field(SECTION_TEXT, "attachment_alignment", "spaces")
    assert vm.apply_changes()
    vm.set_field(SECTION_TEXT, "list_separator", " | ")

    assert vm.reset_group("gongwen")
    assert vm.config.text.attachment_alignment == port.get("template_fill.attachment_alignment") == "tab"
    assert vm.config.text.list_separator == " | "
    vm.cancel_changes()
    assert vm.config.text.attachment_alignment == "tab"


def test_text_reset_preserves_attachment_alignment_and_its_draft(composition) -> None:
    vm, port, _controller = composition
    vm.set_field(SECTION_TEXT, "attachment_alignment", "tab")
    assert vm.apply_changes()
    vm.set_field(SECTION_TEXT, "attachment_alignment", "spaces")

    assert vm.reset_group("text")
    assert port.get("template_fill.attachment_alignment") == "tab"
    assert vm.config.text.attachment_alignment == "spaces"
    vm.cancel_changes()
    assert vm.config.text.attachment_alignment == "tab"


def test_default_changes_and_field_formatting_cancel_together_without_persisting(composition) -> None:
    vm, port, _controller = composition
    assert vm.set_default_optimization("document", "gongwen")
    assert vm.set_default_optimization("image", "invoice_cn")
    vm.set_field(SECTION_TEXT, "gongwen_fields_enabled", False)
    assert vm.is_dirty

    vm.cancel_changes()

    assert vm.get_default_optimization("document") == "none"
    assert vm.get_default_optimization("image") == "none"
    assert vm.config.text.gongwen_fields_enabled is True
    assert port.get("document.to_md_optimization") == "none"
    assert port.get("image.to_md_optimization") == "none"
    assert port.get("template_fill.gongwen_fields_enabled") is True
    assert not vm.is_dirty


def test_disabled_resource_can_be_cleared_but_cannot_be_selected(composition) -> None:
    vm, port, _controller = composition
    assert vm.set_default_optimization("document", "gongwen")
    assert vm.apply_changes()
    assert port.set("optimize.types.gongwen.enabled", False)

    assert not vm.set_default_optimization("document", "gongwen")
    assert not vm.set_default_optimization("document", "invoice_cn")
    assert not vm.set_default_optimization("document", "unavailable-resource")
    assert vm.get_default_optimization("document") == "gongwen"
    assert vm.clear_default_optimization("document", "gongwen")
    assert vm.get_default_optimization("document") == "none"
    assert vm.apply_changes()
    assert port.get("document.to_md_optimization") == "none"


@pytest.mark.parametrize(
    ("category", "group", "resource"),
    [("document", "gongwen", "gongwen"), ("layout", "invoice", "invoice_cn"), ("image", "invoice", "invoice_cn")],
)
@pytest.mark.parametrize(
    ("persisted", "draft", "expected_persisted", "expected_draft"),
    [
        ("own", "own", "none", "none"),
        ("own", "contract", "none", "contract"),
        ("contract", "own", "contract", "contract"),
        ("contract", "contract", "contract", "contract"),
        ("none", "own", "none", "none"),
        ("own", "none", "none", "none"),
        ("contract", "receipt", "contract", "receipt"),
    ],
)
def test_page_reset_preserves_other_optimization_drafts_and_uses_new_cancel_baseline(
    composition,
    category: str,
    group: str,
    resource: str,
    persisted: str,
    draft: str,
    expected_persisted: str,
    expected_draft: str,
) -> None:
    vm, port, _controller = composition
    persisted = resource if persisted == "own" else persisted
    draft = resource if draft == "own" else draft
    assert port.set(f"{category}.to_md_optimization", persisted)
    assert port.set("template_fill.gongwen_fields_enabled", False)
    vm.load_from_controller_config()
    vm.begin_session()
    vm.set_conversion_default(category, "to_md_optimization", draft)
    vm.set_field(SECTION_TEXT, "list_separator", " | ")
    vm.set_conversion_default(category, "to_md_keep_images", False)

    assert vm.reset_group(group)

    assert port.get(f"{category}.to_md_optimization") == expected_persisted
    assert vm.get_default_optimization(category) == expected_draft
    assert vm.config.text.list_separator == " | "
    assert getattr(vm.config.conversion_defaults, category)["to_md_keep_images"] is False
    assert vm.config.text.gongwen_fields_enabled is (group == "gongwen")
    assert vm.is_dirty
    vm.cancel_changes()
    assert vm.get_default_optimization(category) == expected_persisted
    assert not vm.is_dirty


@pytest.mark.parametrize(
    ("category", "resource"), [("document", "gongwen"), ("layout", "invoice_cn"), ("image", "invoice_cn")]
)
def test_ordinary_input_page_reset_preserves_optimization_default_and_draft(
    composition, category: str, resource: str
) -> None:
    vm, port, _controller = composition
    assert port.set(f"{category}.to_md_optimization", resource)
    assert port.set(f"{category}.to_md_keep_images", False)
    vm.load_from_controller_config()
    vm.begin_session()
    vm.set_conversion_default(category, "to_md_optimization", "contract")

    assert vm.reset_group(category)

    assert port.get(f"{category}.to_md_optimization") == resource
    assert vm.get_default_optimization(category) == "contract"
    assert getattr(vm.config.conversion_defaults, category)["to_md_keep_images"] is True
    assert vm.is_dirty


def test_invoice_reset_does_not_reset_gongwen_or_other_input_options(composition) -> None:
    vm, port, _controller = composition
    for category, resource in (("document", "gongwen"), ("layout", "invoice_cn"), ("image", "invoice_cn")):
        assert vm.set_default_optimization(category, resource)
    vm.set_field(SECTION_TEXT, "gongwen_fields_enabled", False)
    vm.set_conversion_default("image", "size_limit", 1024)
    vm.set_conversion_default("layout", "render_dpi", 600)
    assert vm.apply_changes()

    assert vm.reset_group("invoice")

    assert vm.get_default_optimization("image") == "none"
    assert vm.get_default_optimization("layout") == "none"
    assert vm.get_default_optimization("document") == "gongwen"
    assert vm.config.text.gongwen_fields_enabled is False
    assert port.get("image.size_limit") == 1024
    assert port.get("layout.render_dpi") == 600
    assert not vm.is_dirty


@pytest.mark.parametrize("outcome", ["success", "false", "raises"])
def test_page_reset_keeps_unrelated_drafts_after_another_writer_changes_disk(
    composition, tmp_path: Path, monkeypatch: pytest.MonkeyPatch, outcome: str
) -> None:
    vm, port, _controller = composition
    vm.set_field(SECTION_TEXT, "list_separator", " | ")
    vm.set_conversion_default("image", "to_md_optimization", "invoice_cn")
    other_writer = ConfigLoader(base_dir=PROJECT_CONFIGS, user_dir=tmp_path / "configs")
    assert other_writer.set_values({"template_fill.list_separator": " / ", "image.to_md_optimization": "receipt"})
    reset_group = port.reset_group

    def reset_with_reported_outcome(group: str) -> bool:
        result = reset_group(group)
        if outcome == "raises":
            raise RuntimeError("reset completion could not be acknowledged")
        return result if outcome == "success" else False

    monkeypatch.setattr(port, "reset_group", reset_with_reported_outcome)

    assert vm.reset_group("gongwen") is (outcome == "success")

    assert vm.config.text.list_separator == " | "
    assert vm.get_default_optimization("image") == "invoice_cn"
    assert port.get("template_fill.list_separator") == " / "
    assert port.get("image.to_md_optimization") == "receipt"
    assert vm.is_dirty
    vm.cancel_changes()
    assert vm.config.text.list_separator == " / "
    assert vm.get_default_optimization("image") == "receipt"
    assert not vm.is_dirty

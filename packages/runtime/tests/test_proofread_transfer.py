"""Rule transfer previews preserve complete sources and validate both choices."""

import pytest

from docwen_core.toml_tools import parse_toml_text
from docwen_runtime.config.proofread_transfer import plan_rule_import
from docwen_runtime.config.validation import ConfigSemanticError, validate_config_file

pytestmark = pytest.mark.unit


@pytest.mark.parametrize("name", ["symbol_map", "typos", "sensitive_words"])
def test_merge_keeps_existing_entries_values_and_comments(name):
    source = '# User rules\n[entries]\nright = ["old"] # existing note\nkeep = ["retained"]\n'
    incoming = '[entries]\nright = ["new", "old"] # imported note\nadded = ["new rule"]\n'
    plan = plan_rule_import(f"proofread/{name}.toml", source, incoming)
    assert plan.source_text == source
    assert plan.merge_text is not None
    merged = parse_toml_text(plan.merge_text)["entries"]
    assert merged == {"right": ["old", "new"], "keep": ["retained"], "added": ["new rule"]}
    assert "existing note; imported note" in plan.merge_text
    assert "User rules" in plan.merge_text
    assert "keep" not in parse_toml_text(plan.replace_text)["entries"]
    assert [change.kind for change in plan.changes] == ["conflict", "added", "existing_only"]


def test_conflicting_closing_symbol_disables_merge_but_allows_replacement():
    plan = plan_rule_import("proofread/pairs.toml", 'items = [["(", ")"]]\n', 'items = [["[", ")"]]\n')
    assert plan.merge_text is None
    assert plan.merge_error
    assert plan.changes[0].kind == "conflict"
    assert parse_toml_text(plan.replace_text)["items"] == [["[", ")"]]


def test_pair_merge_is_idempotent_and_retains_existing_order():
    source = 'items = [["(", ")"], ["[", "]"]]\n'
    incoming = 'items = [["[", "]"], ["{", "}"]]\n'
    plan = plan_rule_import("proofread/pairs.toml", source, incoming)
    assert plan.merge_text is not None
    again = plan_rule_import("proofread/pairs.toml", plan.merge_text, incoming)
    assert again.merge_text == plan.merge_text
    assert parse_toml_text(plan.merge_text)["items"] == [["(", ")"], ["[", "]"], ["{", "}"]]


@pytest.mark.parametrize(
    "text", ["entries = 4", '[entries]\nwrong = "text"', "unrelated = true", "[entries", '[entries]\n"" = ["bad"]']
)
def test_invalid_rule_import_is_rejected_without_mutation(text):
    source = '[entries]\nkeep = ["old"]\n'
    with pytest.raises(ValueError):
        plan_rule_import("proofread/typos.toml", source, text)
    assert parse_toml_text(source)["entries"] == {"keep": ["old"]}


def test_empty_replacement_is_explicit_and_merge_preserves_rules():
    source = '[entries]\nkeep = ["old"]\n'
    plan = plan_rule_import("proofread/typos.toml", source, "[entries]\n")
    assert parse_toml_text(plan.replace_text)["entries"] == {}
    assert plan.merge_text == source


@pytest.mark.parametrize(
    "filename,data",
    [("image.toml", {"ocr_language": "english"}), ("conversion.toml", {"md_to_docx": {"list_separator": ", "}})],
)
def test_removed_config_paths_are_rejected(filename, data):
    with pytest.raises(ConfigSemanticError, match="has been removed"):
        validate_config_file(filename, data, {})


@pytest.mark.parametrize(
    "language", ["auto", "chinese", "chinese_cht", "english", "japanese", "korean", "latin", "cyrillic"]
)
def test_shared_ocr_language_validation(language):
    assert validate_config_file("ocr.toml", {"language": language}, {}) == {"language": language}


@pytest.mark.parametrize("value", [1, True, "unknown", " english "])
def test_shared_ocr_rejects_invalid_values(value):
    with pytest.raises(ConfigSemanticError):
        validate_config_file("ocr.toml", {"language": value}, {})


@pytest.mark.parametrize("separator", ["", " ", ", ", "、"])
def test_template_separator_remains_exact(separator):
    assert validate_config_file("template_fill.toml", {"list_separator": separator}, {})["list_separator"] == separator


@pytest.mark.parametrize("reverse", [False, True])
def test_pair_merge_accepts_both_table_and_array_sources(reverse):
    table = '# Pairs\n[[items]] # paired table\nsource = "("\ntarget = ")" # round pair\n'
    array = 'items = [["[", "]"]] # square pair\n'
    source, incoming = (array, table) if reverse else (table, array)
    plan = plan_rule_import("proofread/pairs.toml", source, incoming)
    assert plan.merge_text is not None
    assert {tuple(pair) for pair in parse_toml_text(plan.merge_text)["items"]} == {("(", ")"), ("[", "]")}
    assert "round pair" in plan.merge_text
    assert "paired table" in plan.merge_text
    assert "square pair" in plan.merge_text
    assert plan_rule_import("proofread/pairs.toml", plan.merge_text, incoming).merge_text == plan.merge_text


def test_dictionary_merge_retains_inner_comments_and_is_idempotent():
    source = '[entries]\nword = [\n  "old#literal", # old inner note\n] # old outer note\n'
    incoming = '[entries]\nword = [\n  "new#literal", # new inner note\n  "new#literal",\n] # new outer note\n'
    plan = plan_rule_import("proofread/typos.toml", source, incoming)
    assert plan.merge_text is not None
    assert parse_toml_text(plan.merge_text)["entries"]["word"] == ["old#literal", "new#literal"]
    for note in ("old inner note", "old outer note", "new inner note", "new outer note"):
        assert note in plan.merge_text
    assert "#literal" not in plan.changes[0].incoming.split("\n# ")[1]
    assert plan_rule_import("proofread/typos.toml", plan.merge_text, incoming).merge_text == plan.merge_text


def test_imported_standalone_notes_appear_in_preview_and_survive_merge():
    source = '[entries]\nkeep = ["old"]\n'
    incoming = '[entries] # imported dictionary\n# standalone note\nadded = ["new"]\n# end note\n'
    plan = plan_rule_import("proofread/typos.toml", source, incoming)
    assert "standalone note" in plan.changes[0].incoming
    assert plan.merge_text is not None
    for note in ("standalone note", "imported dictionary", "end note"):
        assert note in plan.merge_text
    assert plan_rule_import("proofread/typos.toml", plan.merge_text, incoming).merge_text == plan.merge_text

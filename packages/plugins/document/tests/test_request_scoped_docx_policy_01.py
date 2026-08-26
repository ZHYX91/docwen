"""Focused tests split from test_request_scoped_docx_policy.py."""

from __future__ import annotations

from ._request_scoped_docx_policy_support import (
    Any,
    Barrier,
    Document,
    DocumentPlugin,
    DocxToMarkdownConverter,
    Event,
    Path,
    ThreadPoolExecutor,
    _build_concurrency_probe_docx,
    _build_formula_probe_docx,
    _build_malformed_note_probe_docx,
    _build_policy_probe_docx,
    _build_rich_page_break_probe_docx,
    _context,
    _markdown_from_result,
    _ParagraphStyleProbe,
    _request_policy,
    build_docx_markdown_request_policy,
    pytest,
)

pytestmark = pytest.mark.unit


def test_request_style_options_merge_with_document_detector_fields(tmp_path: Path) -> None:
    from docwen_core.docx_parsing.format_features import detect_paragraph_style_type

    input_path = tmp_path / "style-policy.docx"
    Document().save(str(input_path))
    config = {
        "document": {
            "style": {
                "code": {
                    "docx_to_md": {
                        "paragraph_style_aliases": ["DocumentCode", "documentcode"],
                        "character_style_aliases": ["DocumentCodeChar"],
                        "full_paragraph_as_block": False,
                        "fuzzy_match_enabled": False,
                        "shading": {"wps_enabled": False, "word_enabled": False},
                    }
                },
                "quote": {
                    "docx_to_md": {
                        "level_style_aliases": {"Document Quote": 3},
                        "paragraph_style_aliases": ["Document Aside", "document aside"],
                        "character_style_aliases": ["DocumentQuoteChar"],
                        "full_paragraph_as_block": False,
                        "fuzzy_match_enabled": False,
                    }
                },
            }
        }
    }
    context = _context(
        tmp_path,
        input_path,
        request_id="style-policy-merge",
        config=config,
        options={
            "code_block_style_aliases": ["Acme Snippet", "acme snippet"],
            "quote_style_aliases": ["Request Quote", "REQUEST QUOTE"],
            "quote_generic_names": ["Request Aside", "request aside"],
        },
    )

    detector = build_docx_markdown_request_policy(context, context.request.options).style_detector

    assert detector is not None
    assert detector.code_block_style_names == frozenset({"DocumentCode"})
    assert detector.code_character_style_names == frozenset({"DocumentCodeChar"})
    assert detector.code_full_paragraph_as_block is False
    assert detector.wps_shading_enabled is False
    assert detector.word_shading_enabled is False
    assert detector.code_fuzzy_match_enabled is False
    assert detector.code_block_style_fragments == ("Acme Snippet",)
    assert detector.quote_character_style_names == frozenset({"DocumentQuoteChar"})
    assert detector.quote_full_paragraph_as_block is False
    assert detector.quote_fuzzy_match_enabled is False
    assert detector.quote_style_names == (("Document Quote", 3),)
    assert detector.quote_style_patterns == (("Request Quote", 1),)
    assert detector.quote_generic_names == frozenset({"Document Aside", "Request Aside"})
    assert detect_paragraph_style_type(_ParagraphStyleProbe("My Acme Snippet Style"), detector) == (None, None)


def test_sequential_request_style_aliases_never_require_cache_invalidation(tmp_path: Path) -> None:
    """Each admitted snapshot owns aliases; a later request cannot see stale names."""
    from docwen_core.docx_parsing.format_features import detect_paragraph_style_type

    input_path = tmp_path / "style-alias-snapshot.docx"
    Document().save(str(input_path))

    def config_for(alias: str) -> dict[str, Any]:
        return {
            "document": {
                "style": {
                    "code": {
                        "docx_to_md": {
                            "paragraph_style_aliases": [alias],
                            "fuzzy_match_enabled": False,
                        }
                    }
                }
            }
        }

    first_context = _context(
        tmp_path,
        input_path,
        request_id="style-alias-first",
        config=config_for("First Request Code"),
    )
    second_context = _context(
        tmp_path,
        input_path,
        request_id="style-alias-second",
        config=config_for("Second Request Code"),
    )

    first = build_docx_markdown_request_policy(first_context, {}).style_detector
    second = build_docx_markdown_request_policy(second_context, {}).style_detector

    assert first is not None
    assert second is not None
    first_probe = _ParagraphStyleProbe("First Request Code")
    second_probe = _ParagraphStyleProbe("Second Request Code")
    assert detect_paragraph_style_type(first_probe, config=first) == ("code_block", True)
    assert detect_paragraph_style_type(second_probe, config=first) == (None, None)
    assert detect_paragraph_style_type(second_probe, config=second) == ("code_block", True)
    assert detect_paragraph_style_type(first_probe, config=second) == (None, None)


@pytest.mark.parametrize(
    ("note_type", "footnote_loss_count", "endnote_loss_count"),
    [("footnote", 1, 0), ("endnote", 0, 1)],
)
def test_malformed_referenced_note_part_is_a_typed_visible_loss(
    tmp_path: Path,
    note_type: str,
    footnote_loss_count: int,
    endnote_loss_count: int,
) -> None:
    input_path = _build_malformed_note_probe_docx(tmp_path, note_type)
    context = _context(
        tmp_path,
        input_path,
        request_id=f"malformed-{note_type}",
        config=_request_policy(),
    )

    result = DocumentPlugin().convert(context)

    assert result.success is True
    markdown = _markdown_from_result(result)
    assert "body note[^" in markdown
    assert "[^1]:" not in markdown
    warning = next(diagnostic for diagnostic in result.diagnostics if diagnostic.code == "DOCX2MD-NOTE-DEFINITION-LOSS")
    assert warning.level == "warning"
    assert f"footnotes={footnote_loss_count}" in warning.message
    assert f"endnotes={endnote_loss_count}" in warning.message
    assert result.metrics.extra["note_definition_loss_count"] == 1
    primary = next(artifact for artifact in result.artifacts if artifact.is_primary)
    assert primary.metadata["note_definition_loss_count"] == 1


def test_context_snapshot_owns_all_docx_policy(tmp_path: Path) -> None:
    """Formatting, syntax, aliases and export semantics are request-scoped."""
    input_path = _build_policy_probe_docx(tmp_path)
    context = _context(
        tmp_path,
        input_path,
        request_id="request-policy",
        config=_request_policy(),
    )

    markdown = _markdown_from_result(DocumentPlugin().convert(context))

    expected_fragments = {
        "body formatting and inline syntax": "__request bold__ and _request italic_",
        "heading formatting and inline syntax": "## __request heading__",
        "nested list marker and indentation": "  + nested request item",
        "request page-break separator": "\n___\n",
        "request code style alias": "```\nprint('request policy')\n```",
        "request quote style alias level": ">> request aside",
        "formatted table header": "| __Merged Header__ | Top |",
        "empty merged-cell export strategy": "|  | Bottom |",
        "request image export mode": "<!-- image omitted:",
        "inline footnote reference": "note anchor[^1]",
        "footnote definition": "[^1]: request note",
    }
    missing = [label for label, fragment in expected_fragments.items() if fragment not in markdown]
    assert missing == [], f"request policy was not applied for: {missing}\n\n{markdown}"
    assert sum(artifact.kind == "image" for artifact in context.workspace.registered_artifacts) == 1


def test_warmed_document_plugin_keeps_parallel_request_policies_isolated(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """A warmed shared plugin must keep per-request converter state isolated."""
    plugin = DocumentPlugin()
    warm_path = _build_concurrency_probe_docx(tmp_path, name="warm.docx")
    warm_context = _context(
        tmp_path,
        warm_path,
        request_id="warm",
        config={},
        options={"preserve_formatting": True, "page_break_separator": "---"},
    )
    assert plugin.convert(warm_context).success

    parse_barrier = Barrier(2)
    original_parse = DocxToMarkdownConverter._parse_docx  # pyright: ignore[reportPrivateUsage]

    def _barrier_parse(self: DocxToMarkdownConverter, *args: Any, **kwargs: Any) -> Any:
        parse_barrier.wait(timeout=10)
        return original_parse(self, *args, **kwargs)

    monkeypatch.setattr(DocxToMarkdownConverter, "_parse_docx", _barrier_parse)

    path_a = _build_concurrency_probe_docx(
        tmp_path,
        name="thread-a.docx",
        note_text="thread A note",
    )
    path_b = _build_concurrency_probe_docx(
        tmp_path,
        name="thread-b.docx",
        note_text="thread B note",
    )
    context_a = _context(
        tmp_path,
        path_a,
        request_id="thread-a",
        config=_request_policy(preserve_formatting=True, page_break="***"),
        options={"preserve_formatting": True, "page_break_separator": "***"},
    )
    context_b = _context(
        tmp_path,
        path_b,
        request_id="thread-b",
        config=_request_policy(preserve_formatting=False, page_break="___"),
        options={"preserve_formatting": False, "page_break_separator": "___"},
    )

    with ThreadPoolExecutor(max_workers=2) as executor:
        future_a = executor.submit(plugin.convert, context_a)
        future_b = executor.submit(plugin.convert, context_b)
        markdown_a = _markdown_from_result(future_a.result(timeout=20))
        markdown_b = _markdown_from_result(future_b.result(timeout=20))

    assert "__thread bold__" in markdown_a
    assert "\n***\n" in markdown_a
    assert "__thread bold__" not in markdown_b
    assert "thread bold" in markdown_b
    assert "\n___\n" in markdown_b
    assert "thread note[^1]" in markdown_a
    assert "[^1]: thread A note" in markdown_a
    assert "thread B note" not in markdown_a
    assert "thread note[^1]" in markdown_b
    assert "[^1]: thread B note" in markdown_b
    assert "thread A note" not in markdown_b


def test_shared_converter_serializes_requests_to_protect_instance_policy(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """One manually shared converter must not overlap request-owned state."""
    converter = DocxToMarkdownConverter()
    path_a = _build_concurrency_probe_docx(tmp_path, name="shared-a.docx")
    path_b = _build_concurrency_probe_docx(tmp_path, name="shared-b.docx")
    context_a = _context(
        tmp_path,
        path_a,
        request_id="shared-a",
        config=_request_policy(preserve_formatting=True, page_break="---"),
    )
    context_b = _context(
        tmp_path,
        path_b,
        request_id="shared-b",
        config=_request_policy(preserve_formatting=False, page_break="___"),
    )

    request_a_parsing = Event()
    request_b_parsing = Event()
    release_request_a = Event()
    original_parse = DocxToMarkdownConverter._parse_docx  # pyright: ignore[reportPrivateUsage]

    def _controlled_parse(self: DocxToMarkdownConverter, *args: Any, **kwargs: Any) -> Any:
        context = kwargs["context"]
        if context.request.request_id == "shared-a":
            request_a_parsing.set()
            assert release_request_a.wait(timeout=5.0)
        else:
            request_b_parsing.set()
        return original_parse(self, *args, **kwargs)

    monkeypatch.setattr(DocxToMarkdownConverter, "_parse_docx", _controlled_parse)

    with ThreadPoolExecutor(max_workers=2) as executor:
        future_a = executor.submit(converter.convert, context_a)
        assert request_a_parsing.wait(timeout=1.0)
        future_b = executor.submit(converter.convert, context_b)
        overlapped = request_b_parsing.wait(timeout=0.5)
        release_request_a.set()
        markdown_a = _markdown_from_result(future_a.result(timeout=10))
        markdown_b = _markdown_from_result(future_b.result(timeout=10))

    assert not overlapped
    assert "__thread bold__" in markdown_a
    assert "thread head\n\n---\n\nthread tail" in markdown_a
    assert "__thread bold__" not in markdown_b
    assert "thread head\n\n___\n\nthread tail" in markdown_b


def test_captured_request_snapshot_is_stable_until_execution(tmp_path: Path) -> None:
    """A captured context retains its policy until plugin execution."""
    input_path = _build_concurrency_probe_docx(tmp_path, name="snapshot-a.docx")
    context_a = _context(
        tmp_path,
        input_path,
        request_id="snapshot-a",
        config=_request_policy(preserve_formatting=True, page_break="___"),
    )

    markdown = _markdown_from_result(DocumentPlugin().convert(context_a))
    assert "__thread bold__" in markdown
    assert "\n___\n" in markdown
    assert "**thread bold**" not in markdown
    assert "thread head\n---\n\nthread tail" not in markdown


def test_page_break_preserves_request_formatting_and_note_reference(tmp_path: Path) -> None:
    input_path = _build_rich_page_break_probe_docx(tmp_path, name="rich-page-break.docx")
    context = _context(
        tmp_path,
        input_path,
        request_id="rich-page-break",
        config=_request_policy(preserve_formatting=True, page_break="___"),
    )

    markdown = _markdown_from_result(DocumentPlugin().convert(context))

    assert "__rich before__\n\n___\n\n_rich after_[^1]" in markdown
    assert "[^1]: page break note" in markdown
    assert "rich before\n___\n\nrich after" not in markdown


def test_export_owner_ignores_removed_document_mode_duplicates(tmp_path: Path) -> None:
    """Removed document-section duplicates cannot override request export policy."""
    input_path = _build_concurrency_probe_docx(tmp_path, name="owner-priority.docx")
    config = _request_policy()
    config["export"] = {
        "to_md_image_extraction_mode": "omit",
        "to_md_ocr_placement_mode": "image_md",
    }
    config["document"]["to_md_image_extraction_mode"] = "file"
    config["document"]["to_md_ocr_placement_mode"] = "main_md"
    context = _context(
        tmp_path,
        input_path,
        request_id="owner-priority",
        config=config,
    )

    policy = build_docx_markdown_request_policy(context, {})
    modes = policy.resolve_export_modes()

    assert modes["image_extraction_mode"] == "omit"
    assert modes["ocr_placement_mode"] == "image_md"


def test_document_table_merge_owner_precedes_conversion_compatibility_value(tmp_path: Path) -> None:
    input_path = _build_concurrency_probe_docx(tmp_path, name="table-owner.docx")
    config = _request_policy()
    config["conversion"]["table_merge_export_strategy"] = "fill"
    config["document"]["to_md_table_merge_export_strategy"] = "empty"
    context = _context(
        tmp_path,
        input_path,
        request_id="table-owner",
        config=config,
    )

    policy = build_docx_markdown_request_policy(context, {})

    assert policy.resolve_export_modes()["table_merge_export_strategy"] == "empty"


def test_blank_mode_options_preserve_request_snapshot_modes(tmp_path: Path) -> None:
    input_path = _build_concurrency_probe_docx(tmp_path, name="blank-mode-options.docx")
    config = _request_policy()
    config["export"]["to_md_image_extraction_mode"] = "file"
    config["export"]["to_md_ocr_placement_mode"] = "image_md"
    context = _context(
        tmp_path,
        input_path,
        request_id="blank-mode-options",
        config=config,
    )

    policy = build_docx_markdown_request_policy(
        context,
        {"image_mode": "", "ocr_placement": ""},
    )

    assert policy.resolve_export_modes() == {
        "image_extraction_mode": "file",
        "ocr_placement_mode": "image_md",
        "table_merge_export_strategy": "empty",
    }


def test_request_policy_freezes_export_options_at_projection_time(tmp_path: Path) -> None:
    input_path = _build_concurrency_probe_docx(tmp_path, name="frozen-options.docx")
    context = _context(
        tmp_path,
        input_path,
        request_id="frozen-options",
        config=_request_policy(),
    )
    options: dict[str, Any] = {
        "image_mode": "file",
        "ocr_placement": "image_md",
        "table_merge_strategy": "empty",
        "image_link_style": "markdown_link",
    }

    policy = build_docx_markdown_request_policy(context, options)
    options.update(
        {
            "image_mode": "omit",
            "ocr_placement": "main_md",
            "table_merge_strategy": "fill",
            "image_link_style": "wiki_embed",
        }
    )

    assert policy.resolve_export_modes() == {
        "image_extraction_mode": "file",
        "ocr_placement_mode": "image_md",
        "table_merge_export_strategy": "empty",
    }
    assert policy.image_link_style == "markdown_link"


def test_partial_nonempty_snapshot_uses_deterministic_defaults(tmp_path: Path) -> None:
    """A partial captured snapshot fills missing values from pure defaults."""
    input_path = _build_concurrency_probe_docx(tmp_path, name="partial-snapshot.docx")
    context = _context(
        tmp_path,
        input_path,
        request_id="partial-snapshot",
        config={"conversion": {}},
    )

    markdown = _markdown_from_result(DocumentPlugin().convert(context))

    assert "**thread bold**" in markdown
    assert "__thread bold__" not in markdown
    assert "thread head\n\n---\n\nthread tail" in markdown
    assert "thread head\n***\n\nthread tail" not in markdown


def test_empty_snapshot_uses_deterministic_defaults(tmp_path: Path) -> None:
    input_path = _build_concurrency_probe_docx(tmp_path, name="empty-snapshot.docx")
    context = _context(
        tmp_path,
        input_path,
        request_id="empty-snapshot",
        config={},
    )

    markdown = _markdown_from_result(DocumentPlugin().convert(context))

    assert "**thread bold**" in markdown
    assert "__thread bold__" not in markdown
    assert "thread head\n\n---\n\nthread tail" in markdown
    assert "thread head\n___\n\nthread tail" not in markdown


def test_explicit_options_preserve_false_ignore_and_omit_presentation(tmp_path: Path) -> None:
    input_path = _build_policy_probe_docx(tmp_path, name="explicit-options.docx")
    config = _request_policy()
    config["export"]["to_md_image_extraction_mode"] = "file"
    context = _context(
        tmp_path,
        input_path,
        request_id="explicit-options",
        config=config,
        options={
            "preserve_formatting": False,
            "page_break_separator": "ignore",
            "image_mode": "omit",
        },
    )

    markdown = _markdown_from_result(DocumentPlugin().convert(context))

    assert "request bold and request italic" in markdown
    assert "__request bold__" not in markdown
    assert "before request break\nafter request break" in markdown
    assert "before request break\n---\n\nafter request break" not in markdown
    assert "before request break\n***\n\nafter request break" not in markdown
    assert "before request break\n___\n\nafter request break" not in markdown
    assert "ignore" not in markdown
    assert "<!-- image omitted:" in markdown
    assert sum(artifact.kind == "image" for artifact in context.workspace.registered_artifacts) == 1


def test_formula_text_uses_request_syntax_instead_of_hardcoded_markers(tmp_path: Path) -> None:
    input_path = _build_formula_probe_docx(tmp_path)
    context = _context(
        tmp_path,
        input_path,
        request_id="formula-policy",
        config=_request_policy(),
    )

    markdown = _markdown_from_result(DocumentPlugin().convert(context))

    assert "__formula bold__ $x$ _formula italic_" in markdown
    assert "**formula bold**" not in markdown


def test_formula_text_respects_explicit_preserve_formatting_false(tmp_path: Path) -> None:
    input_path = _build_formula_probe_docx(tmp_path)
    context = _context(
        tmp_path,
        input_path,
        request_id="formula-plain",
        config=_request_policy(),
        options={"preserve_formatting": False},
    )

    markdown = _markdown_from_result(DocumentPlugin().convert(context))

    assert "formula bold $x$ formula italic" in markdown
    assert "__formula bold__" not in markdown
    assert "_formula italic_" not in markdown

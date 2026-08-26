"""Focused tests split from test_md_to_docx_notes.py."""

from __future__ import annotations

from ._md_to_docx_notes_support import (
    MD_NO_NOTES,
    MD_WITH_FOOTNOTES,
    WML_NS,
    Any,
    NoteContext,
    NoteWritebackError,
    _create_endnote_element,
    _create_endnote_ref_run,
    _create_footnote_element,
    _create_footnote_ref_run,
    _mk_children,
    _q,
    normalize_note_syntax,
    process_md_body_with_notes,
    pytest,
)

pytestmark = pytest.mark.contract


class TestProcessBodyFoundation:
    """F-F3-007: process_md_body_with_notes extracts definitions correctly."""

    def test_foundation_extracts_footnotes(self):
        """Footnotes are extracted and removed from the AST."""
        cleaned_ast, note_ctx = process_md_body_with_notes(MD_WITH_FOOTNOTES)

        assert "1" in note_ctx._footnote_children
        assert "2" in note_ctx._footnote_children
        assert "3" in note_ctx._footnote_children

        # AST must not contain 'footnotes' node
        ast_types = {n.get("type") for n in cleaned_ast}
        assert "footnotes" not in ast_types
        assert "footnote_item" not in ast_types

        # AST must still contain footnote_ref nodes
        ref_count = 0

        def _count_refs(nodes):
            nonlocal ref_count
            for n in nodes:
                if n.get("type") == "footnote_ref":
                    ref_count += 1
                _count_refs(n.get("children", []))

        _count_refs(cleaned_ast)
        assert ref_count == 4  # [^1], [^2], [^3], [^endnote:x]

    def test_foundation_extracts_endnotes(self):
        """Endnotes are extracted separately from footnotes."""
        _, note_ctx = process_md_body_with_notes(MD_WITH_FOOTNOTES)

        assert note_ctx.has_endnotes
        # The request-local projection maps canonical syntax to Mistune's ENDNOTE-X key.
        assert "X" in note_ctx._endnote_children
        # Endnote key has the ENDNOTE- prefix stripped
        assert "ENDNOTE-X" not in note_ctx._endnote_children

    def test_foundation_no_notes_clean_passthrough(self):
        """Markdown with no notes produces an empty NoteContext."""
        cleaned_ast, note_ctx = process_md_body_with_notes(MD_NO_NOTES)

        assert not note_ctx.has_notes
        assert note_ctx._footnote_children == {}
        assert note_ctx._endnote_children == {}
        # AST should be unchanged (no 'footnotes' node was present)
        ast_types = {n.get("type") for n in cleaned_ast}
        assert "footnotes" not in ast_types

    def test_foundation_preserves_other_ast_nodes(self):
        """Non-footnote AST nodes pass through unchanged."""
        cleaned_ast, _ = process_md_body_with_notes(MD_WITH_FOOTNOTES)

        headings = [n for n in cleaned_ast if n.get("type") == "heading"]
        assert len(headings) >= 2
        paragraphs = [n for n in cleaned_ast if n.get("type") == "paragraph"]
        assert len(paragraphs) >= 3

    def test_foundation_footnote_content_with_formatting(self):
        """Footnote definitions preserve inline formatting AST structure."""
        _, note_ctx = process_md_body_with_notes(MD_WITH_FOOTNOTES)

        # The **bold** text should have its AST structure preserved
        children = note_ctx._footnote_children["2"]
        assert len(children) == 1  # single paragraph
        # Should contain a 'strong' node
        types = {c.get("type") for c in children[0]}
        assert "strong" in types

    def test_cross_project_canonical_and_explicit_syntax(self):
        markdown = """Foot[^plain] explicit[^footnote:explicit] end[^endnote:tail].

[^plain]: Plain footnote.
[^footnote:explicit]: Explicit footnote.
[^endnote:tail]: Canonical endnote.
"""

        cleaned_ast, note_ctx = process_md_body_with_notes(markdown)

        assert len(note_ctx._footnote_children) == 2
        assert len(note_ctx._endnote_children) == 1
        assert (
            sum(child.get("type") == "footnote_ref" for node in cleaned_ast for child in node.get("children", [])) == 3
        )

    def test_retired_endnote_syntax_is_rejected(self):
        markdown = "Retired[^endnote-old].\n\n[^endnote-old]: Retired body.\n"

        with pytest.raises(NoteWritebackError) as raised:
            process_md_body_with_notes(markdown)

        assert raised.value.diagnostic_code == "MD2DOCX-NOTE-SYNTAX-INVALID"
        assert "use the current endnote form" in str(raised.value)

    @pytest.mark.parametrize(
        ("markdown", "message"),
        [
            ("Missing[^x].\n", "Missing footnote definition"),
            ("A[^x].\n\n[^x]: one\n[^x]: two\n", "Duplicate note definition"),
            (
                "A[^x] B[^footnote:x].\n\n[^x]: one\n[^footnote:x]: two\n",
                "Default and explicit footnote definitions",
            ),
        ],
    )
    def test_note_syntax_failures_are_closed(self, markdown: str, message: str):
        with pytest.raises(NoteWritebackError) as raised:
            process_md_body_with_notes(markdown)

        assert raised.value.diagnostic_code == "MD2DOCX-NOTE-SYNTAX-INVALID"
        assert message in str(raised.value)

    def test_two_space_and_tab_continuations_stay_inside_note(self):
        markdown = """Body[^n].

[^n]: first
  ## not a document heading
  Table: not a document caption {#tbl-hidden}
	@[[fig:hidden]] is not a document reference
"""

        projection = normalize_note_syntax(markdown)
        cleaned_ast, note_ctx = process_md_body_with_notes(markdown)

        assert "    ## not a document heading" in projection
        assert "    @[[fig:hidden]]" in projection
        assert all(node.get("type") != "heading" for node in cleaned_ast)
        assert len(note_ctx._footnote_children) == 1


class TestNoteContextFoundation:
    """F-F1-022: NoteContext manages ID mapping and element lifecycle."""

    def test_foundation_empty_context(self):
        """Empty NoteContext has no notes."""
        ctx = NoteContext()
        assert not ctx.has_notes
        assert not ctx.has_footnotes
        assert not ctx.has_endnotes

    def test_foundation_footnote_id_allocation(self):
        """Word IDs are allocated sequentially for footnotes."""
        fns: dict[str, Any] = {
            "a": _mk_children("Content A"),
            "b": _mk_children("Content B"),
        }
        ctx = NoteContext()
        ctx._footnote_children = fns

        id_a = ctx.get_footnote_word_id("a")
        id_b = ctx.get_footnote_word_id("b")

        assert id_a == 1
        assert id_b == 2
        assert len(ctx._footnote_id_map) == 2
        assert len(ctx.footnote_elements) == 2

    def test_foundation_footnote_id_stable(self):
        """Repeated lookups return the same Word ID."""
        fns: dict[str, Any] = {"x": _mk_children("Content X")}
        ctx = NoteContext()
        ctx._footnote_children = fns

        id1 = ctx.get_footnote_word_id("x")
        id2 = ctx.get_footnote_word_id("x")

        assert id1 == id2 == 1
        # Only one element should be created
        assert len(ctx.footnote_elements) == 1

    def test_foundation_undefined_footnote_returns_none(self):
        """Querying an undefined footnote returns None."""
        ctx = NoteContext()
        ctx._footnote_children = {"a": _mk_children("A")}

        assert ctx.get_footnote_word_id("nonexistent") is None

    def test_foundation_endnote_id_allocation(self):
        """Word IDs are allocated sequentially for endnotes."""
        ens: dict[str, list[list[dict]]] = {
            "1": _mk_children("End A"),
            "2": _mk_children("End B"),
        }
        ctx = NoteContext()
        ctx._endnote_children = ens

        id_1 = ctx.get_endnote_word_id("1")
        id_2 = ctx.get_endnote_word_id("2")

        assert id_1 == 1
        assert id_2 == 2
        assert len(ctx.endnote_elements) == 2

    def test_foundation_endnote_with_prefix(self):
        """get_endnote_word_id accepts keys with or without prefix."""
        ens: dict[str, list[list[dict]]] = {"x": _mk_children("End X")}
        ctx = NoteContext()
        ctx._endnote_children = ens

        id1 = ctx.get_endnote_word_id("x")
        id2 = ctx.get_endnote_word_id("ENDNOTE-x")
        id3 = ctx.get_endnote_word_id("endnote-x")

        assert id1 == 1
        assert id2 == 1
        assert id3 == 1
        # Only one element created
        assert len(ctx.endnote_elements) == 1

    def test_foundation_create_footnote_ref_run(self):
        """create_footnote_ref_run returns an OOXML run with footnoteReference."""
        ctx = NoteContext()
        ctx._footnote_children = {"f1": _mk_children("Content")}
        run = ctx.create_footnote_ref_run("f1")

        assert run is not None
        assert run.tag == _q("r")
        # Check for footnoteReference child
        refs = run.findall(f".//{{{WML_NS}}}footnoteReference")
        assert len(refs) == 1
        assert refs[0].get(_q("id")) == "1"

    def test_foundation_create_footnote_ref_run_undefined(self):
        """create_footnote_ref_run returns None for undefined notes."""
        ctx = NoteContext()
        assert ctx.create_footnote_ref_run("nope") is None

    def test_foundation_create_endnote_ref_run(self):
        """create_endnote_ref_run returns an OOXML run with endnoteReference."""
        ctx = NoteContext()
        ctx._endnote_children = {"e1": _mk_children("End")}
        run = ctx.create_endnote_ref_run("e1")

        assert run is not None
        refs = run.findall(f".//{{{WML_NS}}}endnoteReference")
        assert len(refs) == 1
        assert refs[0].get(_q("id")) == "1"

    def test_foundation_has_notes_properties(self):
        """has_footnotes / has_endnotes / has_notes reflect state."""
        empty = NoteContext()
        assert not empty.has_notes
        assert not empty.has_footnotes
        assert not empty.has_endnotes

        fn = NoteContext()
        fn._footnote_children = {"a": _mk_children("A")}
        assert fn.has_notes
        assert fn.has_footnotes
        assert not fn.has_endnotes

        en = NoteContext()
        en._endnote_children = {"b": _mk_children("B")}
        assert en.has_notes
        assert not en.has_footnotes
        assert en.has_endnotes


class TestOoxmlElementCreationFoundation:
    """F-F1-023: OOXML footnote/endnote body element creation."""

    def test_foundation_create_footnote_element_structure(self):
        """Footnote element has correct OOXML structure."""
        elem = _create_footnote_element(
            1,
            _mk_children("Test content"),
            "FootnoteText",
            "FootnoteReference",
        )

        assert elem.tag == _q("footnote")
        assert elem.get(_q("id")) == "1"

        # Should have one paragraph
        paras = elem.findall(f"{{{WML_NS}}}p")
        assert len(paras) == 1

        # First paragraph should have footnote reference
        p = paras[0]
        pPr = p.find(f"{{{WML_NS}}}pPr")
        assert pPr is not None
        pStyle = pPr.find(f"{{{WML_NS}}}pStyle")
        assert pStyle is not None
        assert pStyle.get(_q("val")) == "FootnoteText"

        # Should contain footnoteRef
        footnote_refs = p.findall(f".//{{{WML_NS}}}footnoteRef")
        assert len(footnote_refs) == 1

        # Should contain content text
        texts = [t.text for t in p.findall(f".//{{{WML_NS}}}t") if t.text]
        assert "Test content" in "".join(texts)

    def test_foundation_create_footnote_element_multiline(self):
        """Multi-paragraph footnote content creates multiple w:p elements."""
        elem = _create_footnote_element(
            5,
            [_mk_children("Line 1")[0], _mk_children("Line 2")[0]],
            "FootnoteText",
            "FootnoteReference",
        )

        paras = elem.findall(f"{{{WML_NS}}}p")
        assert len(paras) == 2

        # Only first paragraph has footnoteRef
        refs_p1 = paras[0].findall(f".//{{{WML_NS}}}footnoteRef")
        assert len(refs_p1) == 1
        refs_p2 = paras[1].findall(f".//{{{WML_NS}}}footnoteRef")
        assert len(refs_p2) == 0

    def test_foundation_create_endnote_element_structure(self):
        """Endnote element has correct OOXML structure."""
        elem = _create_endnote_element(
            3,
            _mk_children("End content"),
            "EndnoteText",
            "EndnoteReference",
        )

        assert elem.tag == _q("endnote")
        assert elem.get(_q("id")) == "3"

        paras = elem.findall(f"{{{WML_NS}}}p")
        assert len(paras) == 1

        p = paras[0]
        refs = p.findall(f".//{{{WML_NS}}}endnoteRef")
        assert len(refs) == 1

        texts = [t.text for t in p.findall(f".//{{{WML_NS}}}t") if t.text]
        assert "End content" in "".join(texts)

    def test_foundation_preserves_leading_trailing_spaces(self):
        """Content with leading/trailing spaces gets xml:space='preserve'."""
        elem = _create_footnote_element(
            1,
            _mk_children("  padded  "),
            "FootnoteText",
            "FootnoteReference",
        )

        t_elements = elem.findall(f".//{{{WML_NS}}}t")
        content_ts = [t for t in t_elements if t.text and t.text.strip()]
        if content_ts:
            t = content_ts[0]
            assert t.get("{http://www.w3.org/XML/1998/namespace}space") == "preserve"


class TestReferenceRunFoundation:
    """F-F1-024: Inline footnote/endnote reference run creation."""

    def test_foundation_footnote_ref_run_has_correct_style(self):
        """Reference run includes the correct rStyle."""
        run = _create_footnote_ref_run(7, "FootnoteReference")

        rPr = run.find(f"{{{WML_NS}}}rPr")
        assert rPr is not None
        rStyle = rPr.find(f"{{{WML_NS}}}rStyle")
        assert rStyle is not None
        assert rStyle.get(_q("val")) == "FootnoteReference"

    def test_foundation_footnote_ref_run_has_correct_id(self):
        """Reference run references the correct footnote ID."""
        run = _create_footnote_ref_run(42, "FootnoteReference")

        ref = run.find(f"{{{WML_NS}}}footnoteReference")
        assert ref is not None
        assert ref.get(_q("id")) == "42"

    def test_foundation_endnote_ref_run_has_correct_style(self):
        """Endnote reference run uses EndnoteReference style."""
        run = _create_endnote_ref_run(3, "EndnoteReference")

        rPr = run.find(f"{{{WML_NS}}}rPr")
        assert rPr is not None, "Missing rPr in endnote ref run"
        rStyle = rPr.find(f"{{{WML_NS}}}rStyle")
        assert rStyle is not None, "Missing rStyle in endnote ref run"
        assert rStyle.get(_q("val")) == "EndnoteReference"

    def test_foundation_endnote_ref_run_has_endnote_reference(self):
        """Endnote reference run contains endnoteReference (not footnoteReference)."""
        run = _create_endnote_ref_run(99, "EndnoteReference")

        fn_refs = run.findall(f"{{{WML_NS}}}footnoteReference")
        assert len(fn_refs) == 0
        en_refs = run.findall(f"{{{WML_NS}}}endnoteReference")
        assert len(en_refs) == 1
        assert en_refs[0].get(_q("id")) == "99"
